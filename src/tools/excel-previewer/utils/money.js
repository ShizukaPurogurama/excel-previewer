import {
  getHeaderSignature,
  isEmptyCell,
  isNumericLike,
  normalizeCell,
  normalizeHeaderLabel,
} from './columns.js';
import { MONEY_MAPPING_STORAGE_KEY, SUPPORTED_CURRENCY_CODES } from '../../../constants.js';

const currencyFormatterCache = new Map();

export function normalizeMoneyMappingIndex(value, totalColumns) {
  if (!Number.isInteger(value) || value < 0) return null;
  if (Number.isInteger(totalColumns) && totalColumns > 0 && value >= totalColumns) return null;
  return value;
}

export function normalizeMoneyMapping(mapping, totalColumns) {
  const safeTotal = Number.isInteger(totalColumns) && totalColumns > 0 ? totalColumns : 0;
  return {
    senderCurrency: normalizeMoneyMappingIndex(mapping && mapping.senderCurrency, safeTotal),
    senderAmount: normalizeMoneyMappingIndex(mapping && mapping.senderAmount, safeTotal),
    receiverCurrency: normalizeMoneyMappingIndex(mapping && mapping.receiverCurrency, safeTotal),
    receiverAmount: normalizeMoneyMappingIndex(mapping && mapping.receiverAmount, safeTotal),
    detectedAmountColumns: Array.isArray(mapping && mapping.detectedAmountColumns)
      ? mapping.detectedAmountColumns.filter(
          (index) =>
            Number.isInteger(index) && index >= 0 && (safeTotal === 0 || index < safeTotal)
        )
      : [],
    detectedCurrencyColumns: Array.isArray(mapping && mapping.detectedCurrencyColumns)
      ? mapping.detectedCurrencyColumns.filter(
          (index) =>
            Number.isInteger(index) && index >= 0 && (safeTotal === 0 || index < safeTotal)
        )
      : [],
    source: mapping && mapping.source === 'manual' ? 'manual' : 'auto',
  };
}

export function hasAnyMoneyMapping(mapping) {
  return Boolean(
    mapping &&
      (Number.isInteger(mapping.senderCurrency) ||
        Number.isInteger(mapping.senderAmount) ||
        Number.isInteger(mapping.receiverCurrency) ||
        Number.isInteger(mapping.receiverAmount))
  );
}

export function getSupportedCurrencyCode(value) {
  const normalized = normalizeCell(value).toUpperCase();
  const matches = normalized.match(/\b(?:USD|KHR)\b/);
  if (matches && SUPPORTED_CURRENCY_CODES[matches[0]]) return matches[0];
  if (normalized.indexOf('$') !== -1) return 'USD';
  if (normalized.indexOf('៛') !== -1) return 'KHR';
  if (!matches || !SUPPORTED_CURRENCY_CODES[matches[0]]) return '';
  return matches[0];
}

export function hasCurrencySymbol(value) {
  const text = normalizeCell(value);
  return /[$៛]/.test(text);
}

export function parseCurrencyAmount(value) {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  const text = normalizeCell(value);
  if (!text) return null;

  const normalized = text
    .replace(/^\((.*)\)$/, '-$1')
    .replace(/\b(?:USD|KHR)\b/gi, '')
    .replace(/[$៛,\s]/g, '');

  if (!/^[-+]?(?:\d+\.?\d*|\.\d+)$/.test(normalized)) return null;
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

export function resolveCurrencyFractionDigits(value) {
  const text = normalizeCell(value).replace(/,/g, '');
  const fractionMatch = text.match(/\.(\d+)/);
  return fractionMatch ? Math.min(fractionMatch[1].length, 6) : 0;
}

export function getCurrencyFormatter(currencyCode, fractionDigits) {
  const safeFractionDigits =
    Number.isInteger(fractionDigits) && fractionDigits >= 0 ? fractionDigits : 0;
  const cacheKey = currencyCode + ':' + safeFractionDigits;

  if (!currencyFormatterCache.has(cacheKey)) {
    currencyFormatterCache.set(
      cacheKey,
      new Intl.NumberFormat(undefined, {
        currency: currencyCode,
        currencyDisplay: 'narrowSymbol',
        maximumFractionDigits: safeFractionDigits,
        minimumFractionDigits: safeFractionDigits,
        style: 'currency',
      })
    );
  }

  return currencyFormatterCache.get(cacheKey);
}

export function isAmountColumnHeader(header) {
  return /\b(amount|amt|value|total|net|gross|balance|debit|credit|fee|charge|payment|payout|settlement|principal|price|cost)\b/.test(
    normalizeHeaderLabel(header)
  );
}

export function isCurrencyColumnHeader(header) {
  return /\b(currency|ccy|curr|fx)\b/.test(normalizeHeaderLabel(header));
}

export function isAccountLikeHeader(headerText) {
  return /\b(account|acct|iban|wallet|number|no|accountno|accountnumber|beneficiaryaccount|receiveraccount|senderaccount)\b/.test(
    headerText
  );
}

export function getTransferSide(headerText) {
  if (/\b(sender|from|source|debit|payer|origin)\b/.test(headerText)) return 'sender';
  if (/\b(receiver|to|destination|credit|beneficiary|recipient|payee)\b/.test(headerText)) return 'receiver';
  return '';
}

function scoreMoneyAmountHeader(headerText) {
  let score = 0;
  if (
    /\b(amount|amt|value|total|net|gross|balance|debit|credit|fee|charge|payment|payout|settlement|principal|price|cost)\b/.test(
      headerText
    )
  ) {
    score += 4;
  }
  if (/\b(sender|receiver|recipient|beneficiary|payer|payee|from|to)\b/.test(headerText)) {
    score += 1;
  }
  return score;
}

function scoreMoneyCurrencyHeader(headerText) {
  let score = 0;
  if (/\b(currency|curr|ccy|fx)\b/.test(headerText)) score += 5;
  if (/\b(usd|khr)\b/.test(headerText)) score += 3;
  return score;
}

function pickBestMoneyCandidate(candidates, side, type, excludedIndexes) {
  if (!Array.isArray(candidates) || !candidates.length) return null;

  const exclusions = new Set(Array.isArray(excludedIndexes) ? excludedIndexes : []);
  const scored = candidates
    .map((candidate) => {
      const sideBonus = candidate.side === side ? 2.5 : candidate.side ? 0.5 : 0;
      return {
        index: candidate.index,
        score:
          (type === 'currency' ? candidate.currencyScore : candidate.amountScore) + sideBonus,
      };
    })
    .filter((candidate) => !exclusions.has(candidate.index))
    .sort((left, right) => right.score - left.score || left.index - right.index);

  return scored.length ? scored[0].index : null;
}

export function detectMoneyMapping(headers, rows) {
  const activeHeaders = Array.isArray(headers) ? headers : [];
  const sampleRows = Array.isArray(rows) ? rows.slice(0, 40) : [];

  const stats = activeHeaders.map((header, index) => {
    const values = sampleRows
      .map((row) => (row ? row[index] : null))
      .filter((value) => !isEmptyCell(value));

    const nonEmptyCount = values.length;
    const numericLikeCount = values.filter(
      (value) => parseCurrencyAmount(value) !== null || isNumericLike(value)
    ).length;
    const currencyCodeCount = values.filter((value) =>
      Boolean(getSupportedCurrencyCode(value))
    ).length;
    const currencySymbolCount = values.filter((value) => hasCurrencySymbol(value)).length;
    const headerText = normalizeHeaderLabel(header);
    const side = getTransferSide(headerText);
    const accountLikeHeader = isAccountLikeHeader(headerText);

    return {
      index,
      amountScore:
        scoreMoneyAmountHeader(headerText) +
        (nonEmptyCount ? numericLikeCount / nonEmptyCount : 0) * 6 +
        currencySymbolCount * 2 -
        (accountLikeHeader ? 6 : 0),
      currencyScore:
        scoreMoneyCurrencyHeader(headerText) +
        (nonEmptyCount ? currencyCodeCount / nonEmptyCount : 0) * 8 +
        currencySymbolCount * 3 -
        (accountLikeHeader ? 8 : 0),
      side,
    };
  });

  const amountCandidates = stats
    .filter((item) => item.amountScore >= 2.2)
    .sort((left, right) => right.amountScore - left.amountScore || left.index - right.index);
  const currencyCandidates = stats
    .filter((item) => item.currencyScore >= 2.2)
    .sort((left, right) => right.currencyScore - left.currencyScore || left.index - right.index);

  const detectedAmountColumns = amountCandidates.map((item) => item.index);
  const detectedCurrencyColumns = currencyCandidates.map((item) => item.index);

  const senderAmount = pickBestMoneyCandidate(amountCandidates, 'sender', 'amount', []);
  const receiverAmount =
    amountCandidates.length > 1
      ? pickBestMoneyCandidate(
          amountCandidates,
          'receiver',
          'amount',
          senderAmount !== null ? [senderAmount] : []
        )
      : null;
  const senderCurrency = pickBestMoneyCandidate(currencyCandidates, 'sender', 'currency', []);
  const receiverCurrency = pickBestMoneyCandidate(
    currencyCandidates,
    'receiver',
    'currency',
    currencyCandidates.length > 1 && senderCurrency !== null ? [senderCurrency] : []
  );

  return normalizeMoneyMapping(
    {
      senderCurrency,
      senderAmount,
      receiverCurrency,
      receiverAmount,
      detectedAmountColumns,
      detectedCurrencyColumns,
      source: 'auto',
    },
    activeHeaders.length
  );
}

function mergeMoneyMappings(baseMapping, overrideMapping) {
  const base = normalizeMoneyMapping(baseMapping, 0);
  const override = normalizeMoneyMapping(overrideMapping, 0);

  return {
    senderCurrency:
      override.senderCurrency !== null ? override.senderCurrency : base.senderCurrency,
    senderAmount: override.senderAmount !== null ? override.senderAmount : base.senderAmount,
    receiverCurrency:
      override.receiverCurrency !== null ? override.receiverCurrency : base.receiverCurrency,
    receiverAmount:
      override.receiverAmount !== null ? override.receiverAmount : base.receiverAmount,
    detectedAmountColumns: base.detectedAmountColumns,
    detectedCurrencyColumns: base.detectedCurrencyColumns,
    source: override.source === 'manual' ? 'manual' : base.source || 'auto',
  };
}

export function loadMoneyMappingFromStorage(headers) {
  try {
    const signature = getHeaderSignature(headers);
    if (!signature) return null;

    const raw = window.localStorage.getItem(MONEY_MAPPING_STORAGE_KEY);
    if (!raw) return null;

    const data = JSON.parse(raw);
    return normalizeMoneyMapping(data && data[signature], headers.length);
  } catch (error) {
    return null;
  }
}

export function saveMoneyMappingToStorage(headers, mapping) {
  try {
    const signature = getHeaderSignature(headers);
    if (!signature) return;

    const raw = window.localStorage.getItem(MONEY_MAPPING_STORAGE_KEY);
    const data = raw ? JSON.parse(raw) : {};
    data[signature] = {
      senderCurrency: Number.isInteger(mapping.senderCurrency) ? mapping.senderCurrency : null,
      senderAmount: Number.isInteger(mapping.senderAmount) ? mapping.senderAmount : null,
      receiverCurrency: Number.isInteger(mapping.receiverCurrency) ? mapping.receiverCurrency : null,
      receiverAmount: Number.isInteger(mapping.receiverAmount) ? mapping.receiverAmount : null,
    };
    window.localStorage.setItem(MONEY_MAPPING_STORAGE_KEY, JSON.stringify(data));
  } catch (error) {
    // ignore
  }
}

export function removeMoneyMappingFromStorage(headers) {
  try {
    const signature = getHeaderSignature(headers);
    if (!signature) return;

    const raw = window.localStorage.getItem(MONEY_MAPPING_STORAGE_KEY);
    if (!raw) return;

    const data = JSON.parse(raw);
    delete data[signature];
    window.localStorage.setItem(MONEY_MAPPING_STORAGE_KEY, JSON.stringify(data));
  } catch (error) {
    // ignore
  }
}

export function resolveMoneyMappingForHeaders(headers, rows, moneyMappings) {
  const activeHeaders = Array.isArray(headers) ? headers.slice() : [];
  const detected = detectMoneyMapping(activeHeaders, rows);
  const signature = getHeaderSignature(activeHeaders);
  const stored = signature && moneyMappings ? moneyMappings.get(signature) : null;
  const saved = stored || loadMoneyMappingFromStorage(activeHeaders);

  if (saved) {
    return {
      mapping: mergeMoneyMappings(detected, saved),
      source: 'manual',
    };
  }

  return {
    mapping: detected,
    source: 'auto',
  };
}

export function isMoneyAmountColumn(index, header, moneyMapping) {
  const mapping = normalizeMoneyMapping(moneyMapping, 0);
  return (
    isAmountColumnHeader(header) ||
    index === mapping.senderAmount ||
    index === mapping.receiverAmount ||
    (Array.isArray(mapping.detectedAmountColumns) &&
      mapping.detectedAmountColumns.indexOf(index) !== -1)
  );
}

export function isMoneyCurrencyColumn(index, header, moneyMapping) {
  const mapping = normalizeMoneyMapping(moneyMapping, 0);
  return (
    isCurrencyColumnHeader(header) ||
    index === mapping.senderCurrency ||
    index === mapping.receiverCurrency ||
    (Array.isArray(mapping.detectedCurrencyColumns) &&
      mapping.detectedCurrencyColumns.indexOf(index) !== -1)
  );
}

export function buildAmountCurrencyColumnLookup(headers, moneyMapping) {
  const safeHeaders = Array.isArray(headers) ? headers.slice() : [];
  const mapping = normalizeMoneyMapping(moneyMapping, safeHeaders.length);
  const currencyIndexes = safeHeaders.reduce((indexes, header, index) => {
    if (isMoneyCurrencyColumn(index, header, mapping)) indexes.push(index);
    return indexes;
  }, []);
  const lookup = safeHeaders.map(() => -1);
  const assignedAmountIndexes = new Set();

  function assignAmount(amountIndex, currencyIndex) {
    if (!Number.isInteger(amountIndex) || amountIndex < 0 || amountIndex >= lookup.length) return;
    lookup[amountIndex] = Number.isInteger(currencyIndex) ? currencyIndex : -1;
    assignedAmountIndexes.add(amountIndex);
  }

  function nearestCurrencyIndex(amountIndex) {
    if (!currencyIndexes.length) return -1;

    let bestIndex = currencyIndexes[0];
    let bestDistance = Number.POSITIVE_INFINITY;

    currencyIndexes.forEach((currencyIndex) => {
      const distance = Math.abs(currencyIndex - amountIndex);
      if (distance < bestDistance) {
        bestDistance = distance;
        bestIndex = currencyIndex;
      }
    });

    return bestIndex;
  }

  [
    [mapping.senderAmount, mapping.senderCurrency],
    [mapping.receiverAmount, mapping.receiverCurrency],
  ].forEach((pair) => {
    const amountIndex = pair[0];
    const currencyIndex = pair[1];
    if (Number.isInteger(amountIndex) && amountIndex >= 0) {
      assignAmount(
        amountIndex,
        Number.isInteger(currencyIndex) ? currencyIndex : nearestCurrencyIndex(amountIndex)
      );
    }
  });

  safeHeaders.forEach((header, index) => {
    if (assignedAmountIndexes.has(index) || !isMoneyAmountColumn(index, header, mapping)) return;
    assignAmount(index, nearestCurrencyIndex(index));
  });

  return lookup;
}

function resolvePairedCurrencyCode(options) {
  if (
    !options ||
    !Array.isArray(options.amountCurrencyColumns) ||
    !Array.isArray(options.activeColumns) ||
    !Array.isArray(options.row)
  ) {
    return '';
  }

  const currencyHeaderIndex = options.amountCurrencyColumns[options.activeHeaderIndex];
  if (!Number.isInteger(currencyHeaderIndex) || currencyHeaderIndex < 0) return '';

  const sourceColumnIndex = options.activeColumns[currencyHeaderIndex];
  if (!Number.isInteger(sourceColumnIndex) || sourceColumnIndex < 0) return '';

  return getSupportedCurrencyCode(options.row[sourceColumnIndex]);
}

function resolveAmountCurrencyCode(value, options) {
  const pairedCurrencyCode = resolvePairedCurrencyCode(options);
  if (pairedCurrencyCode) return pairedCurrencyCode;

  const headerCurrencyCode = getSupportedCurrencyCode(options.headers[options.activeHeaderIndex]);
  if (headerCurrencyCode) return headerCurrencyCode;

  return getSupportedCurrencyCode(value);
}

export function formatAmountCellValue(value, options) {
  if (!options || !Array.isArray(options.headers)) return null;

  const header = options.headers[options.activeHeaderIndex] || '';
  if (!isMoneyAmountColumn(options.activeHeaderIndex, header, options.moneyMapping)) return null;

  const currencyCode = resolveAmountCurrencyCode(value, options);
  if (!currencyCode) return null;

  const numericValue = parseCurrencyAmount(value);
  if (numericValue === null) return null;

  const fractionDigits = resolveCurrencyFractionDigits(value);
  return getCurrencyFormatter(currencyCode, fractionDigits).format(numericValue);
}
