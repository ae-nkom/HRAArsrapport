export const numberFormatter = new Intl.NumberFormat("nb-NO", {
  maximumFractionDigits: 1
});

export const integerFormatter = new Intl.NumberFormat("nb-NO", {
  maximumFractionDigits: 0
});

export const percentFormatter = new Intl.NumberFormat("nb-NO", {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2
});

export const percentValueFormatter = new Intl.NumberFormat("nb-NO", {
  minimumFractionDigits: 0,
  maximumFractionDigits: 2
});

export function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

// SAP may export a numeric cell or Norwegian text with spaces and decimal comma.
// Missing or malformed values must not quietly become zero.
export function parseNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value !== "string" || !value.trim()) return null;
  let text = value.trim().replace(/[\s\u00a0\u202f]/g, "");
  if (text.includes(",")) {
    if (!/^[+-]?(?:\d{1,3}(?:\.\d{3})+|\d+),\d+$/.test(text)) return null;
    text = text.replaceAll(".", "").replace(",", ".");
  }
  if (!/^[+-]?\d+(?:\.\d+)?$/.test(text)) return null;
  const number = Number(text);
  return Number.isFinite(number) ? number : null;
}

function isNumber(value) {
  return value !== null && value !== undefined && value !== "" && Number.isFinite(Number(value));
}

export function formatNumber(value) {
  return isNumber(value) ? numberFormatter.format(Number(value)) : "—";
}

export function formatPercent(value) {
  return isNumber(value) ? `${percentFormatter.format(Number(value))} %` : "—";
}

export function formatPercentValue(value) {
  return isNumber(value) ? percentValueFormatter.format(Number(value)) : "—";
}

export function formatInteger(value) {
  return isNumber(value) ? integerFormatter.format(Number(value)) : "—";
}

export function formatCurrency(value) {
  return isNumber(value) ? `${integerFormatter.format(Number(value))} kr` : "—";
}

export function normalizePersonName(value) {
  return String(value ?? "").normalize("NFKC").toLocaleLowerCase("nb-NO")
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .trim();
}

export function mean(values) {
  if (!values.length || values.some((value) => !Number.isFinite(value))) return null;
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

export function quantileSorted(values, quantile) {
  if (!values.length) return 0;
  const index = (values.length - 1) * quantile;
  const lower = Math.floor(index);
  const upper = Math.ceil(index);
  if (lower === upper) return values[lower];
  const weight = index - lower;
  return values[lower] * (1 - weight) + values[upper] * weight;
}
