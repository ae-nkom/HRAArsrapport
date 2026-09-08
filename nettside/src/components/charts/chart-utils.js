export function chartDecimalPlaces(format = "") {
  if (format === "pct1") return 1;
  if (format === "num0") return 0;
  const decimalPart = String(format).split(".")[1] || "";
  return decimalPart.length;
}

export function formatChartValue(value, format = "") {
  if (value === null || value === undefined || value === '') return '—';
  const number = Number(value);
  if (!Number.isFinite(number)) return "—";

  const decimals = chartDecimalPlaces(format);
  const scaled = format === "pct1" ? number * 100 : number;
  const formatted = new Intl.NumberFormat("nb-NO", {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals
  }).format(scaled);

  return format === "pct1" ? `${formatted} %` : formatted;
}

export function formatChartTick(value, format = "") {
  if (value === null || value === undefined || value === '') return '—';
  const number = Number(value);
  if (!Number.isFinite(number)) return "—";
  if (format === "pct1") return formatChartValue(number, format);

  const decimals = chartDecimalPlaces(format);
  if (Math.abs(number) >= 1_000_000) {
    return new Intl.NumberFormat("nb-NO", {
      notation: "compact",
      maximumFractionDigits: 1
    }).format(number);
  }

  return new Intl.NumberFormat("nb-NO", {
    minimumFractionDigits: 0,
    maximumFractionDigits: decimals
  }).format(number);
}

export function chartNumber(value) {
  return value !== null && value !== undefined && value !== '' && Number.isFinite(Number(value));
}

export function buildChartScale(values, integersOnly = false) {
  const numbers = values.filter(chartNumber).map(Number);
  const lowest = Math.min(0, ...numbers);
  const highest = Math.max(0, ...numbers);
  if (lowest >= 0) return { minimum: 0, ...buildNiceScale(highest, 4, integersOnly) };
  const positiveScale = buildNiceScale(highest - lowest, 4, integersOnly);
  const step = positiveScale.ticks[1] - positiveScale.ticks[0];
  const minimum = Math.floor(lowest / step) * step;
  const maximum = Math.ceil(highest / step) * step;
  return { minimum, maximum, ticks: Array.from({ length: Math.round((maximum - minimum) / step) + 1 }, (_, index) => minimum + index * step) };
}

export function scaleFraction(value, minimum, maximum) {
  return (Number(value) - minimum) / (maximum - minimum || 1);
}

export function buildNiceScale(maximum, targetIntervals = 4, integersOnly = false) {
  const safeMaximum = Math.max(0, Number(maximum) || 0);
  if (!safeMaximum) return { maximum: 1, ticks: [0, 0.25, 0.5, 0.75, 1] };

  const roughStep = safeMaximum / Math.max(1, targetIntervals);
  const magnitude = 10 ** Math.floor(Math.log10(roughStep));
  const normalized = roughStep / magnitude;
  const niceNormalized = integersOnly
    ? (normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10)
    : (normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 2.5 ? 2.5 : normalized <= 5 ? 5 : 10);
  const step = Math.max(integersOnly ? 1 : 0, niceNormalized * magnitude);
  const maximumValue = Math.ceil(safeMaximum / step) * step;
  const intervals = Math.round(maximumValue / step);

  return {
    maximum: maximumValue,
    ticks: Array.from({ length: intervals + 1 }, (_, index) => index * step)
  };
}
