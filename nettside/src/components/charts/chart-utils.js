export function chartDecimalPlaces(format = "") {
  if (format === "pct1") return 1;
  if (format === "num0") return 0;
  const decimalPart = String(format).split(".")[1] || "";
  return decimalPart.length;
}

export function formatChartValue(value, format = "") {
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
