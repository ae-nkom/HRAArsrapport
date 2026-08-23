import test from "node:test";
import assert from "node:assert/strict";
import { buildNiceScale, formatChartTick, formatChartValue } from "../src/components/charts/chart-utils.js";

test("diagramverdier bruker norsk tall-, prosent- og desimalformat", () => {
  assert.equal(formatChartValue(5667.752, "#,##0.0"), "5 667,8");
  assert.equal(formatChartValue(0.437, "pct1"), "43,7 %");
  assert.equal(formatChartValue(143, "num0"), "143");
});

test("diagramaksen dekker maksimum med lesbare intervaller", () => {
  assert.deepEqual(buildNiceScale(5667), {
    maximum: 6000,
    ticks: [0, 2000, 4000, 6000]
  });
  assert.equal(formatChartTick(2_000_000, "#,##0"), "2 mill.");
  assert.deepEqual(buildNiceScale(10, 4, true).ticks, [0, 5, 10]);
});
