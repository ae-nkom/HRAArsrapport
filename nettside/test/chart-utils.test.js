import test from "node:test";
import assert from "node:assert/strict";
import { buildNiceScale, buildChartScale, scaleFraction, formatChartTick, formatChartValue } from "../src/components/charts/chart-utils.js";

test("diagramverdier bruker norsk tall-, prosent- og desimalformat", () => {
  assert.equal(formatChartValue(5667.752, "#,##0.0"), "5 667,8");
  assert.equal(formatChartValue(0.437, "pct1"), "43,7 %");
  assert.equal(formatChartValue(143, "num0"), "143");
});

test('diagramaksen viser korreksjoner under null og fyller aksen ved små prosentverdier', () => {
  const scale = buildChartScale([-200, 1200]);
  assert.ok(scale.minimum <= -200);
  assert.ok(scale.maximum >= 1200);
  assert.ok(scale.ticks.includes(0));
  assert.equal(scaleFraction(0.8, 0, 0.8), 1);
  assert.equal(formatChartValue(null, 'num0'), '—');
});

test("diagramaksen dekker maksimum med lesbare intervaller", () => {
  assert.deepEqual(buildNiceScale(5667), {
    maximum: 6000,
    ticks: [0, 2000, 4000, 6000]
  });
  assert.equal(formatChartTick(2_000_000, "#,##0"), "2 mill.");
  assert.deepEqual(buildNiceScale(10, 4, true).ticks, [0, 5, 10]);
});
