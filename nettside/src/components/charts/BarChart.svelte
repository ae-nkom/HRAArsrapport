<script>
  import { buildChartScale, scaleFraction, chartNumber, formatChartTick, formatChartValue } from './chart-utils.js';
  export let data = [];
  export let x = 'x';
  export let y = 'y';
  export let series = '';
  export let seriesOrder = [];
  export let seriesColors = {};
  export let colorPalette = [];
  export let yFmt = '';
  export let yAxisTitle = '';
  export let xAxisTitle = '';
  const palette = ['#345793', '#009c83', '#7750a3', '#c45d26'];

  $: rows = Array.isArray(data) ? data.filter(row => row && chartNumber(row[y])) : [];
  $: categories = [...new Set(rows.map(row => String(row[x] ?? '—')))];
  $: seriesNames = series ? (seriesOrder.length ? seriesOrder : [...new Set(rows.map(row => String(row[series] ?? '')))]) : [''];
  $: scale = buildChartScale(rows.map(row => row[y]), yFmt === 'num0');
  $: middleTick = (scale.minimum + scale.maximum) / 2;
  $: ticks = yFmt === 'num0' && !Number.isInteger(middleTick) ? [scale.minimum, scale.maximum] : [scale.minimum, middleTick, scale.maximum];
  $: groups = categories.map((label, categoryIndex) => ({label, values: seriesNames.map((seriesName, index) => {
    const row = rows.find(row => String(row[x] ?? '—') === label && (!series || String(row[series] ?? '') === seriesName));
    if (!row) return null;
    const amount = Number(row[y]);
    const zero = scaleFraction(0, scale.minimum, scale.maximum) * 1000;
    const end = scaleFraction(amount, scale.minimum, scale.maximum) * 1000;
    return { amount, seriesName, left: Math.min(zero, end), width: Math.abs(end - zero), color: seriesColors[seriesName] || colorPalette[series ? index : categoryIndex] || palette[series ? index % palette.length : 0] };
  }).filter(Boolean)}));
</script>

<div class="local-chart bar-chart" role="group" aria-label={`${yAxisTitle || 'Verdi'} etter ${xAxisTitle || 'kategori'}`}>
  {#if rows.length}
    <div class="chart-heading">
      <span class="chart-unit">{yAxisTitle}</span>
      {#if series}<div class="chart-legend" aria-label="Tegnforklaring">
        {#each seriesNames as label, index}<span><i style={`background:${seriesColors[label] || colorPalette[index] || palette[index % palette.length]}`}></i>{label}</span>{/each}
      </div>{/if}
    </div>
    <div class="scale-row" aria-hidden="true"><div></div><div class="scale-labels">{#each ticks as tick}<span>{formatChartTick(tick, yFmt)}</span>{/each}</div><div></div></div>
    {#each groups as group}
      <div class="bar-group">
        <div class="bar-category">{group.label}</div>
        <div class="bar-values">
          {#each group.values as item}
            <div class="bar-row">
              <svg viewBox="0 0 1000 24" preserveAspectRatio="none" role="img" aria-label={`${group.label}${item.seriesName ? ', ' + item.seriesName : ''}: ${formatChartValue(item.amount, yFmt)}`}>
                <line x1={scaleFraction(0, scale.minimum, scale.maximum) * 1000} x2={scaleFraction(0, scale.minimum, scale.maximum) * 1000} y1="0" y2="24" stroke="#b8c6d3" vector-effect="non-scaling-stroke" />
                <rect x={item.left} y="4" width={item.width} height="16" rx="2" fill={item.color} />
              </svg>
              <span class="bar-value">{formatChartValue(item.amount, yFmt)}</span>
            </div>
          {/each}
        </div>
      </div>
    {/each}
  {:else}<p class="empty-chart">Ingen tall å vise for valgt grunnlag.</p>{/if}
</div>

<style>
  .bar-chart { width: 100%; min-width: 0; color: #314355; font-size: 13px; }
  .chart-heading { display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px; margin-bottom: 18px; }
  .chart-unit { color: #617184; font-size: 12px; }
  .chart-legend { display: flex; flex-wrap: wrap; gap: 16px; }
  .chart-legend span { display: inline-flex; gap: 7px; align-items: center; }
  .chart-legend i { display: block; width: 10px; height: 10px; border-radius: 2px; }
  .scale-row { display: grid; grid-template-columns: minmax(150px, 230px) minmax(0, 1fr) 100px; gap: 18px; margin-bottom: 5px; }
  .scale-labels { display: flex; justify-content: space-between; font-size: 11px; color: #617184; }
  .bar-group { display: grid; grid-template-columns: minmax(150px, 230px) minmax(0, 1fr); gap: 18px; align-items: center; padding: 12px 0; border-top: 1px solid #e9eef3; }
  .bar-category { line-height: 1.5; overflow-wrap: anywhere; font-weight: 500; }
  .bar-values { min-width: 0; display: grid; gap: 4px; }
  .bar-row { min-width: 0; display: grid; grid-template-columns: minmax(0, 1fr) 100px; align-items: center; gap: 18px; }
  svg { width: 100%; height: 24px; overflow: hidden; background: repeating-linear-gradient(to right, transparent 0, transparent calc(25% - 1px), #eef2f6 calc(25% - 1px), #eef2f6 25%); }
  .bar-value { text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; font-size: 13px; color: #182d44; }
  .empty-chart { padding: 24px 0; color: #617184; }
  @media (max-width: 700px) {
    .scale-row { grid-template-columns: minmax(0, 1fr) 88px; gap: 12px; }
    .scale-row > div:first-child { display: none; }
    .bar-group { grid-template-columns: minmax(0, 1fr); gap: 8px; }
    .bar-row { grid-template-columns: minmax(0, 1fr) 88px; gap: 12px; }
  }
</style>
