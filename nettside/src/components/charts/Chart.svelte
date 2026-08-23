<script>
  import { buildNiceScale, formatChartTick, formatChartValue } from "./chart-utils.js";

  export let kind = "bar";
  export let data = [];
  export let x = "x";
  export let y = "y";
  export let series = "";
  export let type = "grouped";
  export let seriesOrder = [];
  export let seriesColors = {};
  export let colorPalette = [];
  export let swapXY = false;
  export let name = "";
  export let min = "min";
  export let intervalBottom = "q1";
  export let midpoint = "median";
  export let intervalTop = "q3";
  export let max = "max";
  export let xAxisTitle = "";
  export let yAxisTitle = "";
  export let value = "";
  export let colorScale = ["#efefef", "#345fed"];
  export let xFmt = "";
  export let yFmt = "";
  export let valueFmt = "";

  const width = 900;
  const height = 390;
  const margin = { top: 28, right: 112, bottom: 82, left: 180 };
  const plotWidth = width - margin.left - margin.right;
  const plotHeight = height - margin.top - margin.bottom;
  const fallbackColors = ["#3028aa", "#f15b0a", "#345fed", "#00a166", "#db2481", "#6085ff"];

  $: safeData = Array.isArray(data) ? data.filter(Boolean) : [];
  $: categories = [...new Set(safeData.map((row) => String(row[x] ?? row[name] ?? "—")))];
  $: seriesNames = series
    ? (seriesOrder.length
        ? seriesOrder.filter((item) => safeData.some((row) => String(row[series] ?? "") === String(item)))
        : [...new Set(safeData.map((row) => String(row[series] ?? "")))])
    : [""];
  $: numericValues = safeData.map((row) => Number(row[y] ?? row[value] ?? 0)).filter(Number.isFinite);
  $: rawMaximum = Math.max(0, ...numericValues, ...safeData.flatMap((row) => [Number(row[max] || 0), Number(row[intervalTop] || 0)]));
  $: valueScale = type === "stacked100" ? { maximum: 1, ticks: [0, 0.25, 0.5, 0.75, 1] } : buildNiceScale(rawMaximum, 4, yFmt === "num0");
  $: maximum = valueScale.maximum;
  $: axisTicks = valueScale.ticks;
  $: xNumericMaximum = Math.max(0, ...safeData.map((row) => Number(row[x] || 0)));
  $: xScale = buildNiceScale(xNumericMaximum, 4, xFmt === "num0");
  $: heatRows = [...new Set(safeData.map((row) => String(row[y] ?? "—")))];
  $: barItems = buildBarItems();
  $: lineSeries = buildLineSeries();
  $: horizontalAxisTitle = swapXY ? yAxisTitle : xAxisTitle;
  $: verticalAxisTitle = swapXY ? xAxisTitle : yAxisTitle;
  $: numericFormat = yFmt || valueFmt;
  $: rotateCategoryLabels = !swapXY && kind === "bar" && categories.length > 4;

  function colorFor(seriesName, index) {
    return seriesColors?.[seriesName] || colorPalette?.[index] || fallbackColors[index % fallbackColors.length];
  }

  function shortLabel(label, limit = 24) {
    const text = String(label);
    return text.length > limit ? `${text.slice(0, limit - 2)}…` : text;
  }

  function lineCategoryLabel(label) {
    const text = String(label);
    const monthAndYear = text.match(/^([a-zæøå]{3})\s+20\d{2}$/i);
    return monthAndYear ? monthAndYear[1] : shortLabel(text, 8);
  }

  function tickX(amount, scaleMaximum = maximum) {
    return margin.left + Number(amount || 0) / Math.max(1, scaleMaximum) * plotWidth;
  }

  function tickY(amount) {
    return margin.top + plotHeight - Number(amount || 0) / Math.max(1, maximum) * plotHeight;
  }

  function buildBarItems() {
    const items = [];
    const categoryBand = (swapXY ? plotHeight : plotWidth) / Math.max(1, categories.length);
    const groupedBand = categoryBand * 0.7 / Math.max(1, seriesNames.length);

    categories.forEach((category, categoryIndex) => {
      const rows = safeData.filter((row) => String(row[x] ?? "—") === category);
      const stackedTotal = rows.reduce((sum, row) => sum + Number(row[y] || 0), 0) || 1;
      let stackOffset = 0;

      seriesNames.forEach((seriesName, seriesIndex) => {
        const row = rows.find((item) => !series || String(item[series] ?? "") === String(seriesName));
        if (!row) return;
        const rawValue = Number(row?.[y] || 0);
        const displayValue = type === "stacked100" ? rawValue / stackedTotal : rawValue;
        const length = Math.max(0, displayValue / Math.max(1, maximum)) * (swapXY ? plotWidth : plotHeight);
        const groupedOffset = (categoryBand - groupedBand * seriesNames.length) / 2 + groupedBand * seriesIndex;
        const stackLength = Math.max(0, stackOffset / Math.max(1, maximum)) * (swapXY ? plotWidth : plotHeight);
        const isStacked = type === "stacked100";

        items.push({
          category,
          value: rawValue,
          displayValue,
          seriesName,
          color: colorFor(seriesName, seriesIndex || categoryIndex),
          x: swapXY ? margin.left + (isStacked ? stackLength : 0) : margin.left + categoryIndex * categoryBand + (isStacked ? categoryBand * 0.15 : groupedOffset),
          y: swapXY ? margin.top + categoryIndex * categoryBand + (isStacked ? categoryBand * 0.15 : groupedOffset) : margin.top + plotHeight - length - (isStacked ? stackLength : 0),
          width: swapXY ? length : (isStacked ? categoryBand * 0.7 : groupedBand),
          height: swapXY ? (isStacked ? categoryBand * 0.7 : groupedBand) : length
        });
        if (isStacked) stackOffset += displayValue;
      });
    });

    return items;
  }

  function buildLineSeries() {
    return seriesNames.map((seriesName, seriesIndex) => {
      const rows = series ? safeData.filter((row) => String(row[series] ?? "") === String(seriesName)) : safeData;
      const points = categories.map((category, index) => {
        const row = rows.find((item) => String(item[x] ?? "—") === category);
        const amount = Number(row?.[y] || 0);
        return {
          category,
          amount,
          x: margin.left + (index + 0.5) * plotWidth / Math.max(1, categories.length),
          y: tickY(amount)
        };
      });
      return { points, color: colorFor(seriesName, seriesIndex), seriesName };
    });
  }

  function heatOpacity(row) {
    const amount = Number(row[value] || 0);
    const heatMax = Math.max(1, ...safeData.map((item) => Number(item[value] || 0)));
    return 0.18 + 0.82 * amount / heatMax;
  }
</script>

<div class="local-chart">
  {#if safeData.length}
    {#if series && seriesNames.filter(Boolean).length}
      <div class="chart-legend" aria-label="Tegnforklaring">
        {#each seriesNames.filter(Boolean) as seriesName, index}
          <span class="legend-item"><span class="legend-swatch" style={`--legend-color:${colorFor(seriesName, index)}`}></span>{seriesName}</span>
        {/each}
      </div>
    {/if}
    <svg viewBox={`0 0 ${width} ${height}`} role="img" aria-label={`${yAxisTitle || "Verdi"} etter ${xAxisTitle || "kategori"}`}>
      {#if kind === "bar" || kind === "line" || kind === "box"}
        {#each axisTicks as tick}
          {#if swapXY || kind === "box"}
            <line x1={tickX(tick)} y1={margin.top} x2={tickX(tick)} y2={margin.top + plotHeight} class="grid-line" />
            <text x={tickX(tick)} y={margin.top + plotHeight + 20} text-anchor="middle" class="tick-label">{formatChartTick(tick, numericFormat)}</text>
          {:else}
            <line x1={margin.left} y1={tickY(tick)} x2={margin.left + plotWidth} y2={tickY(tick)} class="grid-line" />
            <text x={margin.left - 10} y={tickY(tick)} text-anchor="end" dominant-baseline="middle" class="tick-label">{formatChartTick(tick, numericFormat)}</text>
          {/if}
        {/each}
      {/if}

      {#if kind === "scatter"}
        {#each xScale.ticks as tick}
          <line x1={tickX(tick, xScale.maximum)} y1={margin.top} x2={tickX(tick, xScale.maximum)} y2={margin.top + plotHeight} class="grid-line" />
          <text x={tickX(tick, xScale.maximum)} y={margin.top + plotHeight + 20} text-anchor="middle" class="tick-label">{formatChartTick(tick, xFmt)}</text>
        {/each}
        {#each axisTicks as tick}
          <line x1={margin.left} y1={tickY(tick)} x2={margin.left + plotWidth} y2={tickY(tick)} class="grid-line" />
          <text x={margin.left - 10} y={tickY(tick)} text-anchor="end" dominant-baseline="middle" class="tick-label">{formatChartTick(tick, yFmt)}</text>
        {/each}
      {/if}

      <line x1={margin.left} y1={margin.top + plotHeight} x2={margin.left + plotWidth} y2={margin.top + plotHeight} class="axis" />
      <line x1={margin.left} y1={margin.top} x2={margin.left} y2={margin.top + plotHeight} class="axis" />

      {#if kind === "bar"}
        {#each barItems as item}
          <rect x={item.x} y={item.y} width={item.width} height={item.height} rx="3" fill={item.color}>
            <title>{item.category}{item.seriesName ? ` · ${item.seriesName}` : ""}: {formatChartValue(type === "stacked100" ? item.displayValue : item.value, yFmt)}</title>
          </rect>
          {#if swapXY}
            <text x={item.x + item.width + 7} y={item.y + item.height / 2} dominant-baseline="middle" class="value-label">{formatChartValue(type === "stacked100" ? item.displayValue : item.value, yFmt)}</text>
          {:else}
            <text x={item.x + item.width / 2} y={item.y - 7} text-anchor="middle" class="value-label">{formatChartValue(type === "stacked100" ? item.displayValue : item.value, yFmt)}</text>
          {/if}
        {/each}
        {#each categories as category, index}
          {#if swapXY}
            <text x={margin.left - 10} y={margin.top + (index + 0.5) * plotHeight / categories.length} text-anchor="end" dominant-baseline="middle">{shortLabel(category)}</text>
          {:else if rotateCategoryLabels}
            <text x={margin.left + (index + 0.5) * plotWidth / categories.length} y={margin.top + plotHeight + 18} text-anchor="end" transform={`rotate(-30 ${margin.left + (index + 0.5) * plotWidth / categories.length} ${margin.top + plotHeight + 18})`}>{shortLabel(category, 15)}</text>
          {:else}
            <text x={margin.left + (index + 0.5) * plotWidth / categories.length} y={margin.top + plotHeight + 20} text-anchor="middle">{shortLabel(category)}</text>
          {/if}
        {/each}
      {:else if kind === "line"}
        {#each lineSeries as item}
          <polyline points={item.points.map((point) => `${point.x},${point.y}`).join(" ")} fill="none" stroke={item.color} stroke-width="4" stroke-linejoin="round"><title>{item.seriesName || yAxisTitle}</title></polyline>
          {#each item.points as point, index}
            <circle cx={point.x} cy={point.y} r="5" fill="#fff" stroke={item.color} stroke-width="3"><title>{point.category}: {formatChartValue(point.amount, yFmt)}</title></circle>
            <text x={point.x} y={point.y + (index % 2 ? 19 : -11)} text-anchor="middle" class="value-label line-value">{formatChartValue(point.amount, yFmt)}</text>
            <text x={point.x} y={margin.top + plotHeight + 20} text-anchor="middle">{lineCategoryLabel(point.category)}</text>
          {/each}
        {/each}
      {:else if kind === "scatter"}
        {#each safeData as row, index}
          {@const cx = tickX(Number(row[x] || 0), xScale.maximum)}
          {@const cy = tickY(Number(row[y] || 0))}
          <circle {cx} {cy} r="7" fill={colorFor(String(row[series] || ""), index)} opacity="0.86"><title>{row[series]}: {formatChartValue(row[x], xFmt)} · {formatChartValue(row[y], yFmt)}</title></circle>
          <text x={cx + 10} y={cy - 8} class="value-label">{shortLabel(row[series] || "", 17)}</text>
          <text x={cx + 10} y={cy + 8} class="point-detail">{formatChartValue(row[x], xFmt)} / {formatChartValue(row[y], yFmt)}</text>
        {/each}
      {:else if kind === "heatmap"}
        {#each safeData as row}
          {@const column = categories.indexOf(String(row[x] ?? "—"))}
          {@const heatRow = heatRows.indexOf(String(row[y] ?? "—"))}
          {@const opacity = heatOpacity(row)}
          {@const cellX = margin.left + column * plotWidth / categories.length}
          {@const cellY = margin.top + heatRow * plotHeight / heatRows.length}
          <rect x={cellX} y={cellY} width={plotWidth / categories.length - 3} height={plotHeight / heatRows.length - 3} rx="4" fill={colorScale[1]} {opacity}><title>{row[x]} · {row[y]}: {formatChartValue(row[value], valueFmt)}</title></rect>
          <text x={cellX + (plotWidth / categories.length - 3) / 2} y={cellY + (plotHeight / heatRows.length - 3) / 2} text-anchor="middle" dominant-baseline="middle" class:heat-value-light={opacity > 0.58} class="heat-value">{formatChartValue(row[value], valueFmt)}</text>
        {/each}
        {#each heatRows as rowName, index}<text x={margin.left - 10} y={margin.top + (index + 0.5) * plotHeight / heatRows.length} text-anchor="end" dominant-baseline="middle">{shortLabel(rowName)}</text>{/each}
        {#each categories as category, index}<text x={margin.left + (index + 0.5) * plotWidth / categories.length} y={margin.top + plotHeight + 20} text-anchor="middle">{shortLabel(category, 17)}</text>{/each}
      {:else if kind === "box"}
        {#each safeData as row, index}
          {@const cy = margin.top + (index + 0.5) * plotHeight / safeData.length}
          {@const scale = (number) => tickX(number)}
          <line x1={scale(row[min])} y1={cy} x2={scale(row[max])} y2={cy} stroke={colorFor("", index)} stroke-width="3" />
          <rect x={scale(row[intervalBottom])} y={cy - 13} width={Math.max(2, scale(row[intervalTop]) - scale(row[intervalBottom]))} height="26" fill={colorFor("", index)} opacity="0.32" stroke={colorFor("", index)} />
          <line x1={scale(row[midpoint])} y1={cy - 13} x2={scale(row[midpoint])} y2={cy + 13} stroke={colorFor("", index)} stroke-width="4"><title>{row[name]}: median {formatChartValue(row[midpoint], yFmt)}</title></line>
          <text x={margin.left - 10} y={cy} text-anchor="end" dominant-baseline="middle">{shortLabel(row[name])}</text>
          <text x={scale(row[midpoint])} y={cy - 20} text-anchor="middle" class="value-label">Median {formatChartValue(row[midpoint], yFmt)}</text>
        {/each}
      {/if}

      {#if verticalAxisTitle}<text x="18" y={margin.top + plotHeight / 2} transform={`rotate(-90 18 ${margin.top + plotHeight / 2})`} text-anchor="middle" class="axis-title">{verticalAxisTitle}</text>{/if}
      {#if horizontalAxisTitle}<text x={margin.left + plotWidth / 2} y={height - 10} text-anchor="middle" class="axis-title">{horizontalAxisTitle}</text>{/if}
    </svg>
  {:else}
    <p>Ingen data å vise.</p>
  {/if}
</div>

<style>
  .local-chart { width: 100%; min-height: 16rem; display: grid; place-items: center; color: #526176; }
  .chart-legend { width: 100%; display: flex; flex-wrap: wrap; justify-content: flex-end; gap: 8px 18px; padding: 0 12px 2px; color: #334155; font-size: 11px; }
  .legend-item { display: inline-flex; align-items: center; gap: 6px; }
  .legend-swatch { width: 10px; height: 10px; border-radius: 2px; background: var(--legend-color); }
  svg { width: 100%; height: auto; max-height: 30rem; overflow: visible; }
  text { fill: #526176; font: 12px Lato, Inter, system-ui, sans-serif; }
  .axis { stroke: #94a3b8; stroke-width: 1.2; }
  .grid-line { stroke: #e2e8f0; stroke-width: 1; }
  .tick-label { fill: #64748b; font-size: 11px; }
  .axis-title { fill: #334155; font-weight: 700; }
  .value-label { fill: #172554; font-size: 11px; font-weight: 700; paint-order: stroke; stroke: #fff; stroke-width: 3px; stroke-linejoin: round; }
  .line-value { font-size: 10px; }
  .point-detail { fill: #64748b; font-size: 9px; paint-order: stroke; stroke: #fff; stroke-width: 3px; }
  .heat-value { fill: #172554; font-size: 11px; font-weight: 700; }
  .heat-value-light { fill: #fff; }
  @media (max-width: 720px) { text { font-size: 10px; } .local-chart { overflow-x: auto; justify-content: start; } .chart-legend { justify-content: flex-start; } svg { min-width: 48rem; } }
</style>
