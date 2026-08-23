<script>
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

  const width = 900;
  const height = 340;
  const margin = { top: 24, right: 28, bottom: 76, left: 170 };
  const plotWidth = width - margin.left - margin.right;
  const plotHeight = height - margin.top - margin.bottom;
  const fallbackColors = ["#3028aa", "#f15b0a", "#345fed", "#00a166", "#db2481", "#6085ff"];

  $: safeData = Array.isArray(data) ? data.filter(Boolean) : [];
  $: categories = [...new Set(safeData.map((row) => String(row[x] ?? row[name] ?? "—")))];
  $: seriesNames = series
    ? (seriesOrder.length ? seriesOrder.filter((item) => safeData.some((row) => row[series] === item)) : [...new Set(safeData.map((row) => String(row[series] ?? "")))])
    : [""];
  $: numericValues = safeData.map((row) => Number(row[y] ?? row[value] ?? 0)).filter(Number.isFinite);
  $: maximum = Math.max(1, ...numericValues, ...safeData.flatMap((row) => [Number(row[max] || 0), Number(row[intervalTop] || 0)]));
  $: heatRows = [...new Set(safeData.map((row) => String(row[y] ?? "—")))];
  $: barItems = buildBarItems();
  $: lineItems = buildLineItems();

  function colorFor(seriesName, index) {
    return seriesColors?.[seriesName] || colorPalette?.[index] || fallbackColors[index % fallbackColors.length];
  }

  function shortLabel(label) {
    const text = String(label);
    return text.length > 24 ? `${text.slice(0, 22)}…` : text;
  }

  function buildBarItems() {
    const items = [];
    const categoryBand = (swapXY ? plotHeight : plotWidth) / Math.max(1, categories.length);
    const groupedBand = categoryBand * 0.74 / Math.max(1, seriesNames.length);
    categories.forEach((category, categoryIndex) => {
      const rows = safeData.filter((row) => String(row[x] ?? "—") === category);
      const stackedTotal = rows.reduce((sum, row) => sum + Number(row[y] || 0), 0) || 1;
      let stackOffset = 0;
      seriesNames.forEach((seriesName, seriesIndex) => {
        const row = rows.find((item) => !series || String(item[series] ?? "") === String(seriesName));
        const rawValue = Number(row?.[y] || 0);
        const displayValue = type === "stacked100" ? rawValue / stackedTotal : rawValue;
        const domainMax = type === "stacked100" ? 1 : maximum;
        const length = Math.max(0, displayValue / domainMax) * (swapXY ? plotWidth : plotHeight);
        const groupedOffset = (categoryBand - groupedBand * seriesNames.length) / 2 + groupedBand * seriesIndex;
        const stackLength = Math.max(0, stackOffset / domainMax) * (swapXY ? plotWidth : plotHeight);
        const isStacked = type === "stacked100";
        items.push({
          category,
          value: rawValue,
          seriesName,
          color: colorFor(seriesName, seriesIndex || categoryIndex),
          x: swapXY ? margin.left + (isStacked ? stackLength : 0) : margin.left + categoryIndex * categoryBand + (isStacked ? categoryBand * 0.13 : groupedOffset),
          y: swapXY ? margin.top + categoryIndex * categoryBand + (isStacked ? categoryBand * 0.13 : groupedOffset) : margin.top + plotHeight - length - (isStacked ? stackLength : 0),
          width: swapXY ? length : (isStacked ? categoryBand * 0.74 : groupedBand),
          height: swapXY ? (isStacked ? categoryBand * 0.74 : groupedBand) : length
        });
        if (isStacked) stackOffset += displayValue;
      });
    });
    return items;
  }

  function buildLineItems() {
    return seriesNames.map((seriesName, seriesIndex) => {
      const rows = series ? safeData.filter((row) => String(row[series] ?? "") === String(seriesName)) : safeData;
      const points = categories.map((category, index) => {
        const row = rows.find((item) => String(item[x] ?? "—") === category);
        const amount = Number(row?.[y] || 0);
        return `${margin.left + (index + 0.5) * plotWidth / Math.max(1, categories.length)},${margin.top + plotHeight - amount / maximum * plotHeight}`;
      }).join(" ");
      return { points, color: colorFor(seriesName, seriesIndex), seriesName };
    });
  }

  function heatOpacity(row) {
    const amount = Number(row[value] || 0);
    const heatMax = Math.max(1, ...safeData.map((item) => Number(item[value] || 0)));
    return 0.15 + 0.85 * amount / heatMax;
  }
</script>

<div class="local-chart">
  {#if safeData.length}
    <svg viewBox={`0 0 ${width} ${height}`} role="img" aria-label={`${yAxisTitle || "Verdi"} etter ${xAxisTitle || "kategori"}`}>
      <line x1={margin.left} y1={margin.top + plotHeight} x2={margin.left + plotWidth} y2={margin.top + plotHeight} class="axis" />
      {#if kind === "bar"}
        {#each barItems as item}
          <rect x={item.x} y={item.y} width={item.width} height={item.height} rx="3" fill={item.color}>
            <title>{item.category}{item.seriesName ? ` · ${item.seriesName}` : ""}: {item.value.toLocaleString("nb-NO")}</title>
          </rect>
        {/each}
        {#each categories as category, index}
          {#if swapXY}
            <text x={margin.left - 10} y={margin.top + (index + 0.5) * plotHeight / categories.length} text-anchor="end" dominant-baseline="middle">{shortLabel(category)}</text>
          {:else}
            <text x={margin.left + (index + 0.5) * plotWidth / categories.length} y={margin.top + plotHeight + 18} text-anchor="middle">{shortLabel(category)}</text>
          {/if}
        {/each}
      {:else if kind === "line"}
        {#each lineItems as item}
          <polyline points={item.points} fill="none" stroke={item.color} stroke-width="4" stroke-linejoin="round"><title>{item.seriesName || yAxisTitle}</title></polyline>
        {/each}
      {:else if kind === "scatter"}
        {#each safeData as row, index}
          <circle cx={margin.left + Number(row[x] || 0) / Math.max(1, ...safeData.map((item) => Number(item[x] || 0))) * plotWidth} cy={margin.top + plotHeight - Number(row[y] || 0) / maximum * plotHeight} r="6" fill={colorFor(String(row[series] || ""), index)} opacity="0.82"><title>{Number(row[x] || 0).toLocaleString("nb-NO")} · {Number(row[y] || 0).toLocaleString("nb-NO")}</title></circle>
        {/each}
      {:else if kind === "heatmap"}
        {#each safeData as row}
          {@const column = categories.indexOf(String(row[x] ?? "—"))}
          {@const heatRow = heatRows.indexOf(String(row[y] ?? "—"))}
          <rect x={margin.left + column * plotWidth / categories.length} y={margin.top + heatRow * plotHeight / heatRows.length} width={plotWidth / categories.length - 2} height={plotHeight / heatRows.length - 2} rx="4" fill={colorScale[1]} opacity={heatOpacity(row)}><title>{row[x]} · {row[y]}: {Number(row[value] || 0).toLocaleString("nb-NO")}</title></rect>
        {/each}
        {#each heatRows as rowName, index}<text x={margin.left - 10} y={margin.top + (index + 0.5) * plotHeight / heatRows.length} text-anchor="end" dominant-baseline="middle">{shortLabel(rowName)}</text>{/each}
      {:else if kind === "box"}
        {#each safeData as row, index}
          {@const cy = margin.top + (index + 0.5) * plotHeight / safeData.length}
          {@const scale = (number) => margin.left + Number(number || 0) / maximum * plotWidth}
          <line x1={scale(row[min])} y1={cy} x2={scale(row[max])} y2={cy} stroke={colorFor("", index)} stroke-width="3" />
          <rect x={scale(row[intervalBottom])} y={cy - 13} width={Math.max(2, scale(row[intervalTop]) - scale(row[intervalBottom]))} height="26" fill={colorFor("", index)} opacity="0.32" stroke={colorFor("", index)} />
          <line x1={scale(row[midpoint])} y1={cy - 13} x2={scale(row[midpoint])} y2={cy + 13} stroke={colorFor("", index)} stroke-width="4"><title>{row[name]}: median {Number(row[midpoint] || 0).toLocaleString("nb-NO")}</title></line>
          <text x={margin.left - 10} y={cy} text-anchor="end" dominant-baseline="middle">{shortLabel(row[name])}</text>
        {/each}
      {/if}
      {#if yAxisTitle}<text x="18" y={height / 2} transform={`rotate(-90 18 ${height / 2})`} text-anchor="middle" class="axis-title">{yAxisTitle}</text>{/if}
      {#if xAxisTitle}<text x={margin.left + plotWidth / 2} y={height - 10} text-anchor="middle" class="axis-title">{xAxisTitle}</text>{/if}
    </svg>
  {:else}
    <p>Ingen data å vise.</p>
  {/if}
</div>

<style>
  .local-chart { width: 100%; min-height: 16rem; display: grid; place-items: center; color: #526176; }
  svg { width: 100%; height: auto; max-height: 26rem; overflow: visible; }
  text { fill: #526176; font: 12px Lato, Inter, system-ui, sans-serif; }
  .axis { stroke: #cbd5e1; stroke-width: 1; }
  .axis-title { fill: #334155; font-weight: 700; }
  @media (max-width: 720px) { text { font-size: 10px; } .local-chart { overflow-x: auto; justify-content: start; } svg { min-width: 42rem; } }
</style>
