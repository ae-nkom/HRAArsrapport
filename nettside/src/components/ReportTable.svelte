<script>
  export let columns = [];
  export let rows = [];
  export let caption = '';
</script>

<!-- The scrollable table must be reachable by keyboard on narrow screens. -->
<!-- svelte-ignore a11y_no_noninteractive_tabindex -->
<div class="ledger-table-wrap report-table-wrap" tabindex="0" role="region" aria-label={caption}>
  <table class="ledger-table report-table" class:report-table-short={columns.length <= 3}>
    <caption>{caption}</caption>
    <thead><tr>{#each columns as column}<th scope="col" class:text-right={column.numeric}>{column.label}</th>{/each}</tr></thead>
    <tbody>
      {#each rows as row}
        <tr>{#each columns as column, index}
          {#if index === 0}<th scope="row">{row[column.key] ?? '—'}</th>
          {:else}<td class:text-right={column.numeric} class:tabular-nums={column.numeric}>{row[column.key] ?? '—'}</td>{/if}
        {/each}</tr>
      {/each}
    </tbody>
  </table>
</div>

{#if columns.length > 3}<p class="table-scroll-hint">Rull sidelengs for å se alle kolonnene.</p>{/if}
