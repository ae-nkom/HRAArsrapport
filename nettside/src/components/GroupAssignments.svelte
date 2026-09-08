<script>
  import { reportGroups, excludedGroup, unassignedGroup } from '../lib/report-engine.js';
  export let employees = [];
  export let onChange;
  export let disabled = false;
  $: unresolved = employees.filter((employee) => employee.group === unassignedGroup).length;
</script>

{#if employees.length}
  <section class="content-card mt-3" aria-labelledby="group-assignments-title">
    <h2 id="group-assignments-title" class="panel-title">Avklar rapportgrupper</h2>
    <p class="section-note mt-2">Velg gruppe for stillinger som ikke kan plasseres automatisk. Underdirektører med personalansvar hører til direktørgruppen, og underdirektører uten personalansvar til fagsjefgruppen. Valgene gjelder bare dette uttrekket og lagres ikke mellom økter.</p>
    <p class="section-note mt-2" aria-live="polite">{unresolved ? `${unresolved} ansatte mangler gruppevalg.` : 'Alle gruppevalg er avklart.'}</p>
    <div class="ledger-table-wrap mt-3">
      <table class="ledger-table">
        <thead><tr><th scope="col">Ansatt</th><th scope="col">Stilling</th><th scope="col">Rapportgruppe</th></tr></thead>
        <tbody>
          {#each employees as employee (employee.personKey)}
            <tr>
              <td>{employee.name}</td><td>{employee.position}</td>
              <td>
                <select class="report-manual-input" aria-label={`Rapportgruppe for ${employee.name}`} value={employee.group} {disabled} on:change={(event) => onChange(employee.personKey, event.currentTarget.value)}>
                  <option value={unassignedGroup}>Velg rapportgruppe</option>
                  {#each reportGroups as group}<option value={group}>{group}</option>{/each}
                  <option value={excludedGroup}>{excludedGroup}</option>
                </select>
              </td>
            </tr>
          {/each}
        </tbody>
      </table>
    </div>
    <p class="section-note mt-2">Ansatte uten gruppevalg vises i «Uavklart rapportgruppe». «Utelatt fra rapporten» fjerner den ansatte fra bemannings- og lønnstallene, og antallet utelatte oppgis i rapportens merknader. Et ferdig avklart uttrekk kan også ha kolonnen «Rapportgruppe», eller «Personalansvar» med ja/nei for underdirektører.</p>
  </section>
{/if}
