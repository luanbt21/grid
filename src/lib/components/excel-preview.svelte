<script lang="ts">
  import { read, utils } from "xlsx";

  interface Props {
    files: FileList;
  }

  let { files }: Props = $props();

  let file: File | undefined = $state();
  let data: any[] = $state([]);

  let loading = $state(false);

  const columns = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";

  $effect(() => {
    if (file) readExcel(file);
  });
  async function readExcel(file: File) {
    loading = true;
    const arrayBuffer = await file.arrayBuffer();
    const workbook = read(arrayBuffer);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    data = utils.sheet_to_json(worksheet, { header: 1 });
    loading = false;
  }
</script>

<div>
  <h2 class="text-xl font-bold mb-4">Preview</h2>

  <select bind:value={file} class="p-2 border border-gray-300 rounded w-full">
    {#if files?.length}
      {#each files as file}
        <option value={file}>{file.name}</option>
      {/each}
    {/if}
  </select>

  {#if loading}
    <p class="text-gray-600">Loading...</p>
  {:else if data.length > 0}
    <div class="overflow-x-auto max-h-96">
      <table class="min-w-full bg-white border border-gray-300">
        <thead class="bg-gray-100 sticky top-0 z-10">
          <tr>
            {#each columns as header}
              <th class="py-2 px-4 border"
                >{header === undefined ? "" : header}</th
              >
            {/each}
          </tr>
        </thead>
        <tbody>
          {#each data as row, idx}
            <tr>
              <td class="py-2 px-2 border text-center bg-gray-100 sticky left-0"
                >{idx + 1}</td
              >
              {#each row as cell}
                <td class="py-2 px-4 border"
                  >{cell === undefined ? "" : cell}</td
                >
              {/each}
            </tr>
          {/each}
        </tbody>
      </table>
    </div>
  {:else}
    <p class="text-gray-600">
      No data to display. Please select an Excel file.
    </p>
  {/if}
</div>
