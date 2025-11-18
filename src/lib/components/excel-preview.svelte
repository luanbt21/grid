<script lang="ts">
  import { read, utils } from "xlsx";

  let {
    files,
    file,
    titleRowN,
  }: { files: FileList; file?: File; titleRowN?: number } = $props();
  let data: string[] = $state([]);
  let loading = $state(false);

  const columns = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

  $effect(() => {
    if (files.length) file = files[0];
    if (file) readExcel(file);
  });

  async function readExcel(file: File) {
    loading = true;
    const arrayBuffer = await file.arrayBuffer();
    const workbook = read(arrayBuffer);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    data = utils
      .sheet_to_json(worksheet, { header: 1, range: 0, defval: "" })
      .slice(0, 30) as string[];

    loading = false;
  }
</script>

<div>
  <h2 class="text-xl font-bold mb-4">Preview</h2>
  <description class="text-gray-600">
    The preview only displays the first 30 rows.
  </description>

  <select bind:value={file} class="p-2 border rounded w-full">
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
      <table class="min-w-full border-separate border-spacing-0">
        <thead>
          <tr>
            <th class="py-2 px-4 border sticky top-0 left-0 z-50 bg-primary">
            </th>
            {#each columns.slice(0, data[0].length) as header}
              <th
                class="py-2 px-4 border-y border-r sticky top-0 z-10 bg-primary text-primary-foreground"
              >
                {header}
              </th>
            {/each}
          </tr>
        </thead>

        <tbody>
          {#each data as row, idx}
            <tr>
              <td
                class="py-2 px-2 border-x border-b text-center bg-primary text-primary-foreground sticky left-0"
              >
                {idx + 1}
              </td>
              {#each row as cell}
                <td
                  class="py-2 px-4 border-r border-b"
                  class:bg-secondary={idx === titleRowN}
                  >{cell}
                </td>
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
