<script lang="ts">
  import ExcelPreview from "$lib/components/excel-preview.svelte";
  import { Button } from "$lib/components/ui/button/index.js";
  import { uppercase, downloadExcel } from "$lib/utils";
  import { read, utils } from "xlsx";

  let titleRowN = $state(1);
  let keyCol: string | null = $state(null);
  let valueCol: string | null = $state(null);
  let separator = $state(", ");

  let map: Map<string, Set<string>> = $state(new Map());

  let files: FileList | undefined = $state();
  let isLoading = $state(false);
  let warning = "";

  function onFilesChanged() {
    keyCol = null;
    valueCol = null;
    map = new Map();
    isLoading = false;
    warning = "";
  }

  async function process() {
    if (!files) {
      return;
    }

    isLoading = true;
    let csvRows: any[][] = [];

    for (let file of files) {
      const data = await file.arrayBuffer();
      const workbook = read(data, { type: "array" });

      workbook.SheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const jsonData: string[][] = utils.sheet_to_json(worksheet, {
          header: 1,
          range: 0,
          defval: "",
        });

        jsonData.forEach((row, idx) => {
          if (idx < titleRowN - 1) {
            return;
          }

          if (!keyCol || !valueCol) {
            return;
          }

          const keyColNum = keyCol.charCodeAt(0) - 65;
          const valueColNum = valueCol.charCodeAt(0) - 65;

          if (!row[keyColNum]) {
            return;
          }
          const entry = map.get(row[keyColNum]);
          if (entry) {
            entry.add(row[valueColNum]);
          } else {
            map.set(row[keyColNum], new Set([row[valueColNum]]));
          }
        });

        jsonData.forEach((row) => {
          if (!keyCol || !valueCol) {
            return;
          }

          const keyColNum = keyCol.charCodeAt(0) - 65;
          const valueColNum = valueCol.charCodeAt(0) - 65;

          const entry = map.get(row[keyColNum]);
          if (entry) {
            const corresponding = entry
              .values()
              .filter((value) => value !== row[valueColNum])
              .toArray();
            // insert corresponding next to value column
            row.splice(valueColNum + 1, 0, corresponding.join(separator));
          }
          csvRows.push(row);
        });
      });

      downloadExcel(csvRows, "corresponding", file.name);
    }

    isLoading = false;
  }
</script>

<div class="max-w-2xl mx-auto p-6">
  <h1 class="text-2xl font-bold mb-4">Excel Corresponding</h1>

  <div class="mb-4 grid grid-cols-1 gap-4">
    <div class="mb-4">
      <label
        for="file-upload"
        class="block text-sm font-medium text-gray-700 mb-2"
      >
        Choose Excel file
      </label>
      <input
        id="file-upload"
        type="file"
        accept=".xlsx, .xls"
        disabled={isLoading}
        bind:files
        onchange={onFilesChanged}
        class="block w-full text-sm text-gray-500
        file:mr-4 file:py-2 file:px-4
        file:rounded-full file:border-0
        file:text-sm file:font-semibold
        file:bg-violet-50 file:text-violet-700
        hover:file:bg-violet-100
      "
      />
    </div>

    {#if files?.length}
      <ExcelPreview {files} />
    {/if}

    <div>
      <label
        for="title-row"
        class="block text-sm font-medium text-gray-700 mb-2"
      >
        Title Row
      </label>
      <input
        id="title-row"
        type="number"
        min="1"
        bind:value={titleRowN}
        class="mt-1 p-2 block w-full rounded-md border-gray-700 shadow-sm focus:border-indigo-300 focus:ring focus:ring-indigo-200 focus:ring-opacity-50"
      />
    </div>

    <div>
      <label
        for="primary-column"
        class="block text-sm font-medium text-gray-700 mb-2"
      >
        Primary Column
      </label>
      <input
        id="primary-column"
        type="text"
        placeholder="A-Z"
        pattern="[A-Z]"
        maxlength="1"
        use:uppercase
        bind:value={keyCol}
        class="mt-1 p-2 block w-full rounded-md border-gray-700 shadow-sm focus:border-indigo-300 focus:ring focus:ring-indigo-200 focus:ring-opacity-50"
      />
    </div>

    <div>
      <label
        for="primary-column"
        class="block text-sm font-medium text-gray-700 mb-2"
      >
        Value Column
      </label>
      <input
        id="primary-column"
        type="text"
        placeholder="A-Z"
        pattern="[A-Z]"
        maxlength="1"
        use:uppercase
        bind:value={valueCol}
        class="mt-1 p-2 block w-full rounded-md border-gray-700 shadow-sm focus:border-indigo-300 focus:ring focus:ring-indigo-200 focus:ring-opacity-50"
      />
    </div>

    <div>
      <label
        for="primary-column"
        class="block text-sm font-medium text-gray-700 mb-2"
      >
        Separator
      </label>
      <input
        id="primary-column"
        type="text"
        bind:value={separator}
        class="mt-1 p-2 block w-full rounded-md border-gray-700 shadow-sm focus:border-indigo-300 focus:ring focus:ring-indigo-200 focus:ring-opacity-50"
      />
    </div>

    <Button
      disabled={!files?.length || !keyCol || !valueCol || isLoading}
      onclick={process}
    >
      Process
    </Button>
  </div>

  {#if isLoading}
    <p class="text-gray-600">Processing file...</p>
  {:else if map.size > 0}
    <p>number of key: {map.size}</p>
  {/if}
</div>
