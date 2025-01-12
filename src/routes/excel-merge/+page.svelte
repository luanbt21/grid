<script lang="ts">
  import ExcelPreview from "$lib/components/excel-preview.svelte";
  import { Button } from "$lib/components/ui/button/index.js";
  import { Input } from "$lib/components/ui/input";
  import { Label } from "$lib/components/ui/label";
  import { uppercase, downloadExcel } from "$lib/utils";
  import type { ChangeEventHandler } from "svelte/elements";
  import { read, utils } from "xlsx";

  let titleRowN: number | null = $state(null);
  let primaryCol: string | null = $state(null);

  let firstFile = $state("");
  let warning: string | null = $state(null);
  let titleRow: string[] = $state([]);
  let skipRows: {
    data: string[];
    info: { fileName: string; sheetName: string; row: number };
  }[] = $state([]);

  let files: FileList | null = $state(null);
  let isLoading = $state(false);

  const onFilesChanged: ChangeEventHandler<HTMLInputElement> = (event) => {
    if (!event.currentTarget.files) return;
    files = event.currentTarget.files;

    if (files.length <= 1) {
      alert("You should select multiple Excel files");
      return;
    }
    titleRowN = null;
    primaryCol = null;
    firstFile = files[0].name;
    titleRow = [];
    isLoading = false;
    warning = null;
    skipRows = [];
  };

  async function process() {
    if (!titleRowN) {
      alert("You should select a title row number");
      return;
    }
    if (!files?.length) {
      alert("You should select multiple Excel files");
      return;
    }
    warning = null;
    skipRows = [];
    isLoading = true;
    let gotTitle = false;
    let csvRows: any[][] = [];

    for (let file of files) {
      const data = await file.arrayBuffer();
      const workbook = read(data, { type: "array" });

      for (const sheetName of workbook.SheetNames) {
        const worksheet = workbook.Sheets[sheetName];
        const jsonData: any[][] = utils.sheet_to_json(worksheet, {
          header: 1,
          range: 0,
          defval: "",
        });

        jsonData.slice(titleRowN - 1 || 0).forEach((row, idx) => {
          if (gotTitle && idx == 0) {
            for (let i = 0; i < titleRow.length; i++) {
              if (titleRow[i] != row.at(i)) {
                warning = `Excel files have different titles: ${firstFile} vs ${file.name}`;
                return;
              }
            }

            return;
          }

          if (!gotTitle) {
            titleRow = row;
            gotTitle = true;

            // check title row
            let hasBefore = false;
            for (let i = 0; i < row.length; i++) {
              // Determine primary column index based on non-empty title cell
              if (!primaryCol && typeof row[i] === "string" && row[i] !== "") {
                console.log("Primary column: ", row[i], i);
                primaryCol = String.fromCharCode(i + 65);
                hasBefore = true;
                continue;
              }

              if (typeof row[i] === "number") {
                warning = "Your title column have number type";
                continue;
              }

              if (row[i]) {
                hasBefore = true;
                continue;
              }

              if (!row[i] && hasBefore) {
                warning = "Your title column have empty cell";
                continue;
              }

              // ignore when !row[i] && !hasBefore
            }

            if (!hasBefore) {
              warning = "Your title column is empty";
            }
          }

          if (
            gotTitle &&
            !row[primaryCol ? primaryCol.charCodeAt(0) - 65 : 0]
          ) {
            skipRows.push({
              data: row,
              info: {
                fileName: file.name,
                sheetName,
                row: idx + 1,
              },
            });
            console.log("Skipping empty primary column row: ", row);
            return;
          }

          csvRows.push(row);
        });
      }
    }

    skipRows = skipRows;
    isLoading = false;
    downloadExcel(csvRows, "merge", firstFile);
  }
</script>

<div class="max-w-2xl mx-auto p-6">
  <h1 class="text-2xl font-bold mb-4">Excel File Merger</h1>

  <div class="mb-4 grid grid-cols-1 gap-4">
    <div class="grid w-full items-center gap-1.5">
      <Label for="files">Choose Excel files</Label>
      <Input
        id="files"
        type="file"
        multiple
        accept=".xlsx, .xls"
        disabled={isLoading}
        onchange={onFilesChanged}
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
        Primary Column(Will try to use first data column)
      </label>
      <input
        id="primary-column"
        type="text"
        placeholder="A-Z"
        pattern="[A-Z]"
        maxlength="1"
        use:uppercase
        bind:value={primaryCol}
        class="mt-1 p-2 block w-full rounded-md border-gray-700 shadow-sm focus:border-indigo-300 focus:ring focus:ring-indigo-200 focus:ring-opacity-50"
      />
    </div>

    <Button
      disabled={files?.length == 0 || !titleRowN || isLoading}
      onclick={process}
    >
      Merge
    </Button>
  </div>

  {#if isLoading}
    <p class="text-gray-600">Processing file...</p>
  {:else if firstFile}
    {#if warning}
      <p class="text-yellow-500">{warning}</p>
    {/if}
    <details>
      <summary>Details</summary>
      <p>Base file: {firstFile}</p>
      <p>Your table title: {titleRow.join(" | ")}</p>

      <h3>Skipped rows:</h3>
      {#each skipRows as { info }}
        <p>{info.fileName} - {info.sheetName} - {info.row}</p>
      {:else}
        <p>No skipped rows</p>
      {/each}
    </details>

    {#if skipRows.length > 0}
      <div class="overflow-x-auto">
        <label
          for="table"
          class="block text-center font-medium text-gray-700 mb-2"
        >
          Skipped rows:
        </label>
        <table class="min-w-full bg-white border border-gray-300">
          <thead class="bg-gray-200 sticky top-0 z-10">
            <tr>
              {#each titleRow as header}
                <th class="py-2 px-4 border">{header}</th>
              {/each}
            </tr>
          </thead>
          <tbody>
            {#each skipRows as { data, info }}
              <tr
                class="hover:bg-gray-100"
                title={`${info.fileName} - ${info.sheetName} - ${info.row}`}
              >
                {#each data as cell}
                  <td class="py-2 px-4 border">{cell ? cell : ""}</td>
                {:else}
                  <td class="py-2 px-4 border">_empty</td>
                {/each}
              </tr>
            {/each}
          </tbody>
        </table>
      </div>
    {/if}
  {/if}
</div>
