<script lang="ts">
  import ExcelPreview from "$lib/components/excel-preview.svelte";
  import { Button } from "$lib/components/ui/button/index.js";
  import { uppercase, downloadExcel } from "$lib/utils";
  import { read, utils } from "xlsx";

  let titleRowN: number | null = null;
  let primaryCol: string | null = null;

  let firstFile = "";
  let warning: string | null = null;
  let titleRow: string[] = [];
  let skipRows: {
    data: string[];
    info: { fileName: string; sheetName: string; row: number };
  }[] = [];

  let files: FileList;
  let isLoading = false;

  function onFilesChanged() {
    if (files.length <= 1) {
      alert("You should select multiple Excel files");
      return;
    }
    titleRowN = null;
    primaryCol = null;
    firstFile = "";
    titleRow = [];
    isLoading = false;
    warning = null;
    skipRows = [];
  }

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
      interface CustomFileReader extends FileReader {
        fileName?: string;
      }
      const reader: CustomFileReader = new FileReader();
      reader.fileName = file.name;

      await new Promise((resolve) => {
        reader.onload = function (e) {
          if (e.target?.result instanceof ArrayBuffer) {
            if (!firstFile) firstFile = reader.fileName || "";
            const data = new Uint8Array(e.target.result);
            const workbook = read(data, { type: "array" });

            workbook.SheetNames.forEach((sheetName) => {
              const worksheet = workbook.Sheets[sheetName];
              const jsonData: any[][] = utils.sheet_to_json(worksheet, {
                header: 1,
              });

              jsonData.forEach((row, idx) => {
                if (!titleRowN) return;
                if (idx < titleRowN - 1) {
                  return;
                }
                if (!gotTitle) {
                  gotTitle = true;
                  titleRow = row;
                  console.log("Title row: ", row);

                  // check title row
                  let hasBefore = false;
                  let ok = true;
                  for (let i = 0; i < row.length; i++) {
                    // Determine primary column index based on non-empty title cell
                    if (
                      !primaryCol &&
                      typeof row[i] === "string" &&
                      row[i] !== ""
                    ) {
                      console.log("Primary column: ", row[i], i);
                      primaryCol = String.fromCharCode(i + 65);
                      hasBefore = true;
                      continue;
                    }

                    if (typeof row[i] === "number") {
                      warning = "Your title column have number type";
                      ok = false;
                      continue;
                    }

                    if (row[i]) {
                      hasBefore = true;
                      continue;
                    }

                    if (!row[i] && hasBefore) {
                      warning = "Your title column have empty cell";
                      ok = false;
                      continue;
                    }

                    // ignore when !row[i] && !hasBefore
                  }

                  if (!hasBefore) {
                    warning = "Your title column is empty";
                    ok = false;
                  }

                  if (!primaryCol) {
                    for (let i = 0; i < row.length; i++) {
                      if (typeof row[i] === "string" && row[i] !== "") {
                        console.log("Primary column: ", row[i], i);
                        primaryCol = String.fromCharCode(i + 65);
                        break;
                      }
                    }
                  }
                }

                if (gotTitle && idx == titleRowN - 1) {
                  for (let i = 0; i < titleRow.length; i++) {
                    if (titleRow[i] != row.at(i)) {
                      warning = `Excel files have different titles: ${firstFile} vs ${file.name}`;
                      return;
                    }
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
            });
          }

          skipRows = skipRows;
          isLoading = false;
          resolve(null);
        };

        reader.readAsArrayBuffer(file);
      });
    }

    downloadExcel(csvRows, "merge", firstFile);
  }
</script>

<div class="max-w-2xl mx-auto p-6">
  <h1 class="text-2xl font-bold mb-4">Excel File Merger</h1>

  <div class="mb-4 grid grid-cols-1 gap-4">
    <div class="mb-4">
      <label
        for="file-upload"
        class="block text-sm font-medium text-gray-700 mb-2"
      >
        Choose Excel files
      </label>
      <input
        id="file-upload"
        type="file"
        accept=".xlsx, .xls"
        multiple
        disabled={isLoading}
        bind:files
        on:change={onFilesChanged}
        class="block w-full text-sm text-gray-500
        file:mr-4 file:py-2 file:px-4
        file:rounded-full file:border-0
        file:text-sm file:font-semibold
        file:bg-violet-50 file:text-violet-700
        hover:file:bg-violet-100
      "
      />
    </div>

    <ExcelPreview {files} />

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
      on:click={process}
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
