<script lang="ts">
  import { read, utils } from "xlsx";

  let titleRowN = 1; // Assuming first row contains titles
  let primaryCol: string | null = null; // Default primary column index
  let warning: string | null = null;
  let titleRow: string[] = [];

  let files: FileList;
  let isLoading = false;
  let url = "";

  async function handleFileUpload() {
    if (files.length <= 1) {
      alert("You should select multiple Excel files");
      return;
    }

    isLoading = true;
    let gotTitle = false;
    let csvContent = ""; // This will hold our CSV data
    let csvRows: any[][] = []; // Store each row as an array of strings

    for (let file of files) {
      const reader = new FileReader();

      reader.onload = function (e) {
        if (e.target?.result instanceof ArrayBuffer) {
          const data = new Uint8Array(e.target.result);
          const workbook = read(data, { type: "array" });

          // Iterate over all sheets
          workbook.SheetNames.forEach((sheetName) => {
            const worksheet = workbook.Sheets[sheetName];
            const jsonData: any[][] = utils.sheet_to_json(worksheet, {
              header: 1,
            });

            let skip = !gotTitle ? titleRowN - 1 : titleRowN;
            console.log("Skip: ", skip);

            // Process each row
            jsonData.forEach((row, idx) => {
              if (idx < skip) {
                return;
              }
              if (!gotTitle) {
                gotTitle = true;
                titleRow = row;
                console.log("Title row: ", row);

                // // check title row
                // let hasBefore = false;
                // let ok = true;
                // for (let i = 0; i < row.length; i++) {
                //   if (typeof row[i] === "number") {
                //     warning = "Your title column have number type";
                //     ok = false;
                //     continue;
                //   }

                //   if (row[i]) {
                //     hasBefore = true;
                //     continue;
                //   }

                //   if (!row[i] && hasBefore) {
                //     warning = "Your title column have empty cell";
                //     ok = false;
                //     continue;
                //   }

                //   if (!row[i]) {
                //     console.log("Primary column: ", row[i], i);
                //     primaryCol = String.fromCharCode(i + 65);
                //     continue;
                //   }
                // }

                // if (!hasBefore) {
                //   warning = "Your title column is empty";
                //   ok = false;
                // }

                // Determine primary column index based on non-empty title cell
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

              if (
                gotTitle &&
                !row[primaryCol ? primaryCol.charCodeAt(0) - 65 : 0]
              ) {
                console.log("Skipping empty primary column row: ", row);
                return;
              }

              csvRows.push(row);
            });
          });
        }

        csvContent = csvRows.map((row) => row.join(",")).join("\n");
        const blob = new Blob([csvContent], { type: "text/csv" });
        url = URL.createObjectURL(blob);
        isLoading = false;
      };

      reader.readAsArrayBuffer(file);
    }
  }
</script>

<div class="max-w-2xl mx-auto p-6">
  <h1 class="text-2xl font-bold mb-4">Excel File Merger</h1>

  <div class="mb-4 grid grid-cols-1 gap-4">
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
        maxlength="1"
        bind:value={primaryCol}
        class="mt-1 p-2 block w-full rounded-md border-gray-700 shadow-sm focus:border-indigo-300 focus:ring focus:ring-indigo-200 focus:ring-opacity-50"
      />
    </div>
  </div>

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
      on:change={handleFileUpload}
      class="block w-full text-sm text-gray-500
        file:mr-4 file:py-2 file:px-4
        file:rounded-full file:border-0
        file:text-sm file:font-semibold
        file:bg-violet-50 file:text-violet-700
        hover:file:bg-violet-100
      "
    />
  </div>

  {#if isLoading}
    <p class="text-gray-600">Processing file...</p>
  {:else if url}
    <a href={url} download="merged_output.csv">Download</a>
    <details open>Your table title: {titleRow.join(", ")}</details>
    {#if warning}
      <p class="text-yellow-500">{warning}</p>
    {/if}
  {:else}
    <p class="text-gray-600">No file selected</p>
  {/if}
</div>
