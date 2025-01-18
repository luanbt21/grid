<script lang="ts">
import ExcelPreview from "$lib/components/excel-preview.svelte";
import { Button } from "$lib/components/ui/button/index.js";
import { Input } from "$lib/components/ui/input";
import { Label } from "$lib/components/ui/label";
import { downloadExcel } from "$lib/utils";
import type { ChangeEventHandler } from "svelte/elements";
import { read, utils } from "xlsx";

let skipRow = $state(0);
let keyCol: string | null = $state(null);
let valueCol: string | null = $state(null);
let separator = $state(", ");

$effect(() => {
	if (keyCol) keyCol = keyCol.toUpperCase();
});
$effect(() => {
	if (valueCol) valueCol = valueCol.toUpperCase();
});

let map: Map<string, Set<string>> = $state(new Map());

let files: FileList | null = $state(null);
let isLoading = $state(false);
let warning = "";

const onFilesChanged: ChangeEventHandler<HTMLInputElement> = (event) => {
	files = event.currentTarget.files;
	keyCol = null;
	valueCol = null;
	map = new Map();
	isLoading = false;
	warning = "";
};

async function process() {
	if (!files?.length || !keyCol || !valueCol) {
		return;
	}

	isLoading = true;
	let csvRows: unknown[][] = [];
	const keyColNum = keyCol.charCodeAt(0) - 65;
	const valueColNum = valueCol.charCodeAt(0) - 65;

	for (let file of files) {
		const data = await file.arrayBuffer();
		const workbook = read(data);

		for (const sheetName of workbook.SheetNames) {
			const worksheet = workbook.Sheets[sheetName];
			const jsonData: string[][] = utils.sheet_to_json(worksheet, {
				header: 1,
				range: 0,
				defval: "",
			});

			// setup data
			for (const row of jsonData.slice(skipRow || 0)) {
				if (!row[keyColNum]) {
					return;
				}
				const entry = map.get(row[keyColNum]);
				if (entry) {
					entry.add(row[valueColNum]);
				} else {
					map.set(row[keyColNum], new Set([row[valueColNum]]));
				}
			}

			// insert corresponding
			for (const row of jsonData.slice(skipRow || 0)) {
				const entry = map.get(row[keyColNum]);
				console.log(111, row[keyColNum], entry);

				if (entry) {
					const corresponding = entry
						.values()
						.filter((value) => value !== row[valueColNum])
						.toArray();
					// insert corresponding next to value column
					row.splice(valueColNum + 1, 0, corresponding.join(separator));
				}
				csvRows.push(row);
			}

			downloadExcel(csvRows, "corresponding", `${sheetName}-${file.name}`);
		}
	}

	isLoading = false;
}
</script>

<svelte:head>
  <title>Excel Corresponding</title>
</svelte:head>
<div class="max-w-2xl mx-auto p-6">
  <h1 class="text-2xl font-bold mb-4">Excel Corresponding</h1>

  <div class="flex flex-col gap-4">
    <div class="grid w-full items-center gap-1.5">
      <Label for="files">Choose Excel file</Label>
      <Input
        id="files"
        type="file"
        accept=".xlsx, .xls"
        disabled={isLoading}
        onchange={onFilesChanged}
      />
    </div>

    {#if files?.length}
      <ExcelPreview {files} />
    {/if}

    <div class="grid w-full items-center gap-1.5">
      <Label for="skip-row">Skip Row</Label>
      <Input
        id="skip-row"
        type="number"
        min={1}
        bind:value={skipRow}
        disabled={isLoading}
      />
    </div>

    <div class="grid w-full items-center gap-1.5">
      <Label for="primary-column">Primary Column</Label>
      <Input
        id="primary-column"
        placeholder="A-Z"
        pattern="[A-Z]"
        maxlength={1}
        bind:value={keyCol}
      />
    </div>

    <div class="grid w-full items-center gap-1.5">
      <Label for="primary-column">Value Column</Label>
      <Input
        id="primary-column"
        placeholder="A-Z"
        pattern="[A-Z]"
        maxlength={1}
        bind:value={valueCol}
      />
    </div>

    <div class="grid w-full items-center gap-1.5">
      <Label for="separator">Separator</Label>
      <Input id="separator" bind:value={separator} />
    </div>

    <Button
      disabled={!files?.length || !keyCol || !valueCol || isLoading}
      onclick={process}
    >
      Process
    </Button>

    {#if isLoading}
      <p class="text-gray-600">Processing file...</p>
    {:else if map.size > 0}
      <p>number of key: {map.size}</p>
    {/if}
  </div>
</div>
