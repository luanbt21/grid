<script lang="ts">
  import Combobox from "$lib/components/combobox.svelte";
  import ExcelPreview from "$lib/components/excel-preview.svelte";
  import Button from "$lib/components/ui/button/button.svelte";
  import { Checkbox } from "$lib/components/ui/checkbox";
  import { Input } from "$lib/components/ui/input";
  import { Label } from "$lib/components/ui/label";
  import { Toggle } from "$lib/components/ui/toggle";
  import { renderDocx } from "$lib/utils";
  import { patchDetector } from "docx";
  import Presentation from "lucide-svelte/icons/presentation";
  import type { ChangeEventHandler } from "svelte/elements";
  import { read, utils } from "xlsx";
  import { patchDocx } from "./process";

  let patches = $state<Record<string, string>>({});

  let docxFile: File | undefined = $state();
  let excelFile: FileList | undefined = $state();

  let firstRow: Record<string, string | number> = $state({});
  let skipFirstRow = $state(false);

  let previewExcel = $state(true);
  let previewDocx = $state(false);

  let rawData: Record<string, string | number>[] = $state([]);

  const onTemplateChanged: ChangeEventHandler<HTMLInputElement> = async (
    event
  ) => {
    if (!event.currentTarget.files?.length) return;

    docxFile = event.currentTarget.files[0];
    const patchKeys = await patchDetector({
      data: await docxFile.arrayBuffer(),
    });
    for (const patchKey of patchKeys) {
      patches[patchKey] = "";
    }
  };

  const onDataChanged: ChangeEventHandler<HTMLInputElement> = async (event) => {
    if (!event.currentTarget.files?.length) {
      for (const patchKey of Object.keys(patches)) {
        patches[patchKey] = "";
      }
      return;
    }

    excelFile = event.currentTarget.files;
    const buffer = await excelFile[0].arrayBuffer();
    const workbook = read(buffer);

    // TODO: handle some index error here
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    rawData = utils.sheet_to_json(worksheet, {
      header: "A",
      blankrows: false,
      defval: "",
    });

    firstRow = rawData[0];
    for (const [col, title] of Object.entries(firstRow)) {
      for (const patchKey of Object.keys(patches)) {
        if (title === patchKey) {
          patches[patchKey] = col;
          skipFirstRow = true;
        }
      }
    }
  };

  $inspect(111, rawData);

  const process = async () => {
    if (!docxFile || !excelFile) {
      return;
    }

    const data = rawData.slice(skipFirstRow ? 1 : 0).map((row) => {
      return {
        fileName: row.A.toString(),
        patches: Object.entries(patches).map(([patchKey, col]) => {
          return {
            patchName: patchKey,
            value: String(row[col]),
          };
        }),
      };
    });

    patchDocx(await docxFile.arrayBuffer(), data);
  };
</script>

<svelte:head>
  <title>Fill Docx template with Excel</title>
</svelte:head>
<div class="flex flex-col gap-2 max-w-2xl mx-auto p-6">
  <h1 class="text-2xl font-bold mb-4">Fill Docx template with Excel</h1>

  <div class="grid w-full items-center gap-1.5">
    <Label for="template">Choose Docx template file</Label>
    <div class="flex gap-2">
      <Input
        id="template"
        type="file"
        accept=".docx"
        onchange={onTemplateChanged}
      />
      <Toggle bind:pressed={previewDocx}><Presentation /> Preview</Toggle>
    </div>
  </div>

  <div class="grid w-full items-center gap-1.5">
    <Label for="xlsx">Choose Excel file</Label>
    <div class="flex gap-2">
      <Input id="xlsx" type="file" accept=".xlsx" onchange={onDataChanged} />
      <Toggle bind:pressed={previewExcel}><Presentation /> Preview</Toggle>
    </div>
  </div>

  {#if Object.keys(firstRow).length > 0}
    <div class="items-top flex space-x-2">
      <Checkbox id="skip-first-row" bind:checked={skipFirstRow} />
      <div class="grid gap-1.5 leading-none">
        <Label
          for="skip-first-row"
          class="text-sm font-medium leading-none peer-disabled:cursor-not-allowed peer-disabled:opacity-70"
        >
          Skip first row
        </Label>
        <p class="text-muted-foreground text-sm">
          {Object.values(firstRow).join(" | ")}
        </p>
      </div>
    </div>
  {/if}

  {#if !docxFile}
    <h2 class="font-medium">
      Choose a docx template I will find all the
      <b>&lbrace;&lbrace;KEY&rbrace;&rbrace;</b> you want to replace
    </h2>
  {:else if Object.keys(patches).length === 0}
    <h2 class="font-medium">
      No keys found, I will try to find all the
      <b>&lbrace;&lbrace;KEY&rbrace;&rbrace;</b> you want to replace
    </h2>
  {:else}
    <h2 class="font-medium">Found {Object.keys(patches).length} keys</h2>
  {/if}
  {#each Object.keys(patches) as patchKey}
    <div class="mb-4 flex flex-row justify-between">
      <p>{patchKey}</p>
      {#if excelFile?.length}
        <Combobox
          data={Object.entries(firstRow).map(([key, value]) => ({
            value: key,
            label: `${key}: ${value}`,
          }))}
          bind:value={patches[patchKey]}
          placeholder="column"
        />
      {/if}
    </div>
  {/each}

  <Button
    class="w-full"
    disabled={!docxFile ||
      !excelFile ||
      !Object.keys(patches).length ||
      Object.values(patches).some((value) => !value)}
    onclick={process}>Process</Button
  >

  {#if previewDocx && docxFile}
    {#key docxFile}
      <div use:renderDocx={docxFile}></div>
    {/key}
  {/if}

  {#if previewExcel && excelFile?.length}
    <ExcelPreview files={excelFile} />
  {/if}
</div>
