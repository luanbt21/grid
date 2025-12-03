<script lang="ts">
    import ExcelPreview from "$lib/components/excel-preview.svelte";
    import { Button } from "$lib/components/ui/button/index.js";
    import { Input } from "$lib/components/ui/input";
    import { Label } from "$lib/components/ui/label";
    import { downloadExcel } from "$lib/utils";
    import type { ChangeEventHandler } from "svelte/elements";
    import { read, utils } from "xlsx";

    type Row = Record<string, string | number>;

    let skipRow = $state(0);
    let CODE: string | undefined = $state();
    let ICODE: string | undefined = $state();
    let LEFT: string | undefined = $state();
    let RIGHT: string | undefined = $state();

    let separator = $state(", ");

    let map: Map<string, Set<string>> = $state(new Map());

    let files: FileList | null = $state(null);
    let isLoading = $state(false);
    let warning = "";

    const onFilesChanged: ChangeEventHandler<HTMLInputElement> = (event) => {
        files = event.currentTarget.files;
        map = new Map();
        isLoading = false;
        warning = "";
    };

    async function process() {
        if (!files?.length || !CODE || !LEFT || !RIGHT) {
            return;
        }

        isLoading = true;

        for (let file of files) {
            const buffer = await file.arrayBuffer();
            const workbook = read(buffer);

            for (const sheetName of workbook.SheetNames) {
                const worksheet = workbook.Sheets[sheetName];
                const data: Row[] = utils.sheet_to_json(worksheet, {
                    header: "A",
                    range: 0,
                    defval: "",
                    blankrows: true,
                });

                const output: Row[] = [];

                let rowsWithLeft: Row[] = [];
                let rowsWithRight: Row[] = [];

                let sumLeft = 0;
                let sumRight = 0;

                let isLeftFirst: boolean | null = null;
                for (const row of data) {
                    if (
                        !row[CODE] ||
                        Number.isNaN(Number(row[CODE])) ||
                        Number.isNaN(Number(row[LEFT]))
                    ) {
                        continue;
                    }

                    if (row[LEFT]) {
                        if (isLeftFirst === null) {
                            isLeftFirst = true;
                        }
                        rowsWithLeft.push(row);
                        // leftValues.push(Number(row[LEFT]));
                        sumLeft += Number(row[LEFT]);
                    }
                    if (row[RIGHT]) {
                        if (isLeftFirst === null) {
                            isLeftFirst = false;
                        }
                        rowsWithRight.push(row);
                        // rightValues.push(Number(row[RIGHT]));
                        sumRight += Number(row[RIGHT]);
                    }

                    if (sumLeft > 0 && sumLeft === sumRight) {
                        let finalLeftRows: Row[];
                        let finalRightRows: Row[];

                        if (rowsWithLeft.length === rowsWithRight.length) {
                            finalLeftRows = rowsWithLeft.map((leftRow, i) => ({
                                ...leftRow,
                                ...(ICODE && CODE
                                    ? {
                                          [ICODE]:
                                              rowsWithRight[i]?.[CODE] || "",
                                      }
                                    : {}),
                            }));
                            finalRightRows = rowsWithRight.map(
                                (rightRow, i) => ({
                                    ...rightRow,
                                    ...(ICODE && CODE
                                        ? {
                                              [ICODE]:
                                                  rowsWithLeft[i]?.[CODE] || "",
                                          }
                                        : {}),
                                }),
                            );
                        } else if (rowsWithLeft.length > rowsWithRight.length) {
                            // Assuming rowsWithRight.length is 1, spread it
                            const rightRow = rowsWithRight[0];

                            finalLeftRows = rowsWithLeft.map((leftRow) => ({
                                ...leftRow,
                                ...(ICODE && CODE
                                    ? { [ICODE]: rightRow?.[CODE] || "" }
                                    : {}),
                            }));

                            // Create new right rows based on the single right row, but spread across left values
                            finalRightRows = rowsWithLeft.map((leftRow) => ({
                                ...rightRow,
                                ...(ICODE && CODE
                                    ? { [ICODE]: leftRow?.[CODE] || "" }
                                    : {}),
                                ...(RIGHT && LEFT
                                    ? { [RIGHT]: leftRow[LEFT] || "" }
                                    : {}),
                            }));
                        } else {
                            // rowsWithLeft.length < rowsWithRight.length
                            // Assuming rowsWithLeft.length is 1, spread it
                            const leftRow = rowsWithLeft[0];

                            // Create new left rows based on the single left row, but spread across right values
                            finalLeftRows = rowsWithRight.map((rightRow) => ({
                                ...leftRow,
                                ...(ICODE && CODE
                                    ? { [ICODE]: rightRow?.[CODE] || "" }
                                    : {}),
                                ...(RIGHT && LEFT
                                    ? { [LEFT]: rightRow[LEFT] || "" }
                                    : {}),
                            }));

                            finalRightRows = rowsWithRight.map((rightRow) => ({
                                ...rightRow,
                                ...(ICODE && CODE
                                    ? { [ICODE]: leftRow?.[CODE] || "" }
                                    : {}),
                            }));
                        }

                        if (isLeftFirst) {
                            output.push(...finalLeftRows, ...finalRightRows);
                        } else {
                            output.push(...finalRightRows, ...finalLeftRows);
                        }

                        rowsWithLeft = [];
                        rowsWithRight = [];
                        sumLeft = 0;
                        sumRight = 0;
                        isLeftFirst = null;
                    }
                }

                downloadExcel({
                    type: "json",
                    data: output,
                    tool: "crpd",
                    filename: file.name,
                });
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
            <Label for="primary-column">Account Code</Label>
            <Input
                id="primary-column"
                placeholder="A-Z"
                pattern="[A-Z]"
                maxlength={1}
                bind:value={() => CODE, (v) => (CODE = v?.toUpperCase())}
            />
        </div>

        <div class="grid w-full items-center gap-1.5">
            <Label for="primary-column">Insert Coresponding Column</Label>
            <Input
                id="primary-column"
                placeholder="A-Z"
                pattern="[A-Z]"
                maxlength={1}
                bind:value={() => ICODE, (v) => (ICODE = v?.toUpperCase())}
            />
        </div>

        <div class="grid w-full items-center gap-1.5">
            <Label for="primary-column">Value Left Column</Label>
            <Input
                id="primary-column"
                placeholder="A-Z"
                pattern="[A-Z]"
                maxlength={1}
                bind:value={() => LEFT, (v) => (LEFT = v?.toUpperCase())}
            />
        </div>

        <div class="grid w-full items-center gap-1.5">
            <Label for="primary-column">Value Right Column</Label>
            <Input
                id="primary-column"
                placeholder="A-Z"
                pattern="[A-Z]"
                maxlength={1}
                bind:value={() => RIGHT, (v) => (RIGHT = v?.toUpperCase())}
            />
        </div>

        <!-- <div class="grid w-full items-center gap-1.5">
      <Label for="separator">Separator</Label>
      <Input id="separator" bind:value={separator} />
    </div> -->

        <Button
            disabled={!files?.length || !CODE || !LEFT || !RIGHT || isLoading}
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
