import { BlobReader, BlobWriter, ZipWriter } from "@zip.js/zip.js";
import { TextRun, patchDocument } from "docx";
import FileSaver from "file-saver";

export type FilePatch = {
  fileName: string;
  patches: Patch[];
};
export type Patch = {
  patchName: string;
  value: string;
};

export async function patchDocx(docxBuffer: ArrayBuffer, filePatches: FilePatch[]) {
  const zipFileWriter = new BlobWriter();
  const zipWriter = new ZipWriter(zipFileWriter);

  await Promise.all(
    filePatches.map(async (filePatch) => {
      const patches = Object.fromEntries(
        filePatch.patches.map(({ patchName, value }) => [
          patchName,
          {
            type: "paragraph" as const,
            children: [new TextRun(value)],
          },
        ]),
      );

      const result = await patchDocument({
        outputType: "blob",
        data: docxBuffer,
        patches,
      });

      return zipWriter.add(`${filePatch.fileName}.docx`, new BlobReader(result));
    }),
  );

  await zipWriter.close();
  FileSaver.saveAs(await zipFileWriter.getData(), `letter-${Date.now()}.zip`);
}
