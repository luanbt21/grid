import { type ClassValue, clsx } from "clsx";
import { renderAsync } from "docx-preview";
import { twMerge } from "tailwind-merge";
import { utils, writeFile } from "xlsx";

export function renderDocx(element: HTMLElement, docxFile?: File) {
  if (!element || !docxFile) return;
  renderAsync(docxFile, element);
}

export function formatAmount(amount: number, currency: string) {
  if (currency === "VND") {
    return new Intl.NumberFormat("vi-VN", {
      style: "currency",
      currency: "VND",
    }).format(amount);
  }
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
  }).format(amount);
}

export function uppercase(node: HTMLInputElement) {
  const transform = () => {
    node.value = node.value.toUpperCase();
  };
  node.addEventListener("input", transform, { capture: true });
  transform();
  if (/[A-Z]/.test(node.value)) {
    node.value = "";
  }
}

type BaseDownloadProps = {
  filename: string;
  tool: "merge" | "corresponding" | "crpd";
};

export type DownloadExcelProps = BaseDownloadProps &
  (
    | {
        type: "a";
        data: unknown[][];
      }
    | {
        type: "json";
        data: Record<string, string | number>[];
      }
  );

export function downloadExcel({ type, data, filename, tool }: DownloadExcelProps) {
  const worksheet = type === "a" ? utils.aoa_to_sheet(data) : utils.json_to_sheet(data);
  const workbook = utils.book_new();
  utils.book_append_sheet(workbook, worksheet, "Sheet1");
  writeFile(workbook, `${tool}-${filename}`);
}

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}
