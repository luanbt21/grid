function processFiles() {
  const files = document.getElementById("fileInput").files;
  if (files.length === 0) {
    alert("Please select at least one Excel file.");
    return;
  }

  let titleRow = 1; // Assuming first row contains titles
  let primaryCol = 1; // Default primary column index
  let gotTitle = false;
  let csvContent = ""; // This will hold our CSV data
  let csvRows = []; // Store each row as an array of strings

  for (let file of files) {
    // Check if the file is an Excel file
    const fileExt = file.name.split(".").pop();
    if (["xls", "xlsx"].includes(fileExt.toLowerCase())) {
      const reader = new FileReader();

      reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        // Iterate over all sheets
        workbook.SheetNames.forEach((sheetName) => {
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
          });

          // Process each row
          jsonData.forEach((row, idx) => {
            if (!gotTitle && idx === titleRow - 1) {
              gotTitle = true;
              console.log("Title row: ", row);
              csvRows.push(row);
              // Determine primary column index based on non-empty title cell
              for (let i = 0; i < row.length; i++) {
                if (row[i] !== "") {
                  primaryCol = i;
                  break;
                }
              }
            } else if (gotTitle && row[primaryCol] === "") {
              console.log("Skipping empty primary column row: ", row);
            } else if (idx >= titleRow) {
              csvRows.push(row);
            }
          });
        });
      };

      reader.readAsArrayBuffer(file);
    }
  }

function downloadCSV(csvContent: string) {
const blob = new Blob([csvContent], { type: "text/csv" });
const url = URL.createObjectURL(blob);
const downloadLink = document.getElementById("downloadLink");
downloadLink.href = url;
downloadLink.download = "merged_output.csv";
downloadLink.style.display = "block";
downloadLink.textContent = "Download Merged CSV";
}