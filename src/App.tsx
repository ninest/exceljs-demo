import React, { useState } from "react";
import ExcelJS from "exceljs";

function getDataFromSheet(sheet: ExcelJS.Worksheet) {
  return sheet
    .getSheetValues()
    .slice(1)
    .map((r) => r.slice(1));
}

function App(): JSX.Element {
  const [workbook, setWorkbook] = useState<ExcelJS.Workbook>();

  const submit = async (e: React.FormEvent<HTMLFormElement>): Promise<void> => {
    e.preventDefault();

    const fileInput = e.currentTarget.elements.namedItem("excelfile") as HTMLInputElement;

    if (fileInput.files && fileInput.files.length > 0) {
      const excelFile = fileInput.files[0];
      const workbook = new ExcelJS.Workbook();
      const reader = new FileReader();

      reader.onload = async (event: ProgressEvent<FileReader>): Promise<void> => {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        await workbook.xlsx.load(data.buffer);

        setWorkbook(workbook);
      };

      reader.readAsArrayBuffer(excelFile);
    }
  };

  return (
    <>
      <main>
        <form onSubmit={submit}>
          <input type="file" name="excelfile" />
          <input type="submit" />
        </form>

        {workbook && (
          <div>
            {workbook.worksheets.map((sheet) => {
              return (
                <div key={sheet.id}>
                  <h2>{sheet.name}</h2>
                  <pre>{JSON.stringify(getDataFromSheet(sheet), null, 2)}</pre>
                </div>
              );
            })}
          </div>
        )}
      </main>
    </>
  );
}

export default App;
