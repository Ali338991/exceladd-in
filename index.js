const dataSourceUrl = "https://excel-add-in.surge.sh";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    let fileInput = document.getElementById("fileInput");
    fileInput.addEventListener("change", insertSheets);
  }
});

async function insertSheets() {
  try {
    const myFile = document.getElementById("fileInput").files[0];

    if (!myFile) {
      console.error("No file selected.");
      return;
    }

    const reader = new FileReader();

    reader.onload = async (event) => {
      Excel.run(async (context) => {
        try {
          const startIndex = reader.result.toString().indexOf("base64,");
          const workbookContents = reader.result.toString().substr(startIndex + 7);

          const workbook = context.workbook;

          const options = {
            sheetNamesToInsert: ["TemplateAli"],
            positionType: Excel.WorksheetPositionType.after,
            relativeTo: "Sheet1",
          };

          workbook.insertWorksheetsFromBase64(workbookContents, options);

          await context.sync();

          const sheet = context.workbook.worksheets.getItem("TemplateAli");

          let response = await fetch(dataSourceUrl + "/data.json");
          if (response.ok) {
            const json = await response.json();
            const newSalesData = json.salesData.map((item) => [
              item.PRODUCT,
              item.QTR1,
              item.QTR2,
              item.QTR3,
              item.QTR4,
            ]);

            const startRow = 5;
            const address = "B" + startRow + ":F" + (newSalesData.length + startRow - 1);

            const range = sheet.getRange(address);
            range.values = newSalesData;
            sheet.activate();
            await context.sync();
          } else {
            console.error("HTTP-Error: " + response.status);
          }
        } catch (error) {
          console.log('error==>',error);
        }
      });
    };

    reader.readAsBinaryString(myFile);
  } catch (error) {
    console.error(error);
  }
}
