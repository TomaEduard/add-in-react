/* global Excel console */

export async function insertText(text: string) {
  // Write text to the selected cell.
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = context.workbook.getSelectedRange();
      range.values = [[text]];
      range.format.autofitColumns();
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function makeCellYellow() {
  try {
    await Excel.run(async (context) => {
      context.workbook.worksheets.getActiveWorksheet();
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = "yellow";
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function tooltip() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = context.workbook.getSelectedRange();
      await context.sync(); // Ensure the range is loaded

      const cell = range.getCell(0, 0);
      await context.sync(); // Ensure the cell is loaded

      // Create a rectangle shape
      const shape = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
      shape.fill.setSolidColor("red");

      // Position the rectangle above the selected cell
      shape.left = cell.left;
      shape.top = cell.top - shape.height;
      shape.width = cell.width;
      shape.height = 20; // Set a fixed height for the rectangle

      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function displayJson(jsonObject: object) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      let rowIndex = 0;

      for (const [key, value] of Object.entries(jsonObject)) {
        const rangeKey = sheet.getRange(`A${rowIndex + 1}`);
        const rangeValue = sheet.getRange(`B${rowIndex + 1}`);

        rangeKey.values = [[key]];
        rangeValue.values = [[JSON.stringify(value)]];

        // Set wrapText to true for both key and value cells
        rangeKey.format.wrapText = true;
        rangeValue.format.wrapText = true;

        // Set font size to 20px for both key and value cells
        rangeKey.format.font.size = 20;
        rangeValue.format.font.size = 20;

        rowIndex++;
      }

      // Set column widths
      sheet.getRange("A:A").format.columnWidth = 170;
      sheet.getRange("B:B").format.columnWidth = 780;

      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function createNewSheet(sheetName: string) {
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const newSheet = sheets.add(sheetName);
      newSheet.activate();
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}