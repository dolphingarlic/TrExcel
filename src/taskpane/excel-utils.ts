/* global console, Excel */

import { ColumnWidth, Height, Width } from "./constants";

export async function getFreshSheet() {
  try {
    await Excel.run(async context => {
      let worksheets: Excel.WorksheetCollection = context.workbook.worksheets;
      worksheets.load("items/name");
      await context.sync();

      context.application.suspendScreenUpdatingUntilNextSync();

      // Create a temporary worksheet so Excel doesn't freak out about deleting the last worksheet
      const randomName: string = (Math.random() + 1).toString(36).substring(2);
      const tempWorksheet: Excel.Worksheet = worksheets.add(randomName);
      tempWorksheet.activate();

      // Delete old sheet
      for (let i = 0; i < worksheets.items.length; i++) {
        const ws: Excel.Worksheet = worksheets.items[i];
        if (ws.name === "TrExcel") ws.delete();
      }

      // New blank sheet + set as active
      const activeWorksheet: Excel.Worksheet = worksheets.add("TrExcel");
      context.workbook.properties.custom.add("TrExcel", "true");
      activeWorksheet.activate();

      // Format the game board
      const gameBoard = activeWorksheet.getRangeByIndexes(1, 1, Height, Width);
      gameBoard.format.columnWidth = ColumnWidth;
      gameBoard.format.borders.getItem("EdgeTop").style = "Continuous";
      gameBoard.format.borders.getItem("EdgeRight").style = "Continuous";
      gameBoard.format.borders.getItem("EdgeBottom").style = "Continuous";
      gameBoard.format.borders.getItem("EdgeLeft").style = "Continuous";

      const scoreCells = activeWorksheet.getRangeByIndexes(1, Width + 2, 1, 1);
      scoreCells.format.font.bold = true;
      scoreCells.values = [["SCORE: 0"]];

      const nextPieceCells = activeWorksheet.getRangeByIndexes(3, Width + 3, 4, 4);
      nextPieceCells.format.columnWidth = ColumnWidth;
      const nextPieceLabelCell = nextPieceCells.getOffsetRange(1, -1).getResizedRange(-3, -3);
      nextPieceLabelCell.format.font.bold = true;
      nextPieceLabelCell.values = [["NEXT:"]]

      const pocketCells = activeWorksheet.getRangeByIndexes(8, Width + 3, 4, 4);
      pocketCells.format.columnWidth = ColumnWidth;
      pocketCells.format.columnWidth = ColumnWidth;
      const pocketLabelCell = pocketCells.getOffsetRange(1, -1).getResizedRange(-3, -3);
      pocketLabelCell.format.font.bold = true;
      pocketLabelCell.values = [["POCKET:"]]

      // Delete the temporary worksheet
      tempWorksheet.delete();

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
