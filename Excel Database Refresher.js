function main(workbook: ExcelScript.Workbook) {
        changesToMake(workbook);
        refresh(workbook);

        console.log("Exe Success!");
    }

function changesToMake(workbook: ExcelScript.Workbook) {
    // Define the sheets to keep
    const skip = [
        "Info",
        "Changes"
    ];

    // Get all worksheets in the workbook
    const allSheets = workbook.getWorksheets();
   
    // Get the "Changes" sheet (must already exist)
    const changesSheet = workbook.getWorksheet("Changes");
    const changesUsedRange = changesSheet.getUsedRange();
    const startRow = changesUsedRange ? changesUsedRange.getRowCount() : 0; // append below existing data
    let pasteRow = startRow;

    // Loop through each worksheet
    allSheets.forEach(sheet => {

        // Check if the sheet's name is NOT in the sheetsToKeep array
        if (!skip.includes(sheet.getName())) {
            console.log(sheet.getName());

            // Get the used range to determine how many rows there are
            const usedRange = sheet.getUsedRange();
            if (!usedRange) return; // skip empty sheets

            const values = usedRange.getValues();

            // BN = column 66 → zero-based index 65
            const bnIndex = 65;

            // Loop through rows (you can skip headers if needed)
            for (let i = 0; i < values.length; i++) {
                const row = values[i];

                // Check if BN column is boolean FALSE
                if (row[bnIndex] === false) {
                    // Copy columns A:L (indices 0–11)
                    const rowToCopy = row.slice(0, 12);

                    // Paste into "Changes" sheet
                    changesSheet.getRangeByIndexes(pasteRow, 0, 1, 12).setValues([rowToCopy]);
                    pasteRow++;
                }
            }
        }
    });
   
}

function editAll(workbook: ExcelScript.Workbook) {
    // Define the sheets to keep
    const skip = [
        "Info"
    ];

    // Get all worksheets in the workbook
    const allSheets = workbook.getWorksheets();

    // Loop through each worksheet
    allSheets.forEach(sheet => {

        // Check if the sheet's name is NOT in the sheetsToKeep array
        if (!skip.includes(sheet.getName())) {
            console.log(sheet.getName());

            let master = workbook.getWorksheet("Master");
            sheet.getRange("O:O").copyFrom(master.getRange("O:O"), ExcelScript.RangeCopyType.all, false, false);
            sheet.getRange("O:AA").setColumnHidden(true);
        }
    });
}

function refresh(workbook: ExcelScript.Workbook) {
    // Define the sheets to keep
    const skip = [
        "Info",
        "Changes"
    ];

    // Get all worksheets in the workbook
    const allSheets = workbook.getWorksheets();

    // Loop through each worksheet
    allSheets.forEach(sheet => {

        // Check if the sheet's name is NOT in the sheetsToKeep array
        if (!skip.includes(sheet.getName())) {
            console.log(sheet.getName());

            sheet.getRange("M:AA").setColumnHidden(false);

            let clearRange = sheet.getRange("A2:AA2").getExtendedRange(ExcelScript.KeyboardDirection.down);
            clearRange.clear(ExcelScript.ClearApplyTo.contents);

            let sourceRange = sheet.getRange("BA2:BL2").getExtendedRange(ExcelScript.KeyboardDirection.down);
            let destRange1 = sheet.getRange("A2");
            let destRange2 = sheet.getRange("P2");
            let values = sourceRange.getValues();
            let r = values.length;
            let c = values[0].length;
            let pasterange = destRange1.getResizedRange(r-1, c-1);
            let pasterange2 = destRange2.getResizedRange(r-1, c-1);

            pasterange.setValues(values);
            pasterange2.setValues(values);

            sheet.getRange("A:N").getFormat().autofitColumns();

            // Set border for range A2:K4701 on selectedSheet
            pasterange.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(ExcelScript.BorderLineStyle.continuous);
            pasterange.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setColor("000000");
            pasterange.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setWeight(ExcelScript.BorderWeight.thin);
            pasterange.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.continuous);
            pasterange.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical).setColor("000000");
            pasterange.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical).setWeight(ExcelScript.BorderWeight.thin);
           
            // Set border for range A1:K67 on selectedSheet
            pasterange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
            pasterange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setColor("000000");
            pasterange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.medium);
            pasterange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
            pasterange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setColor("000000");
            pasterange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.medium);
            pasterange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
            pasterange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setColor("000000");
            pasterange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.medium);
            pasterange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
            pasterange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setColor("000000");
            pasterange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.medium);

            // Set format for range D:D on selectedSheet
            sheet.getRange("D:D").setNumberFormatLocal("m/d/yyyy");
            // Set format for range I:J on selectedSheet
            sheet.getRange("I:K").setNumberFormatLocal("m/d/yyyy");

           
            // Set horizontal alignment
            sheet.getRange("A:L").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.general);
            sheet.getRange("A:L").getFormat().setIndentLevel(0);

            sheet.getAutoFilter().apply(sheet.getRange("A1:AA1"));

            sheet.getRange("M:AA").setColumnHidden(true);
            sheet.getRange("BA:BN").setColumnHidden(true);
        }
    });
}

function putFormulasinEachSheet(workbook: ExcelScript.Workbook) {
    // Define the sheets to keep
    const skip = [
        "Info",
        "TCS",
        "LJ",
        "ACO",
        "Changes"
    ];

    // Get all worksheets in the workbook
    const allSheets = workbook.getWorksheets();

    // Loop through each worksheet
    allSheets.forEach(sheet => {

        // Check if the sheet's name is NOT in the sheetsToKeep array
        if (!skip.includes(sheet.getName())) {
            console.log(sheet.getName());

            sheet.getAutoFilter().clearCriteria();

            sheet.getRange("O:BL").clear();
            console.log("Cleared!")
           
            sheet.getRange("BA1:BL1").copyFrom(sheet.getRange("A:L"), ExcelScript.RangeCopyType.all, false, false);

            sheet.getRange("A2:M2").getExtendedRange(ExcelScript.KeyboardDirection.down).clear(ExcelScript.ClearApplyTo.contents); //clear all information

            if (sheet.getName() == "Master") {
                let formulaCell = sheet.getRange("BA2"); //add formula
                formulaCell.setFormula('=SORT(VSTACK(FILTER(TCS!A:L, TCS!H:H = "Staying on HBC"), FILTER(LJ!A:L, LJ!H:H  = "Staying on HBC"), FILTER(ACO!A:L, ACO!H:H = "Staying on HBC")), 1, TRUE)');
            }
            else {
                let drName = sheet.getName();
                sheet.getRange("BA1").copyFrom(sheet.getRange("A1:N1"), ExcelScript.RangeCopyType.all, false, false);
                let formulaCell = sheet.getRange("BA2"); //add formula
                formulaCell.setFormula('=SORT(FILTER(Master!BA:BL,Master!BA:BA ="' + drName + '"), 5)');
            }
        }
    });
}