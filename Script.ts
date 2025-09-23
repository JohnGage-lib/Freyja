function main(workbook: ExcelScript.Workbook) {
    //Set the first row to be the 
    let firstSheet = workbook.getWorksheet("CircFineData");
    let secondSheet = workbook.getWorksheet("ItemRecord");
    // Run through the item list and then add the values if they are in the list
    let i = 1;
    let ScannedItemsBarcode: string[] = [];
    let CircItemBarcode: string[] = [];
    let secondSheetCount = 1;
    console.log('Item Barcode Array Running:')
    while (secondSheet.getCell(i, 10).getValue() != null && secondSheet.getCell(i, 10).getValue() != '') {
        ScannedItemsBarcode.push(secondSheet.getCell(i, 10).getValue().toString());
        i++;
    }
    console.log('Item Barcode Length:')
    console.log(ScannedItemsBarcode.length);
    // Reset to reiterate over CircItemBarcode
    i = 1;
    while (firstSheet.getCell(i, 1).getValue() != null && firstSheet.getCell(i, 1).getValue() != '') {
        CircItemBarcode.push(firstSheet.getCell(i, 1).getValue().toString());
        i++
    }
    console.log('Circ Barcode Length:')
    console.log(CircItemBarcode.length);
    secondSheet.getCell(0, 51).setValue('Circ Records');
    ScannedItemsBarcode.forEach((barcode) => {
        if (CircItemBarcode.indexOf(barcode) != null) {
            let test = CircItemBarcode.filter(circbarcode => circbarcode == barcode);
            console.log('Counting...');
            secondSheet.getCell(ScannedItemsBarcode.indexOf(barcode) + 1, 51).setValue(test.length);

        } else {
            secondSheet.getCell(ScannedItemsBarcode.indexOf(barcode) + 1, 51).setValue(0);
        }
    });
    console.log('Starting Text Count Conversion:');
    //Add total from Sierra Note
    ScannedItemsBarcode.forEach((barcode) => {
        let notes = secondSheet.getCell(ScannedItemsBarcode.indexOf(barcode) + 1, 32).getValue().toString().split(' ');
        let startIndex = notes.indexOf('checkouts%3A');
        if (startIndex != null && Number.isNaN(startIndex) && startIndex > 0) {
            let circCount = Number(notes[startIndex + 1]);
            console.log(circCount);
            secondSheet.getCell(ScannedItemsBarcode.indexOf(barcode) + 1, 51).setValue(Number(secondSheet.getCell(ScannedItemsBarcode.indexOf(barcode) + 1, 51).getValue()) + circCount)
        }
    });
}
