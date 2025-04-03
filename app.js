const { HyperFormula } = require('hyperformula'); // Ensure correct import
const XLSX = require('xlsx');

const options = {
    licenseKey: 'gpl-v3'
};



function readExcelFile(filePath) {
    try {
        // Read the workbook
        console.log('Reading Excel file...', filePath);
        const workbook = XLSX.readFile(filePath);
        
        // Convert first sheet to 2D array
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Create HyperFormula instance
        const hfInstance = HyperFormula.buildFromArray(data, { 
            licenseKey: 'gpl-v3' 
        });

        // Get cell value (example)
        const value = hfInstance.getCellValue({ row: 0, col: 1, sheet: 0 });
        console.log('Cell value:', value);

        return data;
    } catch (err) {
        console.error('Error reading Excel file:', err);
    }
}

const data = readExcelFile('polla.xlsx');

console.log('Data from Excel:', data);





console.log('CACA');
n1 = 10;
const tableData = [['10', '20', '=SUM(' +n1+',B1)', '40'], ['50', '60', '70', '80']];



// Create a new instance of HyperFormula
const hfInstance1 = HyperFormula.buildFromArray(data, options);
hfInstance1.addSheet('Sheet2');
hfInstance1.setCellContents({ row: 0, col: 0, sheet: 1 }, [["Nombre", "Edad"], ["Juan", 20], ["Pedro", 30],["Maria", 25]]);


console.log(hfInstance1.getSheetNames()); // ['Sheet1', 'Sheet2']
const result = hfInstance1.getCellValue({ sheet: 1, col: 1, row: 2 });

console.log("El resultado es: ",result);

// Filter the data in Sheet2 where Age > 20
const filter = hfInstance1.getSheetValues(1).filter(row => row[1] > 20);
console.log("Filtered data (Age > 20):", filter);

/*
// Replace the value Juan with Pedro in Sheet2
const changes = hfInstance1.setSheetContent(1, [["Juan"]], [["Pedro"]]);
console.log("Changes made:", changes);
// Get the values from Sheet2
const sheet2Values = hfInstance1.getSheetValues(1);
console.log("Sheet2 values:", sheet2Values);
*/

// Get the calculated value of the cell C1
const value = hfInstance1.getCellValue({ row: 0, col: 2 , sheet: 0 });  
console.log(value); // 30

// Obtener los datos de todas las hojas de HyperFormula
function saveExcelFile(hfInstance, filePath) {
    try {
        const workbook = XLSX.utils.book_new(); // Crear un nuevo libro de trabajo

        // Iterar sobre todas las hojas de HyperFormula
        hfInstance.getSheetNames().forEach((sheetName, index) => {
            const sheetData = hfInstance.getSheetValues(index); // Obtener los valores de la hoja
            const worksheet = XLSX.utils.aoa_to_sheet(sheetData); // Convertir a formato de hoja de XLSX
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetName); // Agregar la hoja al libro
        });

        // Escribir el libro en un archivo
        XLSX.writeFile(workbook, filePath);
        console.log(`Archivo guardado exitosamente en: ${filePath}`);
    } catch (err) {
        console.error('Error al guardar el archivo Excel:', err);
    }
}

// Guardar el archivo después de las ediciones
saveExcelFile(hfInstance1, 'resultado.xlsx');