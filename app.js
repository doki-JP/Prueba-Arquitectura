const { HyperFormula } = require('hyperformula'); // Ensure correct import
const XLSX = require('xlsx');
const express = require('express');
const path = require('path');

const app = express();
const PORT = 3000;
app.use(express.static(__dirname));
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});
app.listen(PORT, () => {
    console.log(`Server is running at http://localhost:${PORT}`);
});

const options = {
    licenseKey: 'gpl-v3' // Licencia gratis
}
// Funcion para leer el archivo Excel
function readExcelFile(filePath) {
    try {
        // Read the workbook
        console.log('Reading Excel file...', filePath);
        const workbook = XLSX.readFile(filePath);
        
        // Converción de Excel a un arreglo
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Encabezados esperados
        const expectedHeaders = ['Vehiculo', 'Fecha', 'Hora', 'Velocidad', 'Estatus', 'Lts', 'ADC', 'Bat(V)', 'Evento', 'GPS', 'Punto de Interes', 'Comentarios'];
        
        // Checar la primera fila de datos
        const headers = data[0] || [];
        
        // Verificación de encabezados (Que cumpla con los encabezados esperados)
        if (!arraysEqual(headers, expectedHeaders)) {
            console.error('¡ALERTA! Los encabezados del archivo no coinciden con el formato esperado.');
            console.error('Encabezados esperados:', expectedHeaders.join(', '));
            console.error('Encabezados encontrados:', headers.join(', '));
            
            // Error para detener la ejecución si los encabezados no coinciden
            // throw new Error('Validación de formato no exitosa');
        }
        

        // Instanciamiento de HyperFormula.
        // Se le pasa el arreglo de datos para que los procese HyperFormula
        // Intenté usar la función buildfromSheets, pero no funcionó.
        const hfInstance = HyperFormula.buildFromArray(data, options);

        // Ejemplo para obtener el valor de una celda
        // Aquí se puede cambiar la celda a la que se quiere acceder, incluso puede ser una celda que contenga una fórmula 
        // (Regresa el resultado de la fórmula).
        const value = hfInstance.getCellValue({ row: 0, col: 1, sheet: 0 });
        console.log('Cell value:', value);
        return data;

    } catch (err) {
        console.error('Error de lectura: ', err);
    }
}

// Función para la comparación de encabezados, compara dos arreglos y regresa true si son iguales
// y false si no lo son.
function arraysEqual(a, b) {
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) {
        if (a[i] !== b[i]) return false;
    }
    return true;
}

const data = readExcelFile('polla.xlsx');

console.log('Data from Excel:', data);


const n1 = 10;
const tableData = [['10', '20', '=SUM(' +n1+',B1)', '40'], ['50', '60', '70', '80']];
const Prueba = HyperFormula.buildFromArray(tableData, options);
// Importante, debes especificar todo, incluso la hoja, si no lo haces, lanza un error. 
// Expected value of type: SimpleCellAddress for config parameter: cellAddress
console.log(Prueba.getCellValue({ row: 0, col: 2, sheet:0 })); // 20 + 10 = 30, aquí ejecuta la fórmula, no regresa el string
console.log(Prueba.getCellValue({ row: 1, col: 0, sheet:0 })); // 50


// Instanciamiento de HyperFormula.
const hfInstance1 = HyperFormula.buildFromArray(data, options);
hfInstance1.addSheet('Sheet2');
// setCellContents: Se le pasan los datos a la hoja que se acaba de crear, en este caso 'Sheet2', y lo agrega.
hfInstance1.setCellContents({ row: 0, col: 0, sheet: 1 }, [["Nombre", "Edad"], ["Juan", 20], ["Pedro", 30],["Maria", 25]]);

console.log(hfInstance1.getSheetNames()); // ['Sheet1', 'Sheet2']
console.log(hfInstance1.getSheetValues(1)); // [['Nombre', 'Edad'], ['Juan', 20], ['Pedro', 30], ['Maria', 25]]
const result = hfInstance1.getCellValue({ sheet: 1, col: 1, row: 2 }); // 30

console.log("El resultado es: ",result);

// Filtrado de datos en la hoja 1
// Se filtran los datos de la hoja 1, en este caso se filtran los que tienen edad mayor a 20
const filter = hfInstance1.getSheetValues(1).filter(row => row[1] > 20);
console.log("Filtered data (Age > 20):", filter);

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