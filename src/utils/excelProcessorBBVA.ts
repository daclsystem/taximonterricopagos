/**
 * Procesador específico para archivos BBVA usando ExcelJS
 */
import * as ExcelJS from 'exceljs';
import { ExcelData, ExcelSheet, ExcelRow } from '../types/excel';

export const processBBVAFile = (file: File): Promise<ExcelData> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = async (e) => {
      try {
        const data = e.target?.result;
        const workbook = new ExcelJS.Workbook();
        
        try {
          await workbook.xlsx.load(data as ArrayBuffer);
        } catch (excelError) {
          console.error('Error con ExcelJS:', excelError);
          throw new Error('No se pudo leer el archivo Excel. Verifica que el archivo no esté corrupto.');
        }
        
        console.log(`=== PROCESANDO ARCHIVO BBVA CON EXCELJS ===`);
        console.log(`Nombres de hojas encontradas:`, workbook.worksheets.map(ws => ws.name));
        
        const sheets: ExcelSheet[] = [];
        let totalRows = 0;
        
        // Procesar la primera hoja
        const worksheet = workbook.worksheets[0];
        console.log(`BBVA: Procesando hoja: ${worksheet.name}`);
        console.log(`BBVA: Dimensiones: ${worksheet.rowCount} filas x ${worksheet.columnCount} columnas`);
        
        // Leer todas las filas
        const allRows: any[][] = [];
        worksheet.eachRow((row, rowNumber) => {
          const rowData: any[] = [];
          row.eachCell((cell, colNumber) => {
            rowData[colNumber - 1] = cell.value;
          });
          allRows[rowNumber - 1] = rowData;
        });
        
        console.log(`BBVA: Total de filas leídas: ${allRows.length}`);
        console.log(`BBVA: Fila 31:`, allRows[30]); // Fila 31 (índice 30)
        console.log(`BBVA: Fila 32:`, allRows[31]); // Fila 32 (índice 31)
        
        // Verificar si la fila 31 tiene los headers de BBVA
        const row31 = allRows[30];
        if (row31 && row31.length > 0) {
          const row31Text = row31.map(cell => String(cell || '').toLowerCase()).join(' ');
          console.log(`BBVA: Texto de la fila 31: "${row31Text}"`);
          
          const hasSel = row31Text.includes('sel');
          const hasNo = row31Text.includes('no.') || row31Text.includes('no');
          const hasCuenta = row31Text.includes('cuenta');
          const hasTitularArchivo = row31Text.includes('titular(archivo)');
          const hasImporte = row31Text.includes('importe');
          
          if (hasSel && hasNo && hasCuenta && hasTitularArchivo && hasImporte) {
            console.log('✓ BBVA: Headers encontrados en fila 31');
            
            const headers = row31.map(cell => String(cell || '').trim());
            const dataStartIndex = 31; // Datos empiezan en fila 32 (índice 31)
            let dataEndIndex = allRows.length;
            
            // Buscar "Estimado Cliente:" para terminar
            for (let i = 40; i < allRows.length; i++) {
              const row = allRows[i];
              if (row && row.length > 0) {
                const rowText = row.map(cell => String(cell || '').toLowerCase()).join(' ');
                if (rowText.includes('estimado cliente')) {
                  dataEndIndex = i;
                  console.log(`BBVA: Fin de datos encontrado en fila ${i + 1}: "Estimado Cliente"`);
                  break;
                }
              }
            }
            
            console.log(`BBVA: Procesando datos desde fila ${dataStartIndex + 1} hasta fila ${dataEndIndex}`);
            
            // Extraer datos de la hoja
            const sheetData: ExcelRow[] = [];
            for (let i = dataStartIndex; i < dataEndIndex; i++) {
              const row = allRows[i];
              if (row && row.length > 0) {
                const rowData: ExcelRow = {};
                row.forEach((cell, colIndex) => {
                  const header = headers[colIndex];
                  if (header) {
                    rowData[header] = String(cell || '').trim();
                  }
                });
                sheetData.push(rowData);
              }
            }
            
            console.log(`BBVA: Datos extraídos: ${sheetData.length} filas`);
            console.log(`BBVA: Primera fila de datos:`, sheetData[0]);
            
            if (sheetData.length > 0) {
              sheets.push({
                name: worksheet.name,
                headers,
                data: sheetData
              });
              totalRows += sheetData.length;
            }
          } else {
            console.log('❌ BBVA: No se encontraron headers en fila 31');
            reject(new Error('No se encontraron headers de BBVA en la fila 31'));
            return;
          }
        } else {
          console.log('❌ BBVA: Fila 31 está vacía');
          reject(new Error('Fila 31 está vacía'));
          return;
        }
        
        const excelData: ExcelData = {
          fileName: file.name,
          sheets,
          totalRows,
          processedAt: new Date()
        };
        
        console.log(`=== RESULTADO FINAL BBVA ===`);
        console.log(`Archivo: ${file.name}`);
        console.log(`Hojas procesadas: ${sheets.length}`);
        console.log(`Total de filas: ${totalRows}`);
        
        resolve(excelData);
      } catch (error) {
        console.error('Error procesando archivo BBVA:', error);
        reject(error);
      }
    };
    
    reader.onerror = () => {
      reject(new Error('Error leyendo el archivo'));
    };
    
    reader.readAsArrayBuffer(file);
  });
};
