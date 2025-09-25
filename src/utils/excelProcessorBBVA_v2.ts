/**
 * Procesador específico para archivos BBVA usando XLSX con configuración robusta
 */
import * as XLSX from 'xlsx';
import { ExcelData, ExcelSheet, ExcelRow } from '../types/excel';

export const processBBVAFile = (file: File): Promise<ExcelData> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        
        console.log(`=== PROCESANDO ARCHIVO BBVA CON XLSX ROBUSTO ===`);
        
        // Leer el archivo con diferentes opciones
        let workbook;
        try {
          workbook = XLSX.read(data, { 
            type: 'array',
            cellDates: true,
            cellNF: false,
            cellText: false
          });
        } catch (xlsxError) {
          console.error('Error con XLSX:', xlsxError);
          throw new Error('No se pudo leer el archivo Excel. Verifica que el archivo no esté corrupto.');
        }
        
        console.log(`Nombres de hojas encontradas:`, workbook.SheetNames);
        
        const sheets: ExcelSheet[] = [];
        let totalRows = 0;
        
        // Buscar en TODAS las hojas hasta encontrar la que tiene los datos de BBVA
        let foundData = false;
        let jsonData: any[][] = [];
        let sheetName = '';
        
        for (let i = 0; i < workbook.SheetNames.length; i++) {
          const currentSheetName = workbook.SheetNames[i];
          const worksheet = workbook.Sheets[currentSheetName];
          
          console.log(`BBVA: Procesando hoja ${i + 1}/${workbook.SheetNames.length}: ${currentSheetName}`);
          console.log(`BBVA: Rango de la hoja: ${worksheet['!ref']}`);
          
          // Leer datos con múltiples intentos
          let currentData: any[][] = [];
          
          // Intento 1: sheet_to_json con header: 1
          try {
            currentData = XLSX.utils.sheet_to_json(worksheet, { 
              header: 1, 
              defval: '', 
              blankrows: true,
              raw: false
            });
            console.log(`BBVA: Hoja ${currentSheetName} - Datos con header:1 - ${currentData.length} filas`);
          } catch (e) {
            console.log(`Error con header:1 en hoja ${currentSheetName}`);
          }
          
          // Intento 2: sheet_to_json con header: 0
          if (currentData.length === 0) {
            try {
              currentData = XLSX.utils.sheet_to_json(worksheet, { 
                header: 0, 
                defval: '', 
                blankrows: true,
                raw: false
              });
              console.log(`BBVA: Hoja ${currentSheetName} - Datos sin header - ${currentData.length} filas`);
            } catch (e) {
              console.log(`Error sin header en hoja ${currentSheetName}`);
            }
          }
          
          // Intento 3: sheet_to_array
          if (currentData.length === 0) {
            try {
              currentData = XLSX.utils.sheet_to_array(worksheet, { 
                defval: '', 
                blankrows: true 
              });
              console.log(`BBVA: Hoja ${currentSheetName} - Datos con sheet_to_array - ${currentData.length} filas`);
            } catch (e) {
              console.log(`Error con sheet_to_array en hoja ${currentSheetName}`);
            }
          }
          
          // Verificar si esta hoja tiene los datos de BBVA (buscar fila 31 con headers)
          if (currentData.length > 31) {
            const row31 = currentData[30]; // Fila 31 (índice 30)
            if (row31 && row31.length > 0) {
              const row31Text = row31.map(cell => String(cell || '').toLowerCase()).join(' ');
              console.log(`BBVA: Hoja ${currentSheetName} - Fila 31: "${row31Text}"`);
              
              // Verificar si tiene los headers de BBVA
              const hasSel = row31Text.includes('sel');
              const hasNo = row31Text.includes('no.') || row31Text.includes('no');
              const hasCuenta = row31Text.includes('cuenta');
              const hasTitularArchivo = row31Text.includes('titular(archivo)');
              const hasImporte = row31Text.includes('importe');
              
              if (hasSel && hasNo && hasCuenta && hasTitularArchivo && hasImporte) {
                console.log(`✓ BBVA: ¡Datos encontrados en hoja ${currentSheetName}!`);
                jsonData = currentData;
                sheetName = currentSheetName;
                foundData = true;
                break;
              }
            }
          }
        }
        
        if (!foundData) {
          throw new Error('No se encontraron datos de BBVA en ninguna hoja');
        }
        
        console.log(`BBVA: Usando hoja: ${sheetName} con ${jsonData.length} filas`);
        
        console.log(`BBVA: Total de filas leídas: ${jsonData.length}`);
        console.log(`BBVA: Fila 31:`, jsonData[30]); // Fila 31 (índice 30)
        console.log(`BBVA: Fila 32:`, jsonData[31]); // Fila 32 (índice 31)
        
        // Verificar si la fila 31 tiene los headers de BBVA
        const row31 = jsonData[30];
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
            let dataEndIndex = jsonData.length;
            
            // Buscar "Estimado Cliente:" para terminar
            for (let i = 40; i < jsonData.length; i++) {
              const row = jsonData[i];
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
              const row = jsonData[i];
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
                name: sheetName,
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
