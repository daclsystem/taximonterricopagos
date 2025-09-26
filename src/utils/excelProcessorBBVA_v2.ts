/**
 * Procesador específico para archivos BBVA usando XLSX con configuración robusta
 */
import * as ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { ExcelData, ExcelSheet, ExcelRow, AbonoRecord, CombinedData } from '../types/excel';

// Función para validar archivos Excel antes de procesarlos
const validateExcelFile = (file: File): boolean => {
  const validTypes = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel'
  ];
  const validExtensions = ['.xlsx', '.xls'];
  
  const hasValidType = validTypes.includes(file.type);
  const hasValidExtension = validExtensions.some(ext => 
    file.name.toLowerCase().endsWith(ext)
  );
  
  return hasValidType || hasValidExtension;
};

// Función para obtener el valor de una columna manejando variaciones de acentos y caracteres extraños
const getColumnValue = (rowObj: any, columnName: string, fallbackName?: string): string => {
  // Intentar con el nombre exacto primero
  if (rowObj[columnName] !== undefined) {
    return String(rowObj[columnName] || '');
  }
  
  // Si hay un nombre alternativo, intentarlo
  if (fallbackName && rowObj[fallbackName] !== undefined) {
    return String(rowObj[fallbackName] || '');
  }
  
  // Buscar en todas las claves que contengan la palabra (sin acentos y sin caracteres extraños)
  const normalizedColumnName = columnName.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  
  for (const key of Object.keys(rowObj)) {
    // Normalizar la clave y limpiar caracteres extraños
    const normalizedKey = key
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '') // Quitar acentos
      .replace(/[^\w\s]/g, '') // Quitar caracteres especiales excepto letras, números y espacios
      .toLowerCase();
    
    // Buscar coincidencias parciales
    if (normalizedKey.includes(normalizedColumnName) || normalizedColumnName.includes(normalizedKey)) {
      return String(rowObj[key] || '');
    }
  }
  
  // Buscar por patrones específicos para casos conocidos
  if (columnName.toLowerCase().includes('situacion')) {
    for (const key of Object.keys(rowObj)) {
      if (key.toLowerCase().includes('situacion') || key.toLowerCase().includes('situaci')) {
        return String(rowObj[key] || '');
      }
    }
  }
  
  return '';
};

export const processBBVAFile = (file: File): Promise<ExcelData> => {
  return new Promise((resolve, reject) => {
    // Validar el archivo antes de procesarlo
    if (!validateExcelFile(file)) {
      reject(new Error('El archivo no es un archivo Excel válido (.xlsx o .xls)'));
      return;
    }
    
    const reader = new FileReader();
    
    reader.onload = async (e) => {
      try {
        const data = e.target?.result;
        
        console.log(`=== PROCESANDO ARCHIVO BBVA CON EXCELJS ===`);
        console.log(`Tipo de archivo: ${file.type}`);
        console.log(`Tamaño del archivo: ${file.size} bytes`);
        
        // Validar que el archivo no esté vacío
        if (!data || (data as ArrayBuffer).byteLength === 0) {
          throw new Error('El archivo está vacío o no se pudo leer correctamente');
        }
        
        // Crear workbook y cargar el archivo con manejo de errores mejorado
        const workbook = new ExcelJS.Workbook();
        
        try {
          // Determinar el tipo de archivo y cargar apropiadamente
          const isXLS = file.name.toLowerCase().endsWith('.xls') || file.type === 'application/vnd.ms-excel';
          
          if (isXLS) {
            console.log('Detectado archivo .xls, cargando con XLSX (legacy)...');
            // Para archivos .xls, usar XLSX que sí los soporta
            const workbookXLSX = XLSX.read(data, { type: 'array' });
            
            // Convertir a formato compatible con nuestro sistema
            const sheets: ExcelSheet[] = [];
            let totalRows = 0;
            
            // Buscar en TODAS las hojas hasta encontrar la que tiene los datos de BBVA
            let foundData = false;
            let jsonData: any[][] = [];
            let sheetName = '';
            
            for (let i = 0; i < workbookXLSX.SheetNames.length; i++) {
              const currentSheetName = workbookXLSX.SheetNames[i];
              const worksheet = workbookXLSX.Sheets[currentSheetName];
              
              console.log(`BBVA: Procesando hoja ${i + 1}/${workbookXLSX.SheetNames.length}: ${currentSheetName}`);
              
              // Leer datos de la hoja
              const currentData = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1, 
                defval: '', 
                blankrows: true,
                raw: false
              }) as any[][];
              
              console.log(`BBVA: Hoja ${currentSheetName} - ${currentData.length} filas`);
              
              // Verificar si esta hoja tiene los datos de BBVA (buscar en todas las filas)
              if (currentData.length > 0) {
                let foundHeaders = false;
                let headerRowIndex = -1;
                let dataEndIndex = currentData.length;
                
                console.log(`BBVA: Buscando headers en hoja ${currentSheetName} con ${currentData.length} filas`);
                
                // Buscar headers en todas las filas de la hoja
                for (let rowIndex = 0; rowIndex < currentData.length; rowIndex++) {
                  const row = currentData[rowIndex];
                  
                  // Verificar si la fila tiene contenido
                  if (row && row.length > 0) {
                    const rowText = row.map(cell => String(cell || '').trim()).join(' ');
                    const rowTextLower = rowText.toLowerCase();
                    
                    console.log(`BBVA: Hoja ${currentSheetName} - Fila ${rowIndex + 1}: "${rowText.substring(0, 100)}..."`);
                    
                    // Si encuentra "Estimado cliente", detener la búsqueda antes de esta fila
                    if (rowTextLower.includes('estimado cliente')) {
                      dataEndIndex = rowIndex;
                      console.log(`BBVA: Encontrado "Estimado cliente" en fila ${rowIndex + 1}, deteniendo búsqueda de headers`);
                      break;
                    }
                    
                    // Verificar si esta fila tiene los headers de BBVA
                    const hasSel = rowTextLower.includes('sel');
                    const hasNo = rowTextLower.includes('no.') || rowTextLower.includes('no');
                    const hasCuenta = rowTextLower.includes('cuenta');
                    const hasTitularArchivo = rowTextLower.includes('titular(archivo)');
                    const hasImporte = rowTextLower.includes('importe');
                    
                    if (hasSel && hasNo && hasCuenta && hasTitularArchivo && hasImporte) {
                      console.log(`✓ BBVA: ¡Headers encontrados en fila ${rowIndex + 1} en hoja ${currentSheetName}!`);
                      jsonData = currentData;
                      sheetName = currentSheetName;
                      foundData = true;
                      foundHeaders = true;
                      headerRowIndex = rowIndex;
                      // Guardar también el índice de fin de datos
                      (jsonData as any).headerRowIndex = headerRowIndex;
                      (jsonData as any).dataEndIndex = dataEndIndex;
                      break;
                    }
                  } else {
                    // Si la fila está vacía, continuar buscando
                    console.log(`BBVA: Fila ${rowIndex + 1} está vacía, continuando búsqueda...`);
                  }
                }
                
                if (foundHeaders) {
                  break;
                }
              }
            }
            
            if (!foundData) {
              throw new Error('No se encontraron datos de BBVA en ninguna hoja');
            }
            
            console.log(`BBVA: Usando hoja: ${sheetName} con ${jsonData.length} filas`);
            
            // Obtener los índices que se calcularon durante la búsqueda
            const headerRowIndex = (jsonData as any).headerRowIndex;
            const dataEndIndex = (jsonData as any).dataEndIndex || jsonData.length;
            
            console.log(`BBVA: Total de filas leídas: ${jsonData.length}`);
            console.log(`BBVA: Headers encontrados en fila ${headerRowIndex + 1} (índice ${headerRowIndex})`);
            console.log(`BBVA: Fin de datos en fila ${dataEndIndex + 1} (índice ${dataEndIndex})`);
            console.log(`BBVA: Fila de headers:`, jsonData[headerRowIndex]);
            console.log(`BBVA: Fila siguiente:`, jsonData[headerRowIndex + 1]);
            
            // Verificar si la fila de headers tiene los headers de BBVA
            const headerRow = jsonData[headerRowIndex];
            if (headerRow && headerRow.length > 0) {
              const headerRowText = headerRow.map(cell => String(cell || '').toLowerCase()).join(' ');
              console.log(`BBVA: Texto de la fila de headers: "${headerRowText}"`);
              
              const hasSel = headerRowText.includes('sel');
              const hasNo = headerRowText.includes('no.') || headerRowText.includes('no');
              const hasCuenta = headerRowText.includes('cuenta');
              const hasTitularArchivo = headerRowText.includes('titular(archivo)');
              const hasImporte = headerRowText.includes('importe');
              
              if (hasSel && hasNo && hasCuenta && hasTitularArchivo && hasImporte) {
                console.log(`✓ BBVA: Headers confirmados en fila ${headerRowIndex + 1}`);
                
                const headers = headerRow.map(cell => String(cell || '').trim());
                const dataStartIndex = headerRowIndex + 1; // Datos empiezan en la fila siguiente
                
                console.log(`BBVA: Procesando datos desde fila ${dataStartIndex + 1} hasta fila ${dataEndIndex + 1}`);
                
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
                    data: sheetData,
                    rowCount: sheetData.length
                  });
                  totalRows += sheetData.length;
                }
              } else {
                console.log(`❌ BBVA: No se encontraron headers en fila ${headerRowIndex + 1}`);
                throw new Error(`No se encontraron headers de BBVA en la fila ${headerRowIndex + 1}`);
              }
            } else {
              console.log(`❌ BBVA: Fila de headers está vacía`);
              throw new Error('Fila de headers está vacía');
            }
            
            const excelData: ExcelData = {
              fileName: file.name,
              sheets,
              totalRows,
              uploadedAt: new Date()
            };
            
            console.log(`=== RESULTADO FINAL BBVA ===`);
            console.log(`Archivo: ${file.name}`);
            console.log(`Hojas procesadas: ${sheets.length}`);
            console.log(`Total de filas: ${totalRows}`);
            
            resolve(excelData);
            return;
          } else {
            console.log('Detectado archivo .xlsx, cargando con ExcelJS (moderno)...');
            await workbook.xlsx.load(data as ArrayBuffer);
          }
        } catch (loadError) {
          console.error('Error específico al cargar el archivo:', loadError);
          
          // Si es un error de ZIP, intentar con diferentes opciones
          if (loadError instanceof Error && loadError.message.includes('zip')) {
            console.log('Intentando cargar con opciones alternativas...');
            
            // Crear un nuevo workbook y intentar con opciones diferentes
            const workbook2 = new ExcelJS.Workbook();
            try {
              // Intentar con opciones de carga más permisivas
              await workbook2.xlsx.load(data as ArrayBuffer, {
                ignoreNodes: ['xl/styles.xml', 'xl/theme/theme1.xml']
              });
              console.log('Archivo cargado exitosamente con opciones alternativas');
            } catch (secondError) {
              throw new Error(`No se pudo leer el archivo Excel. El archivo puede estar corrupto o no ser un archivo Excel válido. Error: ${loadError instanceof Error ? loadError.message : 'Error desconocido'}`);
            }
          } else {
            throw new Error(`Error al cargar el archivo Excel: ${loadError instanceof Error ? loadError.message : 'Error desconocido'}`);
          }
        }
        
        console.log(`Nombres de hojas encontradas:`, workbook.worksheets.map(ws => ws.name));
        
        const sheets: ExcelSheet[] = [];
        let totalRows = 0;
        
        // Buscar en TODAS las hojas hasta encontrar la que tiene los datos de BBVA
        let foundData = false;
        let jsonData: any[][] = [];
        let sheetName = '';
        
        for (let i = 0; i < workbook.worksheets.length; i++) {
          const worksheet = workbook.worksheets[i];
          const currentSheetName = worksheet.name;
          
          console.log(`BBVA: Procesando hoja ${i + 1}/${workbook.worksheets.length}: ${currentSheetName}`);
          
          // Convertir la hoja a datos
          const currentData: any[][] = [];
          
          worksheet.eachRow((row) => {
            const rowData: any[] = [];
            row.eachCell((cell, colNumber) => {
              rowData[colNumber - 1] = cell.value;
            });
            currentData.push(rowData);
          });
          
          console.log(`BBVA: Hoja ${currentSheetName} - ${currentData.length} filas`);
          
          // Verificar si esta hoja tiene los datos de BBVA (buscar en todas las filas)
          if (currentData.length > 0) {
            let foundHeaders = false;
            let headerRowIndex = -1;
            let dataEndIndex = currentData.length;
            
            console.log(`BBVA: Buscando headers en hoja ${currentSheetName} con ${currentData.length} filas`);
            
            // Buscar headers en todas las filas de la hoja
            for (let rowIndex = 0; rowIndex < currentData.length; rowIndex++) {
              const row = currentData[rowIndex];
              
              // Verificar si la fila tiene contenido
              if (row && row.length > 0) {
                const rowText = row.map(cell => String(cell || '').trim()).join(' ');
                const rowTextLower = rowText.toLowerCase();
                
                console.log(`BBVA: Hoja ${currentSheetName} - Fila ${rowIndex + 1}: "${rowText.substring(0, 100)}..."`);
                
                // Si encuentra "Estimado cliente", detener la búsqueda antes de esta fila
                if (rowTextLower.includes('estimado cliente')) {
                  dataEndIndex = rowIndex;
                  console.log(`BBVA: Encontrado "Estimado cliente" en fila ${rowIndex + 1}, deteniendo búsqueda de headers`);
                  break;
                }
                
                // Verificar si esta fila tiene los headers de BBVA
                const hasSel = rowTextLower.includes('sel');
                const hasNo = rowTextLower.includes('no.') || rowTextLower.includes('no');
                const hasCuenta = rowTextLower.includes('cuenta');
                const hasTitularArchivo = rowTextLower.includes('titular(archivo)');
                const hasImporte = rowTextLower.includes('importe');
                
                if (hasSel && hasNo && hasCuenta && hasTitularArchivo && hasImporte) {
                  console.log(`✓ BBVA: ¡Headers encontrados en fila ${rowIndex + 1} en hoja ${currentSheetName}!`);
                  jsonData = currentData;
                  sheetName = currentSheetName;
                  foundData = true;
                  foundHeaders = true;
                  headerRowIndex = rowIndex;
                  // Guardar también el índice de fin de datos
                  (jsonData as any).headerRowIndex = headerRowIndex;
                  (jsonData as any).dataEndIndex = dataEndIndex;
                  break;
                }
              } else {
                // Si la fila está vacía, continuar buscando
                console.log(`BBVA: Fila ${rowIndex + 1} está vacía, continuando búsqueda...`);
              }
            }
            
            if (foundHeaders) {
              break;
            }
          }
        }
        
        if (!foundData) {
          throw new Error('No se encontraron datos de BBVA en ninguna hoja');
        }
        
        console.log(`BBVA: Usando hoja: ${sheetName} con ${jsonData.length} filas`);
        
        // Obtener los índices que se calcularon durante la búsqueda
        const headerRowIndex = (jsonData as any).headerRowIndex;
        const dataEndIndex = (jsonData as any).dataEndIndex || jsonData.length;
        
        console.log(`BBVA: Total de filas leídas: ${jsonData.length}`);
        console.log(`BBVA: Headers encontrados en fila ${headerRowIndex + 1} (índice ${headerRowIndex})`);
        console.log(`BBVA: Fin de datos en fila ${dataEndIndex + 1} (índice ${dataEndIndex})`);
        console.log(`BBVA: Fila de headers:`, jsonData[headerRowIndex]);
        console.log(`BBVA: Fila siguiente:`, jsonData[headerRowIndex + 1]);
        
        // Verificar si la fila de headers tiene los headers de BBVA
        const headerRow = jsonData[headerRowIndex];
        if (headerRow && headerRow.length > 0) {
          const headerRowText = headerRow.map(cell => String(cell || '').toLowerCase()).join(' ');
          console.log(`BBVA: Texto de la fila de headers: "${headerRowText}"`);
          
          const hasSel = headerRowText.includes('sel');
          const hasNo = headerRowText.includes('no.') || headerRowText.includes('no');
          const hasCuenta = headerRowText.includes('cuenta');
          const hasTitularArchivo = headerRowText.includes('titular(archivo)');
          const hasImporte = headerRowText.includes('importe');
          
          if (hasSel && hasNo && hasCuenta && hasTitularArchivo && hasImporte) {
            console.log(`✓ BBVA: Headers confirmados en fila ${headerRowIndex + 1}`);
            
            const headers = headerRow.map(cell => String(cell || '').trim());
            const dataStartIndex = headerRowIndex + 1; // Datos empiezan en la fila siguiente
            
            console.log(`BBVA: Procesando datos desde fila ${dataStartIndex + 1} hasta fila ${dataEndIndex + 1}`);
            
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
                data: sheetData,
                rowCount: sheetData.length
              });
              totalRows += sheetData.length;
            }
          } else {
            console.log(`❌ BBVA: No se encontraron headers en fila ${headerRowIndex + 1}`);
            reject(new Error(`No se encontraron headers de BBVA en la fila ${headerRowIndex + 1}`));
            return;
          }
        } else {
          console.log(`❌ BBVA: Fila de headers está vacía`);
          reject(new Error('Fila de headers está vacía'));
          return;
        }
        
        const excelData: ExcelData = {
          fileName: file.name,
          sheets,
          totalRows,
          uploadedAt: new Date()
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

export const createSingleFileData = (data: ExcelData): CombinedData => {
  const records: AbonoRecord[] = [];
  
  console.log(`=== PROCESANDO ARCHIVO BBVA: ${data.fileName} ===`);
  
  data.sheets.forEach((sheet, sheetIndex) => {
    console.log(`BBVA - Procesando hoja ${sheetIndex}: ${sheet.name}`);
    console.log(`BBVA - Headers de la hoja:`, sheet.headers);
    
    // Debug: mostrar las primeras 3 filas de datos para verificar
    console.log(`BBVA - Primera fila de datos:`, sheet.data[0]);
    console.log(`BBVA - Segunda fila de datos:`, sheet.data[1]);
    console.log(`BBVA - Tercera fila de datos:`, sheet.data[2]);
    
    sheet.data.forEach((row, index) => {
      // Los datos vienen como objetos con las cabeceras como propiedades
      if (typeof row === 'object' && row !== null && !Array.isArray(row)) {
        const rowObj = row as any;
        
        // Mapeo según las cabeceras reales del archivo BBVA:
        // Los datos vienen como: { "SEL": "valor", "NO.": "valor", "CUENTA": "valor", ... }
        
        // Si rowObj es un array, usar índices de columna directamente
        let estado = '';
        if (Array.isArray(rowObj)) {
          // Mapeo por posición de columna (más confiable)
          // Columna "Situación" está en la posición 8 (índice 8)
          const situacionIndex = 8;
          estado = String(rowObj[situacionIndex] || '');
        } else {
          // Mapeo por nombre de columna usando includes para 'Situaci'
          // Buscar cualquier clave que contenga 'Situaci'
          for (const key of Object.keys(rowObj)) {
            if (key.toLowerCase().includes('situaci')) {
              estado = String(rowObj[key] || '');
              break;
            }
          }
        }

        const record: AbonoRecord = {
          id: `${data.fileName}_${index}`,
          // Mapeo correcto para BBVA: Sel / No. / Cuenta / Titular(Archivo) / Importe
          beneficiario: getColumnValue(rowObj, 'Titular(Archivo)'), 
          documento_tipo: '', 
          documento: getColumnValue(rowObj, 'Doc.Identidad').replace(/^L\s*-\s*/i, ''), // Doc.Identidad para BBVA, quitando "L - "
          documento_2: '',
          documento_3: '',
          monto_mn: 0,
          monto: parseFloat(getColumnValue(rowObj, 'Importe')) || 0, 
          tc: '',
          monto_abonado: 0,
          monto_abonado_2: 0,
          cuenta_tipo: '',
          cuenta_numero: getColumnValue(rowObj, 'Cuenta').replace(/-/g, ''), 
          cuenta_nombre: '',
          estado: estado, // Situación mapeada por posición o nombre
          observaciones: '', // Vacío por ahora
          banco: 'BBVA',
          origen: data.fileName
        };
        
        // Log detallado de cada registro para debugging
        console.log(`=== BBVA RECORD ${index} ===`);
        console.log(`Datos originales del Excel:`, rowObj);
        console.log(`🔍 DEBUGGING BBVA:`);
        
        if (Array.isArray(rowObj)) {
          console.log(`📋 Mapeo por posición de columna:`);
          console.log(`  - [0] Sel: "${rowObj[0]}"`);
          console.log(`  - [1] No.: "${rowObj[1]}"`);
          console.log(`  - [2] Cuenta: "${rowObj[2]}"`);
          console.log(`  - [3] Titular(Archivo): "${rowObj[3]}"`);
          console.log(`  - [4] Titular(Banco): "${rowObj[4]}"`);
          console.log(`  - [5] Doc.Identidad: "${rowObj[5]}"`);
          console.log(`  - [6] Importe: "${rowObj[6]}"`);
          console.log(`  - [7] Columna 7: "${rowObj[7]}"`);
          console.log(`  - [8] Situación → estado ✓: "${rowObj[8]}"`);
          console.log(`  - [9] Columna 9: "${rowObj[9]}"`);
        } else {
          console.log(`📋 Mapeo por nombre de columna:`);
          console.log(`  - 'Sel': "${rowObj['Sel']}"`);
          console.log(`  - 'No.': "${rowObj['No.']}"`);
          console.log(`  - 'Cuenta': "${rowObj['Cuenta']}"`);
          console.log(`  - 'Titular(Archivo)': "${rowObj['Titular(Archivo)']}"`);
          console.log(`  - 'Doc.Identidad': "${rowObj['Doc.Identidad']}"`);
          console.log(`  - 'Importe': "${rowObj['Importe']}"`);
          
          // Buscar claves que contengan "situaci" para debugging
          const situacionKeys = Object.keys(rowObj).filter(key => 
            key.toLowerCase().includes('situaci')
          );
          console.log(`  - Claves que contienen 'situaci':`, situacionKeys);
          situacionKeys.forEach(key => {
            console.log(`    - "${key}": "${rowObj[key]}" → estado`);
          });
        }
        
        console.log(`📊 RESULTADO MAPEADO:`);
        console.log(`  - beneficiario: "${record.beneficiario}"`);
        console.log(`  - documento: "${record.documento}"`);
        console.log(`  - cuenta_numero: "${record.cuenta_numero}"`);
        console.log(`  - monto: ${record.monto}`);
        console.log(`  - estado: "${record.estado}"`);
        console.log(`  - observaciones: "${record.observaciones}"`);
        console.log(`🔑 TODAS LAS CLAVES:`, Object.keys(rowObj));
        console.log(`========================`);
        
        // Incluir registros que tengan algún dato
        if (record.beneficiario || record.monto > 0 || record.estado || record.cuenta_numero) {
          records.push(record);
        }
      } else if (Array.isArray(row)) {
        // Si viene como array, usar la lógica anterior
        const rowArray = row as any[];
        
        const record: AbonoRecord = {
          id: `${data.fileName}_${index}`,
          // Mapeo correcto para BBVA: SEL / NO. / CUENTA / TITULAR(ARCHIVO) / IMPORTE
          beneficiario: String(rowArray[3] || ''), // Columna D: TITULAR(ARCHIVO)
          documento_tipo: '', 
          documento: String(rowArray[1] || '').replace(/^L\s*-\s*/i, '') || '-', // Columna B: NO., quitando "L - "
          documento_2: '',
          documento_3: '',
          monto_mn: 0,
          monto: parseFloat(String(rowArray[4] || '0')) || 0, // Columna E: IMPORTE
          tc: '',
          monto_abonado: 0,
          monto_abonado_2: 0,
          cuenta_tipo: '',
          cuenta_numero: String(rowArray[2] || '').replace(/-/g, ''), // Columna C: CUENTA
          cuenta_nombre: '',
          estado: String(rowArray[2] || '') || '', // Columna A: SEL
          observaciones: '',
          banco: 'BBVA',
          origen: data.fileName
        };
        
        if (record.beneficiario || record.monto > 0 || record.estado || record.cuenta_numero) {
          records.push(record);
        }
      } else {
        console.log(`BBVA - Fila ${index} formato no reconocido:`, row);
      }
    });
  });
  
  return {
    records,
    totalRecords: records.length,
    sources: [data.fileName],
    processedAt: new Date()
  };
};
