/**
 * Procesador específico para archivos BCP - Sistema de Abonos Taxi Monterrico
 */
import * as ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { ExcelData, ExcelSheet, ExcelRow, AbonoRecord, CombinedData } from '../types/excel';

// Función para crear y descargar Excel modificado con columnas fusionadas
export const createAndDownloadModifiedBCPExcel = async (file: File): Promise<ArrayBuffer> => {
  try {
    console.log('🔄 Creando Excel modificado con columnas fusionadas...');
    
    // Leer el archivo original
    const arrayBuffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    
    if (file.name.toLowerCase().endsWith('.xlsx')) {
      await workbook.xlsx.load(arrayBuffer);
    } else {
      // Para archivos .xls, usar XLSX
      const data = new Uint8Array(arrayBuffer);
      const workbook_xlsx = XLSX.read(data, { type: 'array' });
      const worksheet = workbook_xlsx.Sheets[workbook_xlsx.SheetNames[0]];
      
      // Obtener el rango de datos
      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
      
      // Crear nueva hoja en el workbook principal
      const newWorksheet = workbook.addWorksheet('Sheet1');
      
      // Copiar datos
      for (let row = range.s.r; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
          const cell = worksheet[cellAddress];
          if (cell && cell.v !== undefined) {
            newWorksheet.getCell(row + 1, col + 1).value = cell.v;
          }
        }
      }
    }
    
    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) {
      throw new Error('No se pudo obtener la hoja de trabajo del archivo Excel');
    }
    
    // Obtener la fila de encabezados
    const headerRow = worksheet.getRow(1);
    const headers: string[] = [];
    headerRow.eachCell((cell, colNumber) => {
      headers[colNumber - 1] = String(cell.value || '');
    });
    
    console.log('📋 Encabezados originales:', headers);
    console.log('📊 Total de columnas encontradas:', headers.length);
    console.log('📊 Total de filas en el worksheet:', worksheet.rowCount);
    
    // Identificar columnas duplicadas
    const duplicateColumns: { [key: string]: number[] } = {};
    headers.forEach((header, colIndex) => {
      if (header === 'Documento - Tipo' || header === 'Documento') {
        if (!duplicateColumns[header]) {
          duplicateColumns[header] = [];
        }
        duplicateColumns[header].push(colIndex);
      }
    });
    
    console.log('🔍 Columnas duplicadas encontradas:', duplicateColumns);
    
    // Crear nuevo workbook con columnas fusionadas
    const newWorkbook = new ExcelJS.Workbook();
    const newWorksheet = newWorkbook.addWorksheet('Sheet1');
    
    // Crear nuevos encabezados (sin duplicados)
    const newHeaders: string[] = [];
    const usedHeaders = new Set<string>();
    
    headers.forEach((header) => {
      if (header === 'Documento - Tipo' || header === 'Documento') {
        if (!usedHeaders.has(header)) {
          newHeaders.push(header);
          usedHeaders.add(header);
          console.log(`✅ Agregando encabezado: "${header}"`);
        } else {
          console.log(`⏭️ Saltando encabezado duplicado: "${header}"`);
        }
      } else {
        newHeaders.push(header);
        console.log(`📝 Agregando encabezado normal: "${header}"`);
      }
    });
    
    console.log('📋 Nuevos encabezados:', newHeaders);
    
    // Escribir encabezados
    newHeaders.forEach((header, index) => {
      newWorksheet.getCell(1, index + 1).value = header;
    });
    
    // Procesar cada fila de datos
    const rowCount = worksheet.rowCount || 0;
    console.log(`📊 Procesando ${rowCount - 1} filas de datos...`);
    
    // Verificar que hay datos antes de procesar
    if (rowCount <= 1) {
      console.warn('⚠️ No hay filas de datos para procesar');
      throw new Error('No hay datos para procesar en el archivo');
    }
    
    for (let rowIndex = 2; rowIndex <= rowCount; rowIndex++) {
      const row = worksheet.getRow(rowIndex);
      const newRowData: any[] = [];
      
      // Verificar que la fila existe y tiene datos
      if (!row || !row.hasValues) {
        console.log(`⏭️ Fila ${rowIndex} vacía, saltando...`);
        continue;
      }
      
      // Resetear usedHeaders para cada fila
      const rowUsedHeaders = new Set<string>();
      
      // Mapear columnas originales a nuevas
      let newColIndex = 0;
      headers.forEach((header, originalColIndex) => {
        const cellValue = row.getCell(originalColIndex + 1).value;
        const cellStr = String(cellValue || '').trim();
        
        if (header === 'Documento - Tipo' || header === 'Documento') {
          if (!rowUsedHeaders.has(header)) {
            // Primera aparición de la columna duplicada
            newRowData[newColIndex] = cellValue;
            rowUsedHeaders.add(header);
            console.log(`✅ Fila ${rowIndex}: "${header}" = "${cellStr}" (primera aparición)`);
            newColIndex++;
          } else {
            // Columnas duplicadas subsecuentes - verificar si tienen datos mejores
            const existingValue = newRowData[newColIndex - 1];
            const existingStr = String(existingValue || '').trim();
            
            // Si el valor actual tiene datos y el existente está vacío, reemplazar
            if ((!existingStr || existingStr === '-' || existingStr === '') && 
                cellStr && cellStr !== '-' && cellStr !== '') {
              newRowData[newColIndex - 1] = cellValue;
              console.log(`🔄 Fila ${rowIndex}: Reemplazando "${existingStr}" con "${cellStr}" para ${header}`);
            } else {
              console.log(`⏭️ Fila ${rowIndex}: Manteniendo "${existingStr}" para ${header} (${cellStr})`);
            }
          }
        } else {
          // Columna normal
          newRowData[newColIndex] = cellValue;
          console.log(`📝 Fila ${rowIndex}: "${header}" = "${cellStr}"`);
          newColIndex++;
        }
      });
      
      // Escribir la fila procesada
      newRowData.forEach((value, colIndex) => {
        newWorksheet.getCell(rowIndex, colIndex + 1).value = value;
      });
      
      console.log(`📋 Fila ${rowIndex} procesada:`, newRowData);
    }
    
    // Verificar que el nuevo worksheet tenga datos
    console.log(`🔍 Verificando datos en el nuevo worksheet...`);
    console.log(`Filas en nuevo worksheet: ${newWorksheet.rowCount}`);
    console.log(`Columnas en nuevo worksheet: ${newWorksheet.columnCount}`);
    
    // Mostrar algunas celdas para verificar
    for (let row = 1; row <= Math.min(3, newWorksheet.rowCount || 0); row++) {
      const rowData: any[] = [];
      for (let col = 1; col <= Math.min(5, newWorksheet.columnCount || 0); col++) {
        const cellValue = newWorksheet.getCell(row, col).value;
        rowData.push(cellValue);
      }
      console.log(`Fila ${row}:`, rowData);
    }
    
    // Generar el archivo Excel
    const buffer = await newWorkbook.xlsx.writeBuffer();
    console.log(`📦 Buffer generado: ${buffer.byteLength} bytes`);
    
    // Crear el archivo (sin descargar)
    // COMENTADO: Ya no se descarga automáticamente, solo se retorna el buffer
    // const blob = new Blob([buffer], { 
    //   type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    // });
    
    // const url = window.URL.createObjectURL(blob);
    // const link = document.createElement('a');
    // link.href = url;
    // link.download = `BCP_Modificado_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.xlsx`;
    // document.body.appendChild(link);
    // link.click();
    // document.body.removeChild(link);
    // window.URL.revokeObjectURL(url);
    
    console.log('✅ Excel modificado creado exitosamente (sin descargar)');
    
    // Retornar el buffer para procesamiento posterior
    return buffer;
    
  } catch (error) {
    console.error('❌ Error creando Excel modificado:', error);
    throw new Error('No se pudo crear el archivo Excel modificado');
  }
};

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

export const isValidExcelFile = (file: File): boolean => {
  const validTypes = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel'
  ];
  const validExtensions = ['.xlsx', '.xls'];
  
  return validTypes.includes(file.type) || 
         validExtensions.some(ext => file.name.toLowerCase().endsWith(ext));
};

export const formatFileSize = (bytes: number): string => {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
};

export const processBCPFile = (file: File): Promise<ExcelData> => {
  return new Promise(async (resolve, reject) => {
    // Validar el archivo antes de procesarlo
    if (!validateExcelFile(file)) {
      reject(new Error('El archivo no es un archivo Excel válido (.xlsx o .xls)'));
      return;
    }
    
    // Crear archivo modificado internamente (sin descargar)
    let modifiedFileBuffer: ArrayBuffer | null = null;
    try {
      console.log('🔄 Creando archivo modificado sin columnas duplicadas...');
      modifiedFileBuffer = await createAndDownloadModifiedBCPExcel(file);
      console.log('📥 Archivo modificado creado internamente (sin descargar)');
    } catch (error) {
      console.warn('⚠️ No se pudo crear el Excel modificado:', error);
      modifiedFileBuffer = null;
      // Continuar con el procesamiento normal aunque falle la creación
    }
    
    const reader = new FileReader();
    
    reader.onload = async (e) => {
      try {
        let data = e.target?.result;
        
        // Si tenemos el archivo modificado, usarlo en lugar del original
        if (modifiedFileBuffer) {
          console.log(`=== PROCESANDO ARCHIVO BCP MODIFICADO (SIN DUPLICADOS) ===`);
          data = modifiedFileBuffer;
        } else {
          console.log(`=== PROCESANDO ARCHIVO BCP ORIGINAL ===`);
        }
        
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
            
            workbookXLSX.SheetNames.forEach(sheetName => {
              const worksheet = workbookXLSX.Sheets[sheetName];
              const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1, 
                defval: '', 
                blankrows: true,
                raw: false
              });
              
              if (jsonData.length > 0) {
                const headers = (jsonData[0] as any[])?.map((cell: any) => String(cell || '').trim()) || [];
                const rows: ExcelRow[] = [];
                
                for (let i = 1; i < jsonData.length; i++) {
                  const row = jsonData[i] as any[];
                  if (row && row.length > 0) {
                    const rowData: ExcelRow = {};
                    row.forEach((cell: any, colIndex: number) => {
                      const header = headers[colIndex];
                      if (header) {
                        rowData[header] = String(cell || '').trim();
                      }
                    });
                    rows.push(rowData);
                  }
                }
                
                sheets.push({
                  name: sheetName,
                  headers,
                  data: rows,
                  rowCount: rows.length
                });
                totalRows += rows.length;
              }
            });
            
            const excelData: ExcelData = {
              fileName: file.name,
              sheets,
              totalRows,
              uploadedAt: new Date()
            };
            
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
        
        const sheets: ExcelSheet[] = [];
        let totalRows = 0;
        
        console.log(`Nombres de hojas encontradas:`, workbook.worksheets.map(ws => ws.name));
        
        // Para BCP, usar la hoja principal (la primera que tiene datos)
        const worksheet = workbook.worksheets[0];
        const sheetName = worksheet.name;
        
        console.log(`BCP: Usando hoja principal: ${sheetName}`);
        
        // Convertir la hoja a datos
        const jsonData: any[][] = [];
        
        worksheet.eachRow((row) => {
          const rowData: any[] = [];
          row.eachCell((cell, colNumber) => {
            rowData[colNumber - 1] = cell.value;
          });
          jsonData.push(rowData);
        });
        
        console.log(`BCP: Procesando hoja: ${sheetName} con ${jsonData.length} filas`);
        console.log(`BCP: Primeras 5 filas:`, jsonData.slice(0, 5));
        
        // Log completo del contenido del archivo para debugging
        console.log(`=== CONTENIDO COMPLETO DEL ARCHIVO BCP ===`);
        console.log(`Total de filas: ${jsonData.length}`);
        jsonData.forEach((row, index) => {
          console.log(`Fila ${index + 1}:`, row);
        });
        console.log(`=== FIN DEL CONTENIDO COMPLETO ===`);
        
        if (jsonData.length === 0) {
          throw new Error('No se encontraron datos en el archivo Excel');
        }
        
        // Buscar la fila de headers específica para BCP
        let headerRowIndex = 0;
        let headers: string[] = [];
        
        // Buscar headers en las primeras 10 filas
        for (let i = 0; i < Math.min(10, jsonData.length); i++) {
          const row = jsonData[i];
          if (row && row.length > 0) {
            const potentialHeaders = row.map(cell => String(cell || '').trim());
            
            // Verificar si esta fila contiene las cabeceras específicas de BCP
            const hasBCPHeaders = potentialHeaders.some(header => 
              header && header.length > 0 && 
              (header.toLowerCase().includes('beneficiario') && header.toLowerCase().includes('nombre') ||
               header.toLowerCase().includes('documento') && header.toLowerCase().includes('tipo') ||
               header.toLowerCase().includes('monto') ||
               header.toLowerCase().includes('cuenta') && header.toLowerCase().includes('número') ||
               header.toLowerCase().includes('estado') ||
               header.toLowerCase().includes('observación'))
            );
            
            if (hasBCPHeaders) {
              headerRowIndex = i;
              headers = potentialHeaders;
              console.log(`BCP: Headers específicos encontrados en fila ${i + 1}:`, headers);
              console.log(`BCP: Verificando cabeceras específicas:`);
              headers.forEach((header, idx) => {
                console.log(`  - Columna ${idx}: "${header}"`);
              });
              break;
            }
          }
        }
        
        if (headers.length === 0) {
          // Si no encontramos headers, usar la primera fila
          headers = jsonData[0]?.map(cell => String(cell || '').trim()) || [];
          console.log(`BCP: Usando primera fila como headers:`, headers);
        }
        
        // Procesar datos (empezar después de la fila de headers)
        const dataStartIndex = headerRowIndex + 1;
        const dataEndIndex = jsonData.length;
        
        console.log(`BCP: Procesando datos desde fila ${dataStartIndex + 1} hasta ${dataEndIndex}`);
        
        const rows: ExcelRow[] = [];
        
        for (let i = dataStartIndex; i < dataEndIndex; i++) {
          const row = jsonData[i];
          if (row && Array.isArray(row) && row.length > 0) {
            const rowData: { [key: string]: any } = {};
            
            headers.forEach((header, colIndex) => {
              if (header && header.length > 0) {
                rowData[header] = row[colIndex] || '';
              }
            });
            
            // Solo incluir filas que tengan algún dato
            const hasData = Object.values(rowData).some(value => 
              value !== '' && value !== null && value !== undefined
            );
            
            if (hasData) {
              rows.push(rowData);
            }
          }
        }
        
        console.log(`BCP: Procesadas ${rows.length} filas de datos`);
        
        const sheet: ExcelSheet = {
          name: sheetName,
          headers,
          data: rows,
          rowCount: rows.length
        };
        
        sheets.push(sheet);
        totalRows += rows.length;
        
        const excelData: ExcelData = {
          fileName: file.name,
          sheets,
          totalRows,
          uploadedAt: new Date()
        };
        
        console.log(`BCP: Archivo procesado exitosamente - ${totalRows} filas en ${sheets.length} hojas`);
        resolve(excelData);
        
      } catch (error) {
        console.error('BCP: Error procesando archivo:', error);
        reject(new Error(`Error procesando archivo BCP: ${error instanceof Error ? error.message : 'Error desconocido'}`));
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
  
  console.log(`=== PROCESANDO ARCHIVO BCP: ${data.fileName} ===`);
  console.log(`📥 NOTA: Este archivo ya fue modificado para eliminar columnas duplicadas`);
  
  data.sheets.forEach((sheet, sheetIndex) => {
    console.log(`BCP - Procesando hoja ${sheetIndex}: ${sheet.name}`);
    console.log(`BCP - Headers de la hoja:`, sheet.headers);
    
    // Debug: mostrar las primeras 3 filas de datos para verificar
    console.log(`BCP - Primera fila de datos:`, sheet.data[0]);
    console.log(`BCP - Segunda fila de datos:`, sheet.data[1]);
    console.log(`BCP - Tercera fila de datos:`, sheet.data[2]);
    
    // Debug: verificar si hay datos en las columnas C y D
    console.log(`BCP - Verificando columnas C y D:`);
    console.log(`  - Primera fila completa:`, sheet.data[0]);
    console.log(`  - Columna C (índice 2): "${sheet.data[0]?.[2]}"`);
    console.log(`  - Columna D (índice 3): "${sheet.data[0]?.[3]}"`);
    console.log(`  - Segunda fila completa:`, sheet.data[1]);
    console.log(`  - Columna C (índice 2): "${sheet.data[1]?.[2]}"`);
    console.log(`  - Columna D (índice 3): "${sheet.data[1]?.[3]}"`);
    
    // Debug adicional: mostrar todas las columnas disponibles
    if (sheet.data && sheet.data[0] && Array.isArray(sheet.data[0])) {
      console.log(`BCP - Todas las columnas disponibles:`);
      sheet.data[0].forEach((value: any, index: number) => {
        console.log(`  - Columna ${index}: "${value}"`);
      });
    } else {
      console.log(`BCP - sheet.data[0] no es un array:`, sheet.data[0]);
      console.log(`BCP - Tipo de sheet.data[0]:`, typeof sheet.data[0]);
    }
    
    sheet.data.forEach((row, index) => {
      // Los datos vienen como objetos con las cabeceras como propiedades
      if (typeof row === 'object' && row !== null && !Array.isArray(row)) {
        let rowObj = row as any;
        
        // Renombrar columnas duplicadas para evitar conflictos
        const processedRowObj: any = {};
        const columnCounts: { [key: string]: number } = {};
        
        // Procesar cada columna del objeto
        for (const [key, value] of Object.entries(rowObj)) {
          let processedKey = key;
          
          // Si es una columna duplicada (Documento - Tipo o Documento)
          if (key === 'Documento - Tipo' || key === 'Documento') {
            // Contar cuántas veces hemos visto esta columna
            columnCounts[key] = (columnCounts[key] || 0) + 1;
            
            // Si es la segunda o más aparición, renombrarla agregando 's'
            if (columnCounts[key] > 1) {
              processedKey = key === 'Documento - Tipo' ? 'Documento - Tipos' : 'Documentos';
              console.log(`🔄 Renombrando columna duplicada: "${key}" → "${processedKey}"`);
            }
          }
          
          processedRowObj[processedKey] = value;
        }
        
        // Ahora intercambiar los nombres: las columnas con datos se renombran, las vacías mantienen nombres originales
        const finalRowObj: any = {};
        for (const [key, value] of Object.entries(processedRowObj)) {
          let finalKey = key;
          const valueStr = String(value || '').trim();
          
          // Si la columna tiene datos y es una duplicada, renombrarla
          if ((key === 'Documento - Tipo' || key === 'Documento') && valueStr && valueStr !== '-' && valueStr !== '') {
            finalKey = key === 'Documento - Tipo' ? 'Documento - Tipos' : 'Documentos';
            console.log(`🔄 Intercambiando columna con datos: "${key}" → "${finalKey}"`);
          }
          // Si la columna está vacía y es una renombrada, volver al nombre original
          else if ((key === 'Documento - Tipos' || key === 'Documentos') && (!valueStr || valueStr === '-' || valueStr === '')) {
            finalKey = key === 'Documento - Tipos' ? 'Documento - Tipo' : 'Documento';
            console.log(`🔄 Intercambiando columna vacía: "${key}" → "${finalKey}"`);
          }
          
          finalRowObj[finalKey] = value;
        }
        
        // Usar el objeto final procesado
        rowObj = finalRowObj;
        
        // Mapeo según las cabeceras reales del archivo BCP:
        // Los datos vienen como: { "Beneficiario - Nombre": "valor", "Documento - Tipo": "valor", ... }
        
        // Función para obtener el valor de una columna manejando variaciones
        const getColumnValue = (rowObj: any, columnName: string, fallbackName?: string): string => {
          // Si es un array, usar índices de columna directamente según la imagen
          if (Array.isArray(rowObj)) {
            // Mapeo por posición de columna según la imagen del archivo BCP
            if (columnName.toLowerCase().includes('beneficiario - nombre')) {
              return String(rowObj[1] || ''); // Columna B (índice 1)
            } else if (columnName.toLowerCase().includes('documento - tipo')) {
              return String(rowObj[2] || ''); // Columna C (índice 2)
            } else if (columnName.toLowerCase().includes('documento')) {
              return String(rowObj[3] || ''); // Columna D (índice 3)
            } else if (columnName.toLowerCase().includes('monto - moneda')) {
              return String(rowObj[6] || ''); // Columna G (índice 6)
            } else if (columnName.toLowerCase().includes('monto')) {
              return String(rowObj[7] || ''); // Columna H (índice 7)
            } else if (columnName.toLowerCase().includes('cuenta - n')) {
              return String(rowObj[13] || ''); // Columna N (índice 13)
            } else if (columnName.toLowerCase().includes('estado')) {
              return String(rowObj[14] || ''); // Columna O (índice 14)
            } else if (columnName.toLowerCase().includes('observacion')) {
              return String(rowObj[15] || ''); // Columna P (índice 15)
            }
            return '';
          }
          
          // Si es un objeto, usar la lógica de búsqueda
          // Intentar con el nombre exacto primero
          if (rowObj[columnName] !== undefined) {
            return String(rowObj[columnName] || '');
          }
          
          // Si hay un nombre alternativo, intentarlo
          if (fallbackName && rowObj[fallbackName] !== undefined) {
            return String(rowObj[fallbackName] || '');
          }
          
          // Buscar columnas renombradas (Documento - Tipos, Documentos) solo si las originales están vacías
          if (columnName.toLowerCase().includes('documento - tipo')) {
            const originalValue = String(rowObj['Documento - Tipo'] || '').trim();
            if (originalValue && originalValue !== '-' && originalValue !== '') {
              return originalValue;
            } else if (rowObj['Documento - Tipos'] !== undefined) {
              return String(rowObj['Documento - Tipos'] || '');
            }
          } else if (columnName.toLowerCase().includes('documento')) {
            const originalValue = String(rowObj['Documento'] || '').trim();
            if (originalValue && originalValue !== '-' && originalValue !== '') {
              return originalValue;
            } else if (rowObj['Documentos'] !== undefined) {
              return String(rowObj['Documentos'] || '');
            }
          }
          
          // Buscar en todas las claves que contengan la palabra
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
              const value = String(rowObj[key] || '').trim();
              if (value && value !== '-' && value !== '') {
                return value;
              }
            }
          }
          
          return '';
        };

        const record: AbonoRecord = {
          id: `${data.fileName}_${index}`,
          // Mapeo correcto para BCP según las claves reales del archivo
          beneficiario: getColumnValue(rowObj, 'Beneficiario - Nombre'), 
          documento_tipo: getColumnValue(rowObj, 'Documento - Tipo'), 
          documento: getColumnValue(rowObj, 'Documento'), 
          documento_2: '',
          documento_3: '',
          monto_mn: 0,
          monto: parseFloat(getColumnValue(rowObj, 'Monto')) || 0, 
          tc: '',
          monto_abonado: 0,
          monto_abonado_2: 0,
          cuenta_tipo: '',
          cuenta_numero: getColumnValue(rowObj, 'Cuenta - Número').replace(/-/g, ''), // Usar 'Cuenta - Número' que aparece en las claves
          cuenta_nombre: '',
          estado: getColumnValue(rowObj, 'Estado'), 
          observaciones: getColumnValue(rowObj, 'Observación'), // Usar 'Observación' con tilde
          banco: 'BCP',
          origen: data.fileName
        };
        
        // Log detallado de cada registro para debugging
        console.log(`=== BCP RECORD ${index} ===`);
        console.log(`Datos originales del Excel:`, rowObj);
        console.log(`🔍 DEBUGGING BCP (ARCHIVO MODIFICADO):`);
        console.log(`Tipo de datos: ${Array.isArray(rowObj) ? 'ARRAY' : 'OBJETO'}`);
        
        if (Array.isArray(rowObj)) {
          console.log(`📋 Mapeo por posición de columna (ARRAY) según imagen:`);
          console.log(`  - [0] Columna A (índice): "${rowObj[0]}"`);
          console.log(`  - [1] Beneficiario - Nombre (B) ✓: "${rowObj[1]}"`);
          console.log(`  - [2] Documento - Tipo (C) ✓: "${rowObj[2]}"`);
          console.log(`  - [3] Documento (D) ✓: "${rowObj[3]}"`);
          console.log(`  - [4] Columna E (vacía): "${rowObj[4]}"`);
          console.log(`  - [5] Columna F (vacía): "${rowObj[5]}"`);
          console.log(`  - [6] Monto - Moneda (G): "${rowObj[6]}"`);
          console.log(`  - [7] Monto (H) ✓: "${rowObj[7]}"`);
          console.log(`  - [8] T/C (I): "${rowObj[8]}"`);
          console.log(`  - [9] Monto abo (J): "${rowObj[9]}"`);
          console.log(`  - [10] Monto abo (K): "${rowObj[10]}"`);
          console.log(`  - [11] Cuenta - T (L): "${rowObj[11]}"`);
          console.log(`  - [12] Cuenta - M (M): "${rowObj[12]}"`);
          console.log(`  - [13] Cuenta - N (N) ✓: "${rowObj[13]}"`);
          console.log(`  - [14] Estado (O) ✓: "${rowObj[14]}"`);
          console.log(`  - [15] Observación (P) ✓: "${rowObj[15]}"`);
        } else {
          console.log(`📋 Mapeo por nombre de columna (OBJETO):`);
          console.log(`  - 'Beneficiario - Nombre': "${rowObj['Beneficiario - Nombre']}"`);
          console.log(`  - 'Cuenta - Número': "${rowObj['Cuenta - Número']}"`);
          console.log(`  - 'Monto': "${rowObj['Monto']}"`);
          console.log(`  - 'Estado': "${rowObj['Estado']}"`);
          console.log(`  - 'Observación': "${rowObj['Observación']}"`);
          
          // Debug para columnas duplicadas y renombradas
          const documentoKeys = Object.keys(rowObj).filter(key => key.includes('Documento'));
          console.log(`🔍 COLUMNAS 'Documento' ENCONTRADAS:`, documentoKeys);
          documentoKeys.forEach(key => {
            console.log(`  - "${key}": "${rowObj[key]}"`);
          });
          
          const documentoTipoKeys = Object.keys(rowObj).filter(key => key.includes('Documento - Tipo'));
          console.log(`🔍 COLUMNAS 'Documento - Tipo' ENCONTRADAS:`, documentoTipoKeys);
          documentoTipoKeys.forEach(key => {
            console.log(`  - "${key}": "${rowObj[key]}"`);
          });
          
          // Mostrar columnas renombradas
          if (rowObj['Documento - Tipos'] !== undefined) {
            console.log(`🔄 COLUMNA RENOMBRADA: "Documento - Tipos": "${rowObj['Documento - Tipos']}"`);
          }
          if (rowObj['Documentos'] !== undefined) {
            console.log(`🔄 COLUMNA RENOMBRADA: "Documentos": "${rowObj['Documentos']}"`);
          }
          
          // Mostrar lógica de selección
          console.log(`🎯 LÓGICA DE SELECCIÓN:`);
          const docTipoOriginal = String(rowObj['Documento - Tipo'] || '').trim();
          const docTipoRenombrado = String(rowObj['Documento - Tipos'] || '').trim();
          console.log(`  - Documento - Tipo original: "${docTipoOriginal}" (${docTipoOriginal && docTipoOriginal !== '-' ? 'VÁLIDO' : 'VACÍO'})`);
          console.log(`  - Documento - Tipos renombrado: "${docTipoRenombrado}" (${docTipoRenombrado && docTipoRenombrado !== '-' ? 'VÁLIDO' : 'VACÍO'})`);
          
          const docOriginal = String(rowObj['Documento'] || '').trim();
          const docRenombrado = String(rowObj['Documentos'] || '').trim();
          console.log(`  - Documento original: "${docOriginal}" (${docOriginal && docOriginal !== '-' ? 'VÁLIDO' : 'VACÍO'})`);
          console.log(`  - Documentos renombrado: "${docRenombrado}" (${docRenombrado && docRenombrado !== '-' ? 'VÁLIDO' : 'VACÍO'})`);
        }
        
        console.log(`📊 RESULTADO MAPEADO:`);
        console.log(`  - beneficiario: "${record.beneficiario}"`);
        console.log(`  - documento_tipo: "${record.documento_tipo}"`);
        console.log(`  - documento: "${record.documento}"`);
        console.log(`  - cuenta_numero: "${record.cuenta_numero}"`);
        console.log(`  - monto: ${record.monto}`);
        console.log(`  - estado: "${record.estado}"`);
        console.log(`🔑 TODAS LAS CLAVES:`, Object.keys(rowObj));
        console.log(`========================`);
        
        // Incluir registros que tengan algún dato
        if (record.beneficiario || record.monto > 0 || record.estado || record.cuenta_numero) {
          records.push(record);
        }
      } else if (Array.isArray(row)) {
        // Si viene como array, usar la lógica anterior
        const rowArray = row as any[];
        
        console.log(`=== BCP ARRAY RECORD ${index} ===`);
        console.log(`Datos originales del array:`, rowArray);
        console.log(`📋 Mapeo por posición de columna (ARRAY):`);
        console.log(`  - [0] Columna A: "${rowArray[0]}"`);
        console.log(`  - [1] Columna B: "${rowArray[1]}"`);
        console.log(`  - [2] Documento - Tipo (C) ✓: "${rowArray[2]}"`);
        console.log(`  - [3] Documento (D) ✓: "${rowArray[3]}"`);
        console.log(`  - [4] Columna E: "${rowArray[4]}"`);
        console.log(`  - [5] Columna F: "${rowArray[5]}"`);
        console.log(`  - [6] Monto: "${rowArray[6]}"`);
        console.log(`  - [11] Cuenta: "${rowArray[11]}"`);
        console.log(`  - [12] Estado: "${rowArray[12]}"`);
        
        const record: AbonoRecord = {
          id: `${data.fileName}_${index}`,
          // Mapeo correcto para BCP: Beneficiario - Nombre / Documento / Cuenta - Número / Monto / Estado
          beneficiario: String(rowArray[0] || ''), 
          documento_tipo: String(rowArray[2] || ''), // Columna C: Documento - Tipo
          documento: String(rowArray[3] || '') || '-', // Columna D: Documento
          documento_2: '',
          documento_3: '',
          monto_mn: 0,
          monto: parseFloat(String(rowArray[6] || '0')) || 0, 
          tc: '',
          monto_abonado: 0,
          monto_abonado_2: 0,
          cuenta_tipo: '',
          cuenta_numero: String(rowArray[11] || '').replace(/-/g, ''), 
          cuenta_nombre: '',
          estado: String(rowArray[12] || '') || '', // Columna M: Estado
          observaciones: '',
          banco: 'BCP',
          origen: data.fileName
        };
        
        console.log(`📊 RESULTADO MAPEADO (ARRAY):`);
        console.log(`  - beneficiario: "${record.beneficiario}"`);
        console.log(`  - documento_tipo: "${record.documento_tipo}"`);
        console.log(`  - documento: "${record.documento}"`);
        console.log(`  - cuenta_numero: "${record.cuenta_numero}"`);
        console.log(`  - monto: ${record.monto}`);
        console.log(`  - estado: "${record.estado}"`);
        console.log(`========================`);
        
        if (record.beneficiario || record.monto > 0 || record.estado || record.cuenta_numero) {
          records.push(record);
        }
      } else {
        console.log(`BCP - Fila ${index} formato no reconocido:`, row);
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

export const exportToCSV = (data: any[], filename: string): void => {
  if (!data || data.length === 0) {
    console.warn('No hay datos para exportar');
    return;
  }
  
  const headers = Object.keys(data[0]);
  const csvContent = [
    headers.join(','),
    ...data.map(row => 
      headers.map(header => {
        const value = row[header];
        // Escapar comillas y envolver en comillas si contiene comas
        const stringValue = String(value || '');
        if (stringValue.includes(',') || stringValue.includes('"') || stringValue.includes('\n')) {
          return `"${stringValue.replace(/"/g, '""')}"`;
        }
        return stringValue;
      }).join(',')
    )
  ].join('\n');
  
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);
  link.setAttribute('href', url);
  link.setAttribute('download', `${filename}.csv`);
  link.style.visibility = 'hidden';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};
