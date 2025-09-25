/**
 * Utilidades para procesar archivos Excel - Sistema de Abonos Taxi Monterrico
 */
import * as XLSX from 'xlsx';
import * as ExcelJS from 'exceljs';
import { ExcelData, ExcelSheet, ExcelRow, AbonoRecord, CombinedData } from '../types/excel';

// Mapeo inteligente de campos comunes para abonos
const fieldMappings: { [key: string]: string[] } = {
  beneficiario: [
    'titular(archivo)', 'titular(banco)', 'beneficiario - nombre', 'beneficiario', 'titular',
    'cliente', 'client', 'nombre', 'name', 'cliente_nombre', 'pasajero'
  ],
  documento_tipo: [
    'documento - tipo', 'documento - tipo documento', 'tipo documento', 'doc tipo'
  ],
  documento: [
    'documento', 'doc.identidad', 'documento - documento', 'numero documento'
  ],
  documento_2: [
    'documento 2', 'documento - documento 2', 'segundo documento'
  ],
  documento_3: [
    'documento 3', 'documento - documento 3', 'tercer documento'
  ],
  monto_mn: [
    'monto - m/n', 'monto mn', 'monto moneda nacional', 'monto soles'
  ],
  monto: [
    'monto', 'amount', 'importe', 'valor', 'precio', 'total', 'suma',
    'monto - monto', 'importe cargado por abonos', 'importe situación'
  ],
  tc: [
    't/c', 'tipo cambio', 'tc', 'cambio'
  ],
  monto_abonado: [
    'monto abonado', 'monto - abonado', 'abonado'
  ],
  monto_abonado_2: [
    'monto abonado 2', 'monto - abonado 2', 'segundo abonado'
  ],
  cuenta_tipo: [
    'cuenta - t', 'cuenta tipo', 'tipo cuenta'
  ],
  cuenta_numero: [
    'cuenta - n', 'cuenta numero', 'numero cuenta', 'cuenta - cuenta', 'cuenta', 'no.'
  ],
  cuenta_nombre: [
    'cuenta - nombre', 'nombre cuenta'
  ],
  estado: [
    'estado', 'status', 'situacion', 'condicion', 'situación', 'situación de proceso'
  ],
  observaciones: [
    'observaciones', 'observacion', 'comentarios', 'notas', 'obs'
  ],
  banco: [
    'banco', 'entidad', 'bank', 'institución', 'institucion'
  ]
};

const findBestMatch = (headers: string[], targetField: string, isBBVAFile: boolean = false, isBCPFile: boolean = false): string | null => {
  const possibleMatches = fieldMappings[targetField] || [];
  
  // Mapeo específico para archivos BBVA basado en la estructura del reporte
  if (isBBVAFile) {
    const bbvaMappings: { [key: string]: string[] } = {
      beneficiario: ['titular', 'titular(archivo)', 'titular(banco)'],
      documento: ['doc.identidad', 'doc identidad', 'documento'],
      monto: ['importe'],
      monto_abonado: ['importe'],
      estado: ['situación', 'situacion', 'situ', 'estado'],
      cuenta_numero: ['cuenta'],
      banco: ['banco']
    };
    
    const bbvaMatches = bbvaMappings[targetField] || [];
    for (const match of bbvaMatches) {
      const found = headers.find(header => 
        header?.toLowerCase().includes(match.toLowerCase()) ||
        match.toLowerCase().includes(header?.toLowerCase() || '')
      );
      if (found) return found;
    }
  } else if (isBCPFile) {
    // Mapeo específico para archivos BCP - basado en la estructura real del Excel
    const bcpMappings: { [key: string]: string[] } = {
      beneficiario: ['beneficiario - nombre', 'beneficiario', 'cliente', 'nombre', 'titular'],
      documento_tipo: ['documento - tipo', 'documento-tipo', 'tipo documento', 'doc tipo', 'tipo', 'documento tipo', 'documentotipo'],
      documento: ['documento', 'numero documento', 'doc', 'numero', 'documento - número', 'documento-número', 'documentonúmero'],
      monto_mn: ['monto - moneda', 'moneda', 'monto moneda'],
      monto: ['monto', 'importe', 'amount', 'monto - monto'],
      tc: ['t/c', 'tipo cambio', 'tc'],
      monto_abonado: ['monto abonado - moneda', 'monto abonado', 'abonado'],
      cuenta_tipo: ['cuenta - t', 'cuenta tipo', 'tipo cuenta'],
      cuenta_numero: ['cuenta - número', 'cuenta numero', 'numero cuenta', 'cuenta', 'cuenta - n'],
      estado: ['estado', 'status', 'situación', 'situacion'],
      observaciones: ['observación', 'observacion', 'obs'],
      banco: ['banco', 'entidad', 'institución', 'institucion']
    };
    
    const bcpMatches = bcpMappings[targetField] || [];
    
    // Para documento_tipo (se mapea a columna "Documento"), buscar "documento - tipo"
    if (targetField === 'documento_tipo') {
      const found = headers.find((header) => {
        const headerLower = header?.toLowerCase() || '';
        return headerLower.includes('documento') && headerLower.includes('tipo') && !headerLower.includes('dcumento');
      });
      if (found) return found;
    }
    
    // Para documento (se mapea a columna "# Documento"), buscar "documento" que NO contenga "tipo"
    if (targetField === 'documento') {
      const found = headers.find((header) => {
        const headerLower = header?.toLowerCase() || '';
        return headerLower === 'documento' || (headerLower.includes('documento') && !headerLower.includes('tipo'));
      });
      if (found) return found;
    }
    
    for (const match of bcpMatches) {
      const found = headers.find(header => {
        const headerLower = header?.toLowerCase() || '';
        const matchLower = match.toLowerCase();
        
        // Limpiar espacios, guiones y caracteres especiales para comparación
        const headerClean = headerLower.replace(/[-\s_]/g, '');
        const matchClean = matchLower.replace(/[-\s_]/g, '');
        
        // Múltiples formas de búsqueda
        return headerClean.includes(matchClean) || 
               matchClean.includes(headerClean) ||
               headerLower.includes(matchLower) ||
               matchLower.includes(headerLower) ||
               // Búsqueda por palabras separadas
               matchLower.split(' ').every(word => headerLower.includes(word)) ||
               headerLower.split(' ').every(word => matchLower.includes(word)) ||
               // Búsqueda exacta sin espacios ni guiones
               headerClean === matchClean;
      });
      if (found) {
        console.log(`BCP: Campo '${targetField}' mapeado a '${found}' usando patrón '${match}'`);
        return found;
      }
    }
  }
  
  for (const match of possibleMatches) {
    const found = headers.find(header => 
      header?.toLowerCase().includes(match.toLowerCase()) ||
      match.toLowerCase().includes(header?.toLowerCase() || '')
    );
    if (found) return found;
  }
  
  return null;
};

export const processExcelFile = (file: File, bankType?: 'BBVA' | 'BCP'): Promise<ExcelData> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'array' });
        const sheets: ExcelSheet[] = [];
        let totalRows = 0;
        
        console.log(`=== PROCESANDO ARCHIVO BBVA ===`);
        console.log(`Nombres de hojas encontradas:`, workbook.SheetNames);
        
        // Para BBVA, usar la hoja principal (la primera que tiene datos)
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        console.log(`BBVA: Usando hoja principal: ${sheetName}`);
        
        // Leer datos con configuración más robusta
        let jsonData = XLSX.utils.sheet_to_json(worksheet, { 
          header: 1, 
          defval: '', 
          blankrows: true,
          range: undefined // Leer toda la hoja
        });
        console.log(`BBVA: Datos con header:1 - ${jsonData.length} filas`);
        
        // Si no hay datos, intentar sin header
        if (jsonData.length === 0) {
          jsonData = XLSX.utils.sheet_to_json(worksheet, { 
            header: 0, 
            defval: '', 
            blankrows: true,
            range: undefined
          });
          console.log(`BBVA: Datos sin header - ${jsonData.length} filas`);
        }
        
        // Si aún no hay datos, usar sheet_to_array
        if (jsonData.length === 0) {
          jsonData = XLSX.utils.sheet_to_array(worksheet, { 
            defval: '', 
            blankrows: true 
          });
          console.log(`BBVA: Datos con sheet_to_array - ${jsonData.length} filas`);
        }
        
        // Si aún no hay datos, intentar leer con rango específico
        if (jsonData.length === 0) {
          const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z1000');
          console.log(`BBVA: Rango detectado: ${worksheet['!ref']}`);
          jsonData = XLSX.utils.sheet_to_json(worksheet, { 
            header: 1, 
            defval: '', 
            blankrows: true,
            range: range
          });
          console.log(`BBVA: Datos con rango específico - ${jsonData.length} filas`);
        }
        
        console.log(`BBVA: Procesando hoja: ${sheetName} con ${jsonData.length} filas`);
        console.log(`BBVA: Rango de la hoja: ${worksheet['!ref']}`);
        console.log(`BBVA: Primeras 10 filas:`, jsonData.slice(0, 10));
        console.log(`BBVA: Filas 30-35:`, jsonData.slice(30, 35));
        
        if (jsonData.length === 0) {
          console.log(`BBVA: Saltando hoja ${sheetName} porque está vacía`);
          return;
        }
        
        // Usar el tipo de banco especificado por el usuario
        const isBBVAFile = bankType === 'BBVA';
        const isBCPFile = bankType === 'BCP';
        
        console.log(`=== TIPO DE ARCHIVO ESPECIFICADO ===`);
        console.log(`Tipo de banco: ${bankType}`);
        console.log(`Es BBVA: ${isBBVAFile}`);
        console.log(`Es BCP: ${isBCPFile}`);
        
        let headers: string[];
        let dataStartIndex: number;
        let dataEndIndex = jsonData.length;
        
        if (isBBVAFile) {
          console.log('Archivo BBVA detectado - buscando headers en fila 31');
          
          // Buscar la fila 31 (índice 30) para verificar si tiene los headers correctos
          const row31 = jsonData[30]; // Fila 31 (índice 30)
          console.log(`BBVA: Fila 31 (índice 30):`, row31);
          
          if (row31 && row31.length > 0) {
            const row31Text = row31.map(cell => String(cell || '').toLowerCase()).join(' ');
            console.log(`BBVA: Texto de la fila 31: "${row31Text}"`);
            
            // Verificar si la fila 31 tiene los headers de BBVA
            const hasSel = row31Text.includes('sel');
            const hasNo = row31Text.includes('no.') || row31Text.includes('no');
            const hasCuenta = row31Text.includes('cuenta');
            const hasTitularArchivo = row31Text.includes('titular(archivo)');
            const hasImporte = row31Text.includes('importe');
            
            if (hasSel && hasNo && hasCuenta && hasTitularArchivo && hasImporte) {
              console.log('✓ BBVA: Headers encontrados en fila 31');
              headers = row31.map(cell => String(cell || '').trim());
              dataStartIndex = 31; // Datos empiezan en fila 32 (índice 31)
              console.log(`BBVA: Headers detectados:`, headers);
            } else {
              console.log('❌ BBVA: No se encontraron headers en fila 31');
              return;
            }
          } else {
            console.log('❌ BBVA: Fila 31 está vacía');
            return;
          }
                      
                      // Buscar "Estimado Cliente:" para terminar (como especificaste)
                      // Empezar desde la fila 40 para evitar encontrar texto temprano
                      for (let i = Math.max(dataStartIndex + 10, 40); i < jsonData.length; i++) {
                        const row = jsonData[i] as any[];
                        if (row && row.length > 0) {
                          const rowText = row.map(cell => String(cell || '').toLowerCase()).join(' ');
                          if (rowText.includes('estimado cliente')) {
                            dataEndIndex = i;
                            console.log(`BBVA: Fin de datos encontrado en fila ${i + 1} - texto: "${rowText.substring(0, 50)}..."`);
                            break;
                          }
                        }
                      }
                      
                      // Si no se encontró "Estimado Cliente", usar un rango más amplio
                      if (dataEndIndex === jsonData.length || dataEndIndex <= dataStartIndex + 5) {
                        dataEndIndex = Math.min(dataStartIndex + 50, jsonData.length); // Usar 50 filas desde el inicio
                        console.log(`BBVA: No se encontró "Estimado Cliente", usando rango hasta fila ${dataEndIndex + 1}`);
                      }
                      
                      console.log(`BBVA: Rango final - inicio: ${dataStartIndex + 1}, fin: ${dataEndIndex + 1}, total filas: ${dataEndIndex - dataStartIndex}`);
                      
                      // Asegurar que dataEndIndex sea válido
                      if (dataEndIndex <= dataStartIndex) {
                        dataEndIndex = Math.min(dataStartIndex + 100, jsonData.length); // Aumentar rango a 100 filas
                        console.log(`BBVA: Corrigiendo dataEndIndex a ${dataEndIndex + 1}`);
                      }
                    } else if (isBCPFile) {
                      // Para archivos BCP: buscar headers con lógica mejorada
                      console.log('Archivo BCP detectado - buscando estructura de datos');
                      let headerRowIndex = -1;
                      
                      for (let i = 0; i < Math.min(20, jsonData.length); i++) {
                        const row = jsonData[i] as any[];
                        if (row && row.length > 5) {
                          const rowText = row.map(cell => String(cell || '').toLowerCase()).join(' ');
                          console.log(`BCP: Revisando fila ${i + 1}: "${rowText.substring(0, 100)}..."`);
                          
                          if (rowText.includes('beneficiario') || rowText.includes('titular') || 
                              rowText.includes('cuenta') || rowText.includes('monto') ||
                              rowText.includes('documento') || rowText.includes('importe') ||
                              rowText.includes('cliente') || rowText.includes('nombre')) {
                            headerRowIndex = i;
                            console.log(`BCP: Headers encontrados en fila ${i + 1}`);
                            break;
                          }
                        }
                      }
                      
                      if (headerRowIndex === -1) {
                        headerRowIndex = 0;
                        console.log('BCP: No se encontraron headers específicos, usando fila 1');
                      }
                      
                      headers = jsonData[headerRowIndex] as string[];
                      dataStartIndex = headerRowIndex + 1;
                      console.log(`BCP: Headers encontrados:`, headers);
                    } else {
                      // Para otros archivos: usar lógica genérica
                      console.log('Archivo genérico detectado - buscando estructura de datos');
                      let headerRowIndex = -1;
                      
                      for (let i = 0; i < Math.min(20, jsonData.length); i++) {
                        const row = jsonData[i] as any[];
                        if (row && row.length > 5) {
                          const rowText = row.map(cell => String(cell || '').toLowerCase()).join(' ');
                          
                          if (rowText.includes('beneficiario') || rowText.includes('titular') || 
                              rowText.includes('cuenta') || rowText.includes('monto') ||
                              rowText.includes('documento') || rowText.includes('importe') ||
                              rowText.includes('cliente') || rowText.includes('nombre')) {
                            headerRowIndex = i;
                            console.log(`Genérico: Headers encontrados en fila ${i + 1}`);
                            break;
                          }
                        }
                      }
                      
                      if (headerRowIndex === -1) {
                        headerRowIndex = 0;
                        console.log('Genérico: No se encontraron headers específicos, usando fila 1');
                      }
                      
                      headers = jsonData[headerRowIndex] as string[];
                      dataStartIndex = headerRowIndex + 1;
                      console.log(`Genérico: Headers encontrados:`, headers);
                    }
                    
                    // Limpiar headers
                    const cleanHeaders = headers.map((header, index) => {
                      const headerStr = String(header || '').trim();
                      if (!headerStr || headerStr === '') {
                        return `Columna_${index + 1}`;
                      }
                      return headerStr;
                    });
                    
                    console.log(`=== HEADERS ORIGINALES ===`);
                    console.log('Headers del Excel:', cleanHeaders);
                    console.log(`========================`);
                    
                    console.log(`Procesando ${sheetName}: dataStart=${dataStartIndex}, dataEnd=${dataEndIndex}, isBBVA=${isBBVAFile}`);
                    console.log(`Headers BBVA:`, cleanHeaders);
                    
                    const relevantData = jsonData.slice(dataStartIndex, dataEndIndex);
                    console.log(`BBVA: Datos relevantes: ${relevantData.length} filas encontradas`);
                    console.log(`BBVA: Rango de datos: desde fila ${dataStartIndex + 1} hasta fila ${dataEndIndex + 1}`);
                    
                    // Debug: mostrar las primeras 5 filas de datos relevantes
                    console.log('BBVA: Primeras 5 filas de datos relevantes:');
                    relevantData.slice(0, 5).forEach((row, idx) => {
                      console.log(`  Fila ${dataStartIndex + idx + 1}:`, row);
                    });
                    
                    const dataRows = relevantData
                      .filter(row => {
                        const rowArray = row as any[];
                        if (!rowArray || rowArray.length === 0) return false;
                        
                        const firstCell = String(rowArray[0] || '').trim();
                        
                        if (isBBVAFile) {
                          // Para BBVA: ser más permisivo con el filtrado
                          // Verificar que tenga datos en columnas importantes según la tabla BBVA
                          // Headers: ['Sel', 'No.', 'Cuenta', 'Banco', 'Titular(Archivo)', 'Titular(Banco)', 'Doc.Identidad', 'Importe', 'Situación']
                          const hasCuenta = rowArray[2] && String(rowArray[2] || '').trim() !== '';
                          const hasTitular = rowArray[4] && String(rowArray[4] || '').trim() !== '';
                          const hasImporte = rowArray[7] && String(rowArray[7] || '').trim() !== '';
                          const hasEstado = rowArray[8] && String(rowArray[8] || '').trim() !== '';
                          
                          // Incluir si tiene al menos uno de estos campos importantes
                          const hasImportantData = hasCuenta || hasTitular || hasImporte || hasEstado;
                          
                          // También incluir si la primera celda es un número (1, 2, 3, etc.)
                          const isNumber = /^\d+$/.test(firstCell);
                          
                          // Incluir si tiene datos importantes O si empieza con número
                          return hasImportantData || isNumber;
                        } else if (isBCPFile) {
                          // Para BCP: incluir TODAS las filas que tengan datos en cualquier columna
                          // Verificar que no sea una fila completamente vacía
                          const hasAnyData = rowArray.some(cell => String(cell || '').trim() !== '');
                          
                          // Incluir filas que tengan al menos algún dato
                          return hasAnyData;
                        } else {
                          // Para otros archivos: verificar contenido general
                          return rowArray.some(cell => String(cell || '').trim() !== '');
                        }
                      })
                      .map(row => {
                        const rowObj: ExcelRow = {};
                        cleanHeaders.forEach((header, index) => {
                          const cellValue = (row as any[])[index] || '';
                          rowObj[header] = String(cellValue).trim();
                        });
                        return rowObj;
                      });
            
            console.log(`BBVA: ${dataRows.length} filas procesadas después del filtrado`);
            if (isBBVAFile && dataRows.length === 0) {
              console.log('BBVA: ADVERTENCIA - No se encontraron filas válidas. Revisar filtros.');
              console.log('BBVA: Primeras 3 filas relevantes:', relevantData.slice(0, 3));
              console.log('BBVA: Headers mapeados:', cleanHeaders);
              
              // Debug adicional: mostrar todas las filas relevantes para ver qué está pasando
              console.log('BBVA: TODAS las filas relevantes:');
              relevantData.forEach((row, idx) => {
                console.log(`  Fila ${dataStartIndex + idx + 1}:`, row);
              });
            } else if (isBBVAFile && dataRows.length > 0) {
              console.log('BBVA: ÉXITO - Se encontraron filas válidas');
              console.log('BBVA: Primera fila procesada:', dataRows[0]);
            }
            
            // Mostrar mapeo de campos para debugging
            if (isBBVAFile && dataRows.length > 0) {
              console.log('BBVA: Mapeo de campos exitoso');
              console.log('Primera fila procesada:', dataRows[0]);
              
              // Debug del mapeo de campos específicos
              const mappedFields = {
                beneficiario: findBestMatch(cleanHeaders, 'beneficiario', true, false),
                documento: findBestMatch(cleanHeaders, 'documento', true, false),
                monto: findBestMatch(cleanHeaders, 'monto', true, false),
                estado: findBestMatch(cleanHeaders, 'estado', true, false),
                cuenta_numero: findBestMatch(cleanHeaders, 'cuenta_numero', true, false),
                banco: findBestMatch(cleanHeaders, 'banco', true, false)
              };
              console.log('BBVA: Campos mapeados:', mappedFields);
              console.log('BBVA: Headers disponibles:', cleanHeaders);
              
              // Debug específico para estado y banco
              if (mappedFields.estado) {
                console.log(`BBVA: Campo estado mapeado a: "${mappedFields.estado}"`);
                // Mostrar el valor real del campo
                const estadoValue = dataRows[0][mappedFields.estado];
                console.log(`BBVA: Valor del estado en primera fila: "${estadoValue}"`);
              } else {
                console.log('BBVA: ADVERTENCIA - Campo estado NO mapeado');
              }
              
              if (mappedFields.banco) {
                console.log(`BBVA: Campo banco mapeado a: "${mappedFields.banco}"`);
              } else {
                console.log('BBVA: Campo banco NO mapeado (usará BBVA por defecto)');
              }
              
              // Debug del nombre del archivo
              console.log(`BBVA: Nombre del archivo: "${file.name}"`);
              console.log(`BBVA: ¿Contiene 'bbva'?: ${file.name?.toLowerCase().includes('bbva')}`);
            }
            
            if (isBCPFile && dataRows.length > 0) {
              console.log('BCP: Mapeo de campos exitoso');
              console.log('Primera fila procesada:', dataRows[0]);
              console.log('BCP: Headers disponibles:', cleanHeaders);
              
              // Debug del mapeo de campos específicos para BCP
              const mappedFields = {
                beneficiario: findBestMatch(cleanHeaders, 'beneficiario', false, true),
                documento_tipo: findBestMatch(cleanHeaders, 'documento_tipo', false, true),
                documento: findBestMatch(cleanHeaders, 'documento', false, true),
                monto: findBestMatch(cleanHeaders, 'monto', false, true),
                estado: findBestMatch(cleanHeaders, 'estado', false, true),
                cuenta_numero: findBestMatch(cleanHeaders, 'cuenta_numero', false, true),
                banco: findBestMatch(cleanHeaders, 'banco', false, true)
              };
              console.log('BCP: Campos mapeados:', mappedFields);
              
              // Debug específico para documento_tipo y documento
              if (mappedFields.documento_tipo) {
                console.log(`BCP: Campo documento_tipo mapeado a: "${mappedFields.documento_tipo}"`);
                const docTipoValue = dataRows[0][mappedFields.documento_tipo];
                console.log(`BCP: Valor del documento_tipo en primera fila: "${docTipoValue}"`);
              } else {
                console.log('BCP: ADVERTENCIA - Campo documento_tipo NO mapeado');
              }
              
              if (mappedFields.documento) {
                console.log(`BCP: Campo documento mapeado a: "${mappedFields.documento}"`);
                const docValue = dataRows[0][mappedFields.documento];
                console.log(`BCP: Valor del documento en primera fila: "${docValue}"`);
              } else {
                console.log('BCP: ADVERTENCIA - Campo documento NO mapeado');
              }
              
              // Debug del nombre del archivo
              console.log(`BCP: Nombre del archivo: "${file.name}"`);
              console.log(`BCP: ¿Contiene 'bcp'?: ${file.name?.toLowerCase().includes('bcp')}`);
            }
            
            // Headers fijos para la tabla - adaptados para BBVA y BCP
            let fixedHeaders: string[];
            
            if (isBBVAFile) {
              // Headers específicos para BBVA según la tabla "Relación de las cuentas de abono"
              fixedHeaders = [
                'Item',                    // Número de registro (No.)
                'Beneficiario',           // Titular(Archivo)
                'Documento',              // vacío
                '# Documento',            // vacío
                'Monto',                  // Importe
                'Cuenta',                 // Cuenta
                'Estado',                 // Situación
                'Observación',            // vacío
                'Banco'                   // BBVA
              ];
            } else {
              // Headers para BCP - basados en las columnas resaltadas en amarillo
              fixedHeaders = [
                'Beneficiario - Nombre',    // Columna B - resaltada en amarillo
                'Documento - Tipo',         // Columna C - resaltada en amarillo  
                'Documento',                // Columna D - resaltada en amarillo
                'Monto - Moneda',           // Columna G - resaltada en amarillo
                'Monto',                    // Columna H - resaltada en amarillo
                'Cuenta - T',               // Columna L - no resaltada
                'Cuenta - Número',          // Columna N - resaltada en amarillo
                'Estado',                   // Columna O - resaltada en amarillo
                'Observación'               // Columna P - resaltada en amarillo
              ];
            }
            
          sheets.push({
            name: sheetName,
            data: dataRows,
            headers: fixedHeaders,
            rowCount: dataRows.length
          });
          
          totalRows += dataRows.length;
          
          const excelData: ExcelData = {
            fileName: file.name,
            sheets,
            totalRows,
            uploadedAt: new Date()
          };
          
          resolve(excelData);
      } catch (error) {
        let errorMessage = 'Error desconocido';
        if (error instanceof Error) {
          if (error.message.includes('Bad uncompressed size') || 
              error.message.includes('uncompressed size') || 
              error.message.includes('63295 != 0')) {
            errorMessage = 'El archivo Excel está corrupto o no es un archivo Excel válido. Por favor, verifica el archivo.';
          } else {
            errorMessage = error.message;
          }
        }
        reject(new Error(`Error procesando archivo: ${errorMessage}`));
      }
    };
    
    reader.onerror = () => {
      reject(new Error('Error leyendo el archivo'));
    };
    
    reader.readAsArrayBuffer(file);
  });
};

export const combineExcelData = (data1: ExcelData, data2: ExcelData, bankType1?: 'BBVA' | 'BCP', bankType2?: 'BBVA' | 'BCP'): CombinedData => {
  const combinedRecords: AbonoRecord[] = [];
  
  // Procesar primer archivo
  data1.sheets.forEach(sheet => {
    const isBBVA1 = bankType1 === 'BBVA' || data1.fileName?.toLowerCase().includes('bbva') || false;
    
    sheet.data.forEach((row, index) => {
      const rowArray = row as any[];
      let record: AbonoRecord;
      
      if (isBBVA1) {
        // MAPEO ESPECÍFICO PARA BBVA según la tabla "Relación de las cuentas de abono"
        // Headers: ['Sel', 'No.', 'Cuenta', 'Banco', 'Titular(Archivo)', 'Titular(Banco)', 'Doc.Identidad', 'Importe', 'Situación']
        record = {
          id: `${data1.fileName}_${index}`,
          // Mapeo según especificaciones:
          // item = columna 2 (No.)
          // beneficiario = columna 5 (Titular(Archivo))
          // documento = vacío
          // # documento = vacío
          // monto = columna 8 (Importe)
          // cuenta = columna 3 (Cuenta)
          // estado = columna 9 (Situación)
          // observacion = vacío
          // banco = BBVA
          beneficiario: String(rowArray[4] || ''), // Columna 5: Titular(Archivo)
          documento_tipo: '', // vacío
          documento: '', // vacío
          documento_2: '',
          documento_3: '',
          monto_mn: 0,
          monto: parseFloat(String(rowArray[7] || '0')) || 0, // Columna 8: Importe
          tc: '',
          monto_abonado: 0,
          monto_abonado_2: 0,
          cuenta_tipo: '',
          cuenta_numero: String(rowArray[2] || '').replace(/-/g, ''), // Columna 3: Cuenta
          cuenta_nombre: '',
          estado: String(rowArray[8] || ''), // Columna 9: Situación
          observaciones: '', // vacío
          banco: 'BBVA', // Siempre BBVA para archivos BBVA
          origen: data1.fileName
        };
      } else {
        // MAPEO PARA BCP usando headers fijos
        record = {
          id: `${data1.fileName}_${index}`,
          beneficiario: String(row['Beneficiario - Nombre'] || ''), // Columna B
          documento_tipo: String(row['Documento - Tipo'] || '') || '-', // Columna C
          documento: String(row['Documento'] || '') || '-', // Columna D
          documento_2: '',
          documento_3: '',
          monto_mn: 0,
          monto: parseFloat(String(row['Monto'] || '0')) || 0, // Columna H
          tc: '',
          monto_abonado: 0,
          monto_abonado_2: 0,
          cuenta_tipo: '',
          cuenta_numero: String(row['Cuenta - Número'] || '').replace(/-/g, ''), // Columna N
          cuenta_nombre: '',
          estado: String(row['Estado'] || ''), // Columna O
          observaciones: String(row['Observación'] || ''), // Columna P
          banco: bankType1 || (data1.fileName?.toLowerCase().includes('bbva') ? 'BBVA' : 'BCP'),
          origen: data1.fileName
        };
      }
      
      // Debug: mostrar los valores de documento
      console.log(`${isBBVA1 ? 'BBVA' : 'BCP'} Record ${index}: documento_tipo = "${record.documento_tipo}", documento = "${record.documento}"`);
      console.log(`${isBBVA1 ? 'BBVA' : 'BCP'} Row data:`, row);
      
      // Incluir registros que tengan algún dato
      if (record.beneficiario || record.monto > 0 || record.estado || record.cuenta_numero) {
        combinedRecords.push(record);
      }
    });
  });
  
  // Procesar segundo archivo
  data2.sheets.forEach(sheet => {
    const isBBVA2 = bankType2 === 'BBVA' || data2.fileName?.toLowerCase().includes('bbva') || false;
    const isBCP2 = bankType2 === 'BCP' || data2.fileName?.toLowerCase().includes('bcp') || false;
    
    // Debug para BCP - mostrar todos los headers disponibles
    if (isBCP2) {
      console.log(`=== DEBUGGING BCP HEADERS ===`);
      console.log(`Headers disponibles en BCP:`, sheet.headers);
      console.log(`Headers en lowercase:`, sheet.headers.map(h => h?.toLowerCase()));
      console.log(`=============================`);
    }
    
    // Mapeo directo usando los headers fijos
    const mappedFields = {
      beneficiario: 'Beneficiario - Nombre',
      documento_tipo: 'Documento - Tipo',
      documento: 'Documento',
      documento_2: null,
      documento_3: null,
      monto_mn: 'Monto - Moneda',
      monto: 'Monto',
      tc: null,
      monto_abonado: null,
      monto_abonado_2: null,
      cuenta_tipo: 'Cuenta - T',
      cuenta_numero: 'Cuenta - Número',
      cuenta_nombre: null,
      estado: 'Estado',
      observaciones: 'Observación',
      banco: null
    };
    
    // Debug específico para BCP - mostrar qué campos se mapearon
    if (isBCP2) {
      console.log(`=== DEBUGGING BCP MAPPING ===`);
      console.log(`documento_tipo mapeado a: "${mappedFields.documento_tipo}"`);
      console.log(`documento mapeado a: "${mappedFields.documento}"`);
      console.log(`beneficiario mapeado a: "${mappedFields.beneficiario}"`);
      
      // Debug específico para ver si encuentra los campos exactos
      const docTipoHeader = sheet.headers.find(h => h?.toLowerCase().includes('documento') && h?.toLowerCase().includes('tipo'));
      const docHeader = sheet.headers.find(h => h?.toLowerCase().includes('documento') && !h?.toLowerCase().includes('tipo'));
      console.log(`Header que contiene 'documento' y 'tipo': "${docTipoHeader}"`);
      console.log(`Header que contiene 'documento' pero no 'tipo': "${docHeader}"`);
      console.log(`=============================`);
    }
    
    sheet.data.forEach((row, index) => {
      const rowArray = row as any[];
      let record: AbonoRecord;
      
      if (isBBVA2) {
        // MAPEO ESPECÍFICO PARA BBVA según la tabla "Relación de las cuentas de abono"
        // Headers: ['Sel', 'No.', 'Cuenta', 'Banco', 'Titular(Archivo)', 'Titular(Banco)', 'Doc.Identidad', 'Importe', 'Situación']
        record = {
          id: `${data2.fileName}_${index}`,
          // Mapeo según especificaciones:
          // item = columna 2 (No.)
          // beneficiario = columna 5 (Titular(Archivo))
          // documento = vacío
          // # documento = vacío
          // monto = columna 8 (Importe)
          // cuenta = columna 3 (Cuenta)
          // estado = columna 9 (Situación)
          // observacion = vacío
          // banco = BBVA
          beneficiario: String(rowArray[4] || ''), // Columna 5: Titular(Archivo)
          documento_tipo: '', // vacío
          documento: '', // vacío
          documento_2: '',
          documento_3: '',
          monto_mn: 0,
          monto: parseFloat(String(rowArray[7] || '0')) || 0, // Columna 8: Importe
          tc: '',
          monto_abonado: 0,
          monto_abonado_2: 0,
          cuenta_tipo: '',
          cuenta_numero: String(rowArray[2] || '').replace(/-/g, ''), // Columna 3: Cuenta
          cuenta_nombre: '',
          estado: String(rowArray[8] || ''), // Columna 9: Situación
          observaciones: '', // vacío
          banco: 'BBVA', // Siempre BBVA para archivos BBVA
          origen: data2.fileName
        };
      } else {
        // MAPEO PARA BCP usando NÚMEROS DE COLUMNA EXACTOS
        record = {
          id: `${data2.fileName}_${index}`,
          beneficiario: String(rowArray[1] || ''), // Columna B (Beneficiario)
          documento_tipo: String(rowArray[2] || '') || '-', // Columna C (Documento - Tipo)
          documento: String(rowArray[3] || '') || '-', // Columna D (Documento)
          documento_2: '',
          documento_3: '',
          monto_mn: 0,
          monto: parseFloat(String(rowArray[7] || '0')) || 0, // Columna H (Monto)
          tc: '',
          monto_abonado: 0,
          monto_abonado_2: 0,
          cuenta_tipo: '',
          cuenta_numero: String(rowArray[13] || '').replace(/-/g, ''), // Columna N (Cuenta)
          cuenta_nombre: '',
          estado: String(rowArray[14] || ''), // Columna O (Estado)
          observaciones: String(rowArray[15] || ''), // Columna P (Observación)
          banco: bankType2 || (data2.fileName?.toLowerCase().includes('bbva') ? 'BBVA' : 'BCP'),
          origen: data2.fileName
        };
      }
      
      // Debug: mostrar los valores de documento
      console.log(`${isBBVA2 ? 'BBVA' : 'BCP'} Record ${index}: documento_tipo = "${record.documento_tipo}", documento = "${record.documento}"`);
      console.log(`${isBBVA2 ? 'BBVA' : 'BCP'} Row data:`, row);
      
      // Incluir registros que tengan algún dato
      if (record.beneficiario || record.monto > 0 || record.estado || record.cuenta_numero) {
        combinedRecords.push(record);
      }
    });
  });
  
  return {
    records: combinedRecords,
    totalRecords: combinedRecords.length,
    sources: [data1.fileName, data2.fileName],
    processedAt: new Date()
  };
};

export const createSingleFileData = (data: ExcelData, bankType?: 'BBVA' | 'BCP'): CombinedData => {
  const records: AbonoRecord[] = [];
  
  // Procesar el archivo único
  console.log(`=== PROCESANDO ARCHIVO: ${data.fileName} ===`);
  console.log(`Tipo de banco especificado: ${bankType}`);
  console.log(`Total de hojas: ${data.sheets.length}`);
  data.sheets.forEach((sheet, sheetIndex) => {
    console.log(`Hoja ${sheetIndex}: ${sheet.name} - Filas: ${sheet.data.length}`);
    console.log(`Headers de la hoja ${sheetIndex}:`, sheet.headers);
    
    const isBBVA = bankType === 'BBVA';
    
    // Para BBVA, solo procesar hojas que tengan datos
    if (isBBVA && sheet.data.length === 0) {
      console.log(`BBVA: Saltando hoja ${sheetIndex} porque no tiene datos`);
      return;
    }
    
    // Debug: mostrar las primeras 3 filas de datos para verificar
    if (isBBVA) {
      console.log(`BBVA - Primera fila de datos:`, sheet.data[0]);
      console.log(`BBVA - Segunda fila de datos:`, sheet.data[1]);
      console.log(`BBVA - Tercera fila de datos:`, sheet.data[2]);
    } else {
      console.log(`BCP - Primera fila de datos:`, sheet.data[0]);
      console.log(`BCP - Segunda fila de datos:`, sheet.data[1]);
      console.log(`BCP - Tercera fila de datos:`, sheet.data[2]);
      
      // Debug: verificar si hay datos en las columnas C y D
      console.log(`BCP - Verificando columnas C y D:`);
      console.log(`  - Columna C (índice 2): "${sheet.data[0]?.[2]}"`);
      console.log(`  - Columna D (índice 3): "${sheet.data[0]?.[3]}"`);
      console.log(`  - Columna C (índice 2): "${sheet.data[1]?.[2]}"`);
      console.log(`  - Columna D (índice 3): "${sheet.data[1]?.[3]}"`);
    }
    
    sheet.data.forEach((row, index) => {
      let record: AbonoRecord;
      
      if (isBBVA) {
        // Para BBVA: los datos vienen como array de columnas
        const rowArray = Array.isArray(row) ? row : Object.values(row);
        
        // MAPEO ESPECÍFICO PARA BBVA según la tabla "Relación de las cuentas de abono"
        // Headers: ['Sel', 'No.', 'Cuenta', 'Banco', 'Titular(Archivo)', 'Titular(Banco)', 'Doc.Identidad', 'Importe', 'Situación']
        record = {
          id: `${data.fileName}_${index}`,
          // Mapeo según especificaciones:
          // item = columna 2 (No.) - índice 1
          // beneficiario = columna 5 (Titular(Archivo)) - índice 4
          // documento = vacío
          // # documento = vacío
          // monto = columna 8 (Importe) - índice 7
          // cuenta = columna 3 (Cuenta) - índice 2
          // estado = columna 9 (Situación) - índice 8
          // observacion = vacío
          // banco = BBVA
          beneficiario: String(rowArray[4] || ''), // Columna 5: Titular(Archivo) - índice 4
          documento_tipo: '', // vacío
          documento: '', // vacío
          documento_2: '',
          documento_3: '',
          monto_mn: 0,
          monto: parseFloat(String(rowArray[7] || '0')) || 0, // Columna 8: Importe - índice 7
          tc: '',
          monto_abonado: 0,
          monto_abonado_2: 0,
          cuenta_tipo: '',
          cuenta_numero: String(rowArray[2] || '').replace(/-/g, ''), // Columna 3: Cuenta - índice 2
          cuenta_nombre: '',
          estado: String(rowArray[8] || ''), // Columna 9: Situación - índice 8
          observaciones: '', // vacío
          banco: 'BBVA', // Siempre BBVA para archivos BBVA
          origen: data.fileName
        };
        
        console.log(`BBVA Record ${index}:`, record);
      } else {
        // Para BCP: usar lógica existente
        const rowArray = row as any[];
        
        record = {
          id: `${data.fileName}_${index}`,
          beneficiario: String(rowArray[1] || ''), // Columna B = 2
          documento_tipo: String(rowArray[2] || '') || '-', // Columna C = 3
          documento: String(rowArray[3] || '') || '-', // Columna D = 4
          documento_2: '',
          documento_3: '',
          monto_mn: 0,
          monto: parseFloat(String(rowArray[7] || '0')) || 0, // Columna H = 8
          tc: '',
          monto_abonado: 0,
          monto_abonado_2: 0,
          cuenta_tipo: '',
          cuenta_numero: String(rowArray[13] || '').replace(/-/g, ''), // Columna N = 14
          cuenta_nombre: '',
          estado: String(rowArray[14] || ''), // Columna O = 15
          observaciones: String(rowArray[15] || ''), // Columna P = 16
          banco: bankType || 'BCP',
          origen: data.fileName
        };
      }
      
      // Debug: mostrar información del archivo y datos
      if (isBBVA) {
        console.log(`=== BBVA ARCHIVO PROCESANDO: ${data.fileName} ===`);
        console.log(`BBVA Record ${index}: beneficiario = "${record.beneficiario}", monto = "${record.monto}"`);
      } else {
        console.log(`=== BCP ARCHIVO PROCESANDO: ${data.fileName} ===`);
        console.log(`BCP Record ${index}: documento_tipo = "${record.documento_tipo}", documento = "${record.documento}"`);
      }
      // Incluir registros que tengan algún dato
      if (record.beneficiario || record.monto > 0 || record.estado || record.cuenta_numero) {
        records.push(record);
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

export const exportCombinedToCSV = (data: CombinedData): void => {
  const headers = [
    'Beneficiario', 'Documento Tipo', 'Documento', 'Documento 2', 'Documento 3',
    'Monto M/N', 'Monto', 'T/C', 'Monto Abonado', 'Monto Abonado 2',
    'Cuenta Tipo', 'Cuenta Número', 'Cuenta Nombre', 'Estado', 'Observaciones', 'Banco', 'Origen'
  ];
  const csvContent = [
    headers.join(','),
    ...data.records.map(record => [
      `"${record.beneficiario}"`,
      `"${record.documento_tipo}"`,
      `"${record.documento}"`,
      `"${record.documento_2}"`,
      `"${record.documento_3}"`,
      record.monto_mn,
      record.monto,
      `"${record.tc}"`,
      record.monto_abonado,
      record.monto_abonado_2,
      `"${record.cuenta_tipo}"`,
      `"${record.cuenta_numero}"`,
      `"${record.cuenta_nombre}"`,
      `"${record.estado}"`,
      `"${record.observaciones}"`,
      `"${record.banco}"`,
      `"${record.origen}"`
    ].join(','))
  ].join('\n');
  
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);
  link.setAttribute('href', url);
  link.setAttribute('download', `abonos_taxi_monterrico_${new Date().toISOString().split('T')[0]}.csv`);
  link.style.visibility = 'hidden';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};

export const isValidExcelFile = (file: File): boolean => {
  const validTypes = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel',
    '.xlsx',
    '.xls'
  ];
  
  return validTypes.some(type => 
    file.type === type || file.name?.toLowerCase().endsWith(type)
  );
};

export const formatFileSize = (bytes: number): string => {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
};

export const exportToCSV = (data: ExcelRow[], filename: string): void => {
  if (data.length === 0) return;
  
  const headers = Object.keys(data[0]);
  const csvContent = [
    headers.join(','),
    ...data.map(row => 
      headers.map(header => `"${String(row[header] || '').replace(/"/g, '""')}"`).join(',')
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