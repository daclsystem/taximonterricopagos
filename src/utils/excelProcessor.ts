/**
 * Utilidades para procesar archivos Excel - Sistema de Abonos Taxi Monterrico
 */
import * as XLSX from 'xlsx';
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
      const found = headers.find((header, index) => {
        const headerLower = header?.toLowerCase() || '';
        return headerLower.includes('documento') && headerLower.includes('tipo') && !headerLower.includes('dcumento');
      });
      if (found) return found;
    }
    
    // Para documento (se mapea a columna "# Documento"), buscar "documento" que NO contenga "tipo"
    if (targetField === 'documento') {
      const found = headers.find((header, index) => {
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

export const processExcelFile = (file: File): Promise<ExcelData> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'array' });
        const sheets: ExcelSheet[] = [];
        let totalRows = 0;
        
        workbook.SheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          if (worksheet) {
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            if (jsonData.length === 0) return;
            
                    // Detectar si es archivo BBVA o BCP
                    let isBBVAFile = false;
                    let isBCPFile = false;
                    let selRowIndex = -1;
                    
                    // Primero verificar por nombre de archivo
                    const fileName = file.name.toLowerCase();
                    console.log(`=== DETECTANDO TIPO DE ARCHIVO ===`);
                    console.log(`Nombre del archivo: "${file.name}"`);
                    console.log(`¿Contiene 'bbva'?: ${fileName.includes('bbva')}`);
                    console.log(`¿Contiene 'bcp'?: ${fileName.includes('bcp')}`);
                    
                    if (fileName.includes('bbva')) {
                      isBBVAFile = true;
                      console.log('Archivo BBVA detectado por nombre');
                    } else if (fileName.includes('bcp')) {
                      isBCPFile = true;
                      console.log('Archivo BCP detectado por nombre');
                    } else {
                      // Si no se puede determinar por nombre, buscar patrones en el contenido
                      for (let i = 0; i < jsonData.length; i++) {
                        const row = jsonData[i] as any[];
                        if (row && row.length > 0) {
                          const foundSel = row.some(cell => {
                            const cellStr = String(cell || '').toLowerCase().trim();
                            return cellStr === 'sel' || cellStr === 'sel.';
                          });
                          
                          // Buscar patrones típicos de BBVA
                          const rowText = row.map(cell => String(cell || '').toLowerCase()).join(' ');
                          const hasBBVAPatterns = rowText.includes('pagos masivos') || 
                                                rowText.includes('consulta de') ||
                                                rowText.includes('monterrico') ||
                                                rowText.includes('abonos enviados');
                          
                          // Buscar patrones típicos de BCP
                          const hasBCPPatterns = rowText.includes('beneficiario - nombre') ||
                                               rowText.includes('documento - tipo') ||
                                               rowText.includes('monto - moneda') ||
                                               rowText.includes('cuenta - número') ||
                                               rowText.includes('payment report');
                          
                          if (foundSel || hasBBVAPatterns) {
                            isBBVAFile = true;
                            if (foundSel) selRowIndex = i;
                            break;
                          } else if (hasBCPPatterns) {
                            isBCPFile = true;
                            console.log('Archivo BCP detectado por patrones en contenido');
                            break;
                          }
                        }
                      }
                      
                      // Si no se detectó nada, asumir que es BCP por defecto
                      if (!isBBVAFile && !isBCPFile) {
                        isBCPFile = true;
                        console.log('Archivo BCP asumido por defecto');
                      }
                    }
                    
                    let headers: string[];
                    let dataStartIndex: number;
                    let dataEndIndex = jsonData.length;
                    
                    if (isBBVAFile) {
                      console.log('Archivo BBVA detectado - buscando estructura de datos');
                      
                      // Buscar la fila con "Sel" como header
                      let headerRowIndex = -1;
                      if (selRowIndex !== -1) {
                        headerRowIndex = selRowIndex;
                      } else {
                        // Buscar fila que contenga "Sel" y otros headers típicos de BBVA
                        for (let i = 0; i < Math.min(50, jsonData.length); i++) {
                          const row = jsonData[i] as any[];
                          if (row && row.length > 5) {
                            const rowText = row.map(cell => String(cell || '').toLowerCase()).join(' ');
                            if (rowText.includes('sel') && 
                                (rowText.includes('cuenta') || rowText.includes('titular') || rowText.includes('importe'))) {
                              headerRowIndex = i;
                              break;
                            }
                          }
                        }
                      }
                      
                      if (headerRowIndex !== -1) {
                        headers = jsonData[headerRowIndex] as string[];
                        dataStartIndex = headerRowIndex + 1;
                        console.log(`BBVA: Headers encontrados en fila ${headerRowIndex + 1}, datos desde fila ${dataStartIndex + 1}`);
                      } else {
                        // Headers por defecto basados en la imagen
                        headers = ['Sel', 'No.', 'Cuenta', 'Banco', 'Titular(Archivo)', 'Titular(Banco)', 'Doc.Identidad', 'Importe', 'Situación'];
                        dataStartIndex = 31; // Fila 32 por defecto
                        console.log('BBVA: Usando headers por defecto y fila 32');
                      }
                      
                      // Buscar "Estimado Cliente:" o "Los tipos de documentos" para terminar
                      // Empezar desde más adelante para evitar terminar muy temprano
                      for (let i = Math.max(dataStartIndex + 5, 35); i < jsonData.length; i++) {
                        const row = jsonData[i] as any[];
                        if (row && row.length > 0) {
                          const rowText = row.map(cell => String(cell || '').toLowerCase()).join(' ');
                          if (rowText.includes('estimado cliente') || 
                              rowText.includes('los tipos de documentos') ||
                              rowText.includes('tipos de documentos') ||
                              rowText.includes('r: ruc') ||
                              rowText.includes('l: dni')) {
                            dataEndIndex = i;
                            console.log(`BBVA: Fin de datos encontrado en fila ${i + 1} - texto: "${rowText.substring(0, 50)}..."`);
                            break;
                          }
                        }
                      }
                      
                      // Si no se encontró el final, usar un rango más amplio
                      if (dataEndIndex === jsonData.length || dataEndIndex <= dataStartIndex + 5) {
                        // Para BBVA, usar un rango más amplio desde el inicio de datos
                        dataEndIndex = Math.min(dataStartIndex + 50, jsonData.length);
                        console.log(`BBVA: No se encontró fin específico, usando rango hasta fila ${dataEndIndex + 1}`);
                      }
                      
                      // Asegurar que dataEndIndex sea válido
                      if (dataEndIndex <= dataStartIndex) {
                        dataEndIndex = Math.min(dataStartIndex + 50, jsonData.length);
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
                    console.log(`Datos relevantes BBVA: ${relevantData.length} filas encontradas`);
                    
                    const dataRows = relevantData
                      .filter(row => {
                        const rowArray = row as any[];
                        if (!rowArray || rowArray.length === 0) return false;
                        
                        const firstCell = String(rowArray[0] || '').trim();
                        
                        if (isBBVAFile) {
                          // Para BBVA: debe empezar con un número (1, 2, 3, etc.)
                          const isNumber = /^\d+$/.test(firstCell);
                          if (isNumber) {
                            // Verificar que tenga datos en columnas importantes
                            const hasCuenta = rowArray[2] && String(rowArray[2] || '').trim() !== '';
                            const hasTitular = rowArray[4] && String(rowArray[4] || '').trim() !== '';
                            const hasImporte = rowArray[7] && String(rowArray[7] || '').trim() !== '';
                            
                            // Al menos debe tener cuenta, titular o importe
                            return hasCuenta || hasTitular || hasImporte;
                          }
                          
                          // Si no empieza con número, verificar si tiene datos válidos en columnas importantes
                          const hasCuenta = rowArray[2] && String(rowArray[2] || '').trim() !== '';
                          const hasTitular = rowArray[4] && String(rowArray[4] || '').trim() !== '';
                          const hasImporte = rowArray[7] && String(rowArray[7] || '').trim() !== '';
                          
                          return hasCuenta || hasTitular || hasImporte;
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
              console.log('Primeras 3 filas relevantes:', relevantData.slice(0, 3));
              console.log('Headers mapeados:', cleanHeaders);
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
            
            // Headers fijos para la tabla - basados en las columnas resaltadas en amarillo del BCP
            const fixedHeaders = [
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
            
            sheets.push({
              name: sheetName,
              data: dataRows,
              headers: fixedHeaders,
              rowCount: dataRows.length
            });
            
            totalRows += dataRows.length;
          }
        });
        
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
    const isBCP1 = bankType1 === 'BCP' || data1.fileName?.toLowerCase().includes('bcp') || false;
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
    
    sheet.data.forEach((row, index) => {
      // MAPEO USANDO HEADERS REALES DEL ARCHIVO BCP
      const record: AbonoRecord = {
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
      
      // Debug: mostrar los valores de documento
      console.log(`BCP Record ${index}: documento_tipo = "${record.documento_tipo}", documento = "${record.documento}"`);
      console.log(`BCP Row data:`, row);
      
      // Para BCP, incluir TODOS los registros que tengan algún dato
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
      // MAPEO CORRECTO SEGÚN ESPECIFICACIONES - BCP
      const rowArray = row as any[];
      
      const record: AbonoRecord = {
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
      
      // Debug: mostrar los valores de documento
      console.log(`BCP Record ${index}: documento_tipo = "${record.documento_tipo}", documento = "${record.documento}"`);
      console.log(`BCP Row data:`, row);
      
      // Para BCP, incluir TODOS los registros que tengan algún dato
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
  console.log(`Total de hojas: ${data.sheets.length}`);
  data.sheets.forEach((sheet, sheetIndex) => {
    console.log(`Hoja ${sheetIndex}: ${sheet.name} - Filas: ${sheet.data.length}`);
    console.log(`Headers de la hoja ${sheetIndex}:`, sheet.headers);
    
    const isBBVA = bankType === 'BBVA' || data.fileName?.toLowerCase().includes('bbva') || false;
    const isBCP = bankType === 'BCP' || data.fileName?.toLowerCase().includes('bcp') || false;
    const mappedFields = {
      beneficiario: findBestMatch(sheet.headers, 'beneficiario', isBBVA, isBCP),
      documento_tipo: findBestMatch(sheet.headers, 'documento_tipo', isBBVA, isBCP),
      documento: findBestMatch(sheet.headers, 'documento', isBBVA, isBCP),
      documento_2: findBestMatch(sheet.headers, 'documento_2', isBBVA, isBCP),
      documento_3: findBestMatch(sheet.headers, 'documento_3', isBBVA, isBCP),
      monto_mn: findBestMatch(sheet.headers, 'monto_mn', isBBVA, isBCP),
      monto: findBestMatch(sheet.headers, 'monto', isBBVA, isBCP),
      tc: findBestMatch(sheet.headers, 'tc', isBBVA, isBCP),
      monto_abonado: findBestMatch(sheet.headers, 'monto_abonado', isBBVA, isBCP),
      monto_abonado_2: findBestMatch(sheet.headers, 'monto_abonado_2', isBBVA, isBCP),
      cuenta_tipo: findBestMatch(sheet.headers, 'cuenta_tipo', isBBVA, isBCP),
      cuenta_numero: findBestMatch(sheet.headers, 'cuenta_numero', isBBVA, isBCP),
      cuenta_nombre: findBestMatch(sheet.headers, 'cuenta_nombre', isBBVA, isBCP),
      estado: findBestMatch(sheet.headers, 'estado', isBBVA, isBCP),
      observaciones: findBestMatch(sheet.headers, 'observaciones', isBBVA, isBCP),
      banco: findBestMatch(sheet.headers, 'banco', isBBVA, isBCP)
    };
    
    // Debug: mostrar las primeras 3 filas de datos para verificar
    console.log(`BCP - Primera fila de datos:`, sheet.data[0]);
    console.log(`BCP - Segunda fila de datos:`, sheet.data[1]);
    console.log(`BCP - Tercera fila de datos:`, sheet.data[2]);
    
    // Debug: verificar si hay datos en las columnas C y D
    console.log(`BCP - Verificando columnas C y D:`);
    console.log(`  - Columna C (índice 2): "${sheet.data[0]?.[2]}"`);
    console.log(`  - Columna D (índice 3): "${sheet.data[0]?.[3]}"`);
    console.log(`  - Columna C (índice 2): "${sheet.data[1]?.[2]}"`);
    console.log(`  - Columna D (índice 3): "${sheet.data[1]?.[3]}"`);
    
    sheet.data.forEach((row, index) => {
      // Para archivos BCP, usar NÚMEROS DE COLUMNA EXACTOS
      const rowArray = row as any[];
      
      const record: AbonoRecord = {
        id: `${data.fileName}_${index}`,
        // MAPEO USANDO NÚMEROS DE COLUMNA EXACTOS:
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
        banco: bankType || (data.fileName?.toLowerCase().includes('bbva') ? 'BBVA' : 'BCP'),
        origen: data.fileName
      };
      
      // Debug: mostrar información del archivo y datos
      console.log(`=== BCP ARCHIVO PROCESANDO: ${data.fileName} ===`);
      console.log(`BCP Record ${index}: documento_tipo = "${record.documento_tipo}", documento = "${record.documento}"`);
      console.log(`BCP Row data:`, row);
      console.log(`BCP Headers disponibles:`, Object.keys(row));
      
      // DEBUG ESPECÍFICO: Verificar TODAS las columnas disponibles
      console.log(`BCP TODAS LAS COLUMNAS DISPONIBLES:`);
      Object.keys(row).forEach((key, idx) => {
        console.log(`  ${idx}: "${key}" = "${row[key]}"`);
      });
      
      // Debug: mostrar el banco asignado
      console.log(`BCP Record ${index}: banco = "${record.banco}", fileName = "${data.fileName}"`);
      
      // Para BCP, incluir TODOS los registros que tengan algún dato
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