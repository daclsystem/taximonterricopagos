/**
 * Componente para comparar dos archivos Excel procesados
 */
import React from 'react';
import { GitCompare, FileText, Users, Calendar } from 'lucide-react';
import { ExcelData } from '../types/excel';

interface ComparisonViewProps {
  file1: ExcelData | null;
  file2: ExcelData | null;
}

export const ComparisonView: React.FC<ComparisonViewProps> = ({ file1, file2 }) => {
  if (!file1 || !file2) {
    return (
      <div className="bg-white rounded-lg border border-gray-200 p-8 text-center">
        <GitCompare className="h-12 w-12 text-gray-400 mx-auto mb-4" />
        <h3 className="text-lg font-medium text-gray-900 mb-2">
          Comparación de Archivos
        </h3>
        <p className="text-gray-500">
          Sube ambos archivos Excel para ver la comparación
        </p>
      </div>
    );
  }

  const getUniqueHeaders = (data: ExcelData): string[] => {
    const allHeaders = new Set<string>();
    data.sheets.forEach(sheet => {
      sheet.headers.forEach(header => allHeaders.add(header));
    });
    return Array.from(allHeaders);
  };

  const file1Headers = getUniqueHeaders(file1);
  const file2Headers = getUniqueHeaders(file2);
  const commonHeaders = file1Headers.filter(h => file2Headers.includes(h));
  const uniqueToFile1 = file1Headers.filter(h => !file2Headers.includes(h));
  const uniqueToFile2 = file2Headers.filter(h => !file1Headers.includes(h));

  return (
    <div className="bg-white rounded-lg border border-gray-200 shadow-sm">
      <div className="p-6 border-b border-gray-200">
        <h3 className="text-lg font-semibold text-gray-800 flex items-center">
          <GitCompare className="h-5 w-5 mr-2 text-blue-500" />
          Comparación de Archivos
        </h3>
      </div>

      <div className="p-6 space-y-6">
        {/* Estadísticas generales */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div className="space-y-4">
            <h4 className="font-medium text-gray-900 flex items-center">
              <FileText className="h-4 w-4 mr-2 text-blue-500" />
              {file1.fileName}
            </h4>
            <div className="grid grid-cols-2 gap-4">
              <div className="bg-blue-50 p-3 rounded-lg">
                <p className="text-xs text-blue-600 font-medium">Hojas</p>
                <p className="text-lg font-semibold text-blue-900">{file1.sheets.length}</p>
              </div>
              <div className="bg-green-50 p-3 rounded-lg">
                <p className="text-xs text-green-600 font-medium">Filas</p>
                <p className="text-lg font-semibold text-green-900">{file1.totalRows}</p>
              </div>
              <div className="bg-purple-50 p-3 rounded-lg">
                <p className="text-xs text-purple-600 font-medium">Columnas</p>
                <p className="text-lg font-semibold text-purple-900">{file1Headers.length}</p>
              </div>
              <div className="bg-orange-50 p-3 rounded-lg">
                <p className="text-xs text-orange-600 font-medium">Fecha</p>
                <p className="text-sm font-semibold text-orange-900">
                  {file1.uploadedAt.toLocaleDateString('es-ES')}
                </p>
              </div>
            </div>
          </div>

          <div className="space-y-4">
            <h4 className="font-medium text-gray-900 flex items-center">
              <FileText className="h-4 w-4 mr-2 text-green-500" />
              {file2.fileName}
            </h4>
            <div className="grid grid-cols-2 gap-4">
              <div className="bg-blue-50 p-3 rounded-lg">
                <p className="text-xs text-blue-600 font-medium">Hojas</p>
                <p className="text-lg font-semibold text-blue-900">{file2.sheets.length}</p>
              </div>
              <div className="bg-green-50 p-3 rounded-lg">
                <p className="text-xs text-green-600 font-medium">Filas</p>
                <p className="text-lg font-semibold text-green-900">{file2.totalRows}</p>
              </div>
              <div className="bg-purple-50 p-3 rounded-lg">
                <p className="text-xs text-purple-600 font-medium">Columnas</p>
                <p className="text-lg font-semibold text-purple-900">{file2Headers.length}</p>
              </div>
              <div className="bg-orange-50 p-3 rounded-lg">
                <p className="text-xs text-orange-600 font-medium">Fecha</p>
                <p className="text-sm font-semibold text-orange-900">
                  {file2.uploadedAt.toLocaleDateString('es-ES')}
                </p>
              </div>
            </div>
          </div>
        </div>

        {/* Análisis de columnas */}
        <div className="border-t border-gray-200 pt-6">
          <h4 className="font-medium text-gray-900 mb-4">Análisis de Columnas</h4>
          
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="bg-green-50 border border-green-200 rounded-lg p-4">
              <h5 className="font-medium text-green-800 mb-2">
                Columnas Comunes ({commonHeaders.length})
              </h5>
              {commonHeaders.length > 0 ? (
                <div className="space-y-1 max-h-32 overflow-y-auto">
                  {commonHeaders.map((header, index) => (
                    <span
                      key={index}
                      className="inline-block bg-green-100 text-green-800 text-xs px-2 py-1 rounded-full mr-1 mb-1"
                    >
                      {header}
                    </span>
                  ))}
                </div>
              ) : (
                <p className="text-sm text-green-600">No hay columnas en común</p>
              )}
            </div>

            <div className="bg-blue-50 border border-blue-200 rounded-lg p-4">
              <h5 className="font-medium text-blue-800 mb-2">
                Solo en {file1.fileName.split('.')[0]} ({uniqueToFile1.length})
              </h5>
              {uniqueToFile1.length > 0 ? (
                <div className="space-y-1 max-h-32 overflow-y-auto">
                  {uniqueToFile1.map((header, index) => (
                    <span
                      key={index}
                      className="inline-block bg-blue-100 text-blue-800 text-xs px-2 py-1 rounded-full mr-1 mb-1"
                    >
                      {header}
                    </span>
                  ))}
                </div>
              ) : (
                <p className="text-sm text-blue-600">Todas las columnas son comunes</p>
              )}
            </div>

            <div className="bg-orange-50 border border-orange-200 rounded-lg p-4">
              <h5 className="font-medium text-orange-800 mb-2">
                Solo en {file2.fileName.split('.')[0]} ({uniqueToFile2.length})
              </h5>
              {uniqueToFile2.length > 0 ? (
                <div className="space-y-1 max-h-32 overflow-y-auto">
                  {uniqueToFile2.map((header, index) => (
                    <span
                      key={index}
                      className="inline-block bg-orange-100 text-orange-800 text-xs px-2 py-1 rounded-full mr-1 mb-1"
                    >
                      {header}
                    </span>
                  ))}
                </div>
              ) : (
                <p className="text-sm text-orange-600">Todas las columnas son comunes</p>
              )}
            </div>
          </div>
        </div>

        {/* Resumen de compatibilidad */}
        <div className="border-t border-gray-200 pt-6">
          <div className="bg-gray-50 rounded-lg p-4">
            <h5 className="font-medium text-gray-800 mb-2">Resumen de Compatibilidad</h5>
            <div className="text-sm text-gray-600 space-y-1">
              <p>
                • <strong>Compatibilidad de columnas:</strong>{' '}
                {((commonHeaders.length / Math.max(file1Headers.length, file2Headers.length)) * 100).toFixed(1)}%
              </p>
              <p>
                • <strong>Diferencia de filas:</strong>{' '}
                {Math.abs(file1.totalRows - file2.totalRows).toLocaleString()} filas
              </p>
              <p>
                • <strong>Diferencia de hojas:</strong>{' '}
                {Math.abs(file1.sheets.length - file2.sheets.length)} hojas
              </p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};