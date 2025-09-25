/**
 * Componente para subir archivos Excel con drag & drop
 */
import React, { useCallback, useState } from 'react';
import { Upload, File, AlertCircle, CheckCircle } from 'lucide-react';
import { FileUploadState } from '../types/excel';
import { processExcelFile, isValidExcelFile, formatFileSize } from '../utils/excelProcessor';
import { processBBVAFile } from '../utils/excelProcessorBBVA_v2';

interface FileUploaderProps {
  label: string;
  onFileProcessed: (data: FileUploadState) => void;
  uploadState: FileUploadState;
  bankType?: 'BBVA' | 'BCP';
}

export const FileUploader: React.FC<FileUploaderProps> = ({ 
  label, 
  onFileProcessed, 
  uploadState,
  bankType
}) => {
  const [dragOver, setDragOver] = useState(false);

  const handleFileUpload = useCallback(async (file: File) => {
    // Validar tipo de archivo
    if (!isValidExcelFile(file)) {
      onFileProcessed({
        file: null,
        data: null,
        isLoading: false,
        error: 'Por favor, selecciona un archivo Excel válido (.xlsx o .xls)'
      });
      return;
    }

    // Iniciar carga
    onFileProcessed({
      file,
      data: null,
      isLoading: true,
      error: null
    });

    try {
      let data;
      if (bankType === 'BBVA') {
        data = await processBBVAFile(file);
      } else {
        data = await processExcelFile(file, bankType);
      }
      onFileProcessed({
        file,
        data,
        isLoading: false,
        error: null
      });
    } catch (error) {
      onFileProcessed({
        file: null,
        data: null,
        isLoading: false,
        error: error instanceof Error ? error.message : 'Error procesando archivo'
      });
    }
  }, [onFileProcessed]);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    
    const files = Array.from(e.dataTransfer.files);
    if (files.length > 0) {
      handleFileUpload(files[0]);
    }
  }, [handleFileUpload]);

  const handleFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      handleFileUpload(files[0]);
    }
  }, [handleFileUpload]);

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
  }, []);

  return (
    <div className="w-full">
      <h3 className="text-lg font-semibold text-gray-800 mb-3">{label}</h3>
      
      <div
        className={`
          relative border-2 border-dashed rounded-lg p-6 transition-all duration-200
          ${dragOver 
            ? 'border-blue-500 bg-blue-50' 
            : uploadState.data 
              ? 'border-green-500 bg-green-50'
              : uploadState.error
                ? 'border-red-500 bg-red-50'
                : 'border-gray-300 hover:border-gray-400 hover:bg-gray-50'
          }
        `}
        onDrop={handleDrop}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
      >
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileSelect}
          className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
          disabled={uploadState.isLoading}
        />
        
        <div className="text-center">
          {uploadState.isLoading ? (
            <div className="space-y-2">
              <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-500 mx-auto"></div>
              <p className="text-sm text-gray-600">Procesando archivo...</p>
            </div>
          ) : uploadState.data ? (
            <div className="space-y-2">
              <CheckCircle className="h-8 w-8 text-green-500 mx-auto" />
              <p className="text-sm font-medium text-gray-900">{uploadState.file?.name}</p>
              <p className="text-xs text-gray-500">
                {formatFileSize(uploadState.file?.size || 0)} • {uploadState.data.totalRows} filas
              </p>
            </div>
          ) : uploadState.error ? (
            <div className="space-y-2">
              <AlertCircle className="h-8 w-8 text-red-500 mx-auto" />
              <p className="text-sm text-red-600">{uploadState.error}</p>
              <p className="text-xs text-gray-500">Haz clic o arrastra otro archivo</p>
            </div>
          ) : (
            <div className="space-y-2">
              <Upload className="h-8 w-8 text-gray-400 mx-auto" />
              <p className="text-sm font-medium text-gray-900">
                Arrastra tu archivo Excel aquí
              </p>
              <p className="text-xs text-gray-500">
                O haz clic para seleccionar • .xlsx, .xls
              </p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};