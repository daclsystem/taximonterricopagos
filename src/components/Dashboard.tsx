/**
 * Componente Dashboard - Sistema de Carga de Abonos
 */
import React, { useState } from 'react';
import { Car, Upload, FileSpreadsheet, LogOut } from 'lucide-react';
import { FileUploader } from './FileUploader';
import { CombinedDataViewer } from './CombinedDataViewer';
import { FileUploadState, CombinedData } from '../types/excel';
import { combineExcelData, createSingleFileData } from '../utils/excelProcessor';
import { getSession, clearSession } from '../utils/auth';

interface DashboardProps {
  onLogout: () => void;
}

export const Dashboard: React.FC<DashboardProps> = ({ onLogout }) => {
  const [file1State, setFile1State] = useState<FileUploadState>({
    file: null,
    data: null,
    isLoading: false,
    error: null
  });

  const [file2State, setFile2State] = useState<FileUploadState>({
    file: null,
    data: null,
    isLoading: false,
    error: null
  });

  const [combinedData, setCombinedData] = useState<CombinedData | null>(null);

  const session = getSession();

  const handleFile1Upload = (state: FileUploadState) => {
    setFile1State(state);
    if (state.data) {
      if (file2State.data) {
        // Si hay ambos archivos, combinarlos
        const combined = combineExcelData(state.data, file2State.data, 'BBVA', 'BCP');
        setCombinedData(combined);
      } else {
        // Si solo hay un archivo, crear datos combinados con solo ese archivo
        const combined = createSingleFileData(state.data, 'BBVA');
        setCombinedData(combined);
      }
    }
  };

  const handleFile2Upload = (state: FileUploadState) => {
    setFile2State(state);
    if (state.data) {
      if (file1State.data) {
        // Si hay ambos archivos, combinarlos
        const combined = combineExcelData(file1State.data, state.data, 'BBVA', 'BCP');
        setCombinedData(combined);
      } else {
        // Si solo hay un archivo, crear datos combinados con solo ese archivo
        const combined = createSingleFileData(state.data, 'BCP');
        setCombinedData(combined);
      }
    }
  };

  const resetFiles = () => {
    setFile1State({
      file: null,
      data: null,
      isLoading: false,
      error: null
    });
    setFile2State({
      file: null,
      data: null,
      isLoading: false,
      error: null
    });
    setCombinedData(null);
  };

  const handleLogout = () => {
    clearSession();
    onLogout();
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 via-white to-green-50">
      {/* Header */}
      <header className="bg-white shadow-sm border-b border-gray-200">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center justify-between h-16">
            <div className="flex items-center">
              <img 
                src="https://taximonterrico.com/assets/logo_variante-CoJ5dU2i.png" 
                alt="Taxi Monterrico" 
                className="h-10 w-auto mr-4"
              />
              <div>
                <h1 className="text-2xl font-bold text-gray-900">
                  Carga Abonos
                </h1>
                <p className="text-sm text-gray-600">Sistema de procesamiento de archivos Excel</p>
              </div>
            </div>
            <div className="flex items-center space-x-4">
              {session && (
                <div className="flex items-center space-x-3">
                  <img 
                    src={session.fotop} 
                    alt="Usuario" 
                    className="h-8 w-8 rounded-full object-cover"
                    onError={(e) => {
                      (e.target as HTMLImageElement).src = 'https://via.placeholder.com/32x32?text=U';
                    }}
                  />
                  <span className="text-sm font-medium text-gray-700">
                    {session.idacceso}
                  </span>
                </div>
              )}
              <button
                onClick={handleLogout}
                className="inline-flex items-center px-4 py-2 border border-red-300 rounded-md shadow-sm text-sm font-medium text-red-700 bg-white hover:bg-red-50 focus:outline-none focus:ring-2 focus:ring-red-500 transition-colors"
              >
                <LogOut className="h-4 w-4 mr-2" />
                Cerrar Sesión
              </button>
            </div>
          </div>
        </div>
      </header>

      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        <div className="space-y-8">
          <div className="text-center">
            <div className="flex items-center justify-center mb-4">
              <Car className="h-12 w-12 text-blue-600 mr-3" />
              <FileSpreadsheet className="h-12 w-12 text-green-600" />
            </div>
            <h2 className="text-3xl font-bold text-gray-900 mb-4">
              Sistema de Carga de Abonos
            </h2>
            <p className="text-lg text-gray-600 max-w-3xl mx-auto">
              Sube archivos Excel con formatos de abonos. El sistema los procesará 
              automáticamente y los mostrará en una tabla unificada para Taxi Monterrico.
            </p>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
            <FileUploader
              label="Archivo BBVA"
              onFileProcessed={handleFile1Upload}
              uploadState={file1State}
              bankType="BBVA"
            />
            <FileUploader
              label="Archivo BCP"
              onFileProcessed={handleFile2Upload}
              uploadState={file2State}
              bankType="BCP"
            />
          </div>

          {/* Estado del procesamiento */}
          <div className="bg-white rounded-lg border border-gray-200 p-6">
            <h3 className="text-lg font-semibold text-gray-800 mb-4">Estado del Procesamiento</h3>
            <div className="space-y-3">
              <div className="flex items-center">
                <div className={`w-4 h-4 rounded-full mr-3 ${
                  file1State.data ? 'bg-green-500' : file1State.error ? 'bg-red-500' : 'bg-gray-300'
                }`}></div>
                <span className="text-sm text-gray-700">
                  Archivo BBVA: {file1State.data ? 'Procesado' : file1State.error ? 'Error' : 'Pendiente'}
                </span>
              </div>
              <div className="flex items-center">
                <div className={`w-4 h-4 rounded-full mr-3 ${
                  file2State.data ? 'bg-green-500' : file2State.error ? 'bg-red-500' : 'bg-gray-300'
                }`}></div>
                <span className="text-sm text-gray-700">
                  Archivo BCP: {file2State.data ? 'Procesado' : file2State.error ? 'Error' : 'Pendiente'}
                </span>
              </div>
              <div className="flex items-center">
                <div className={`w-4 h-4 rounded-full mr-3 ${
                  combinedData ? 'bg-green-500' : 'bg-gray-300'
                }`}></div>
                <span className="text-sm text-gray-700">
                  Procesamiento: {combinedData ? 'Completado' : 'Pendiente'}
                </span>
              </div>
            </div>
          </div>

          {/* Tabla de Abonos */}
          <CombinedDataViewer data={combinedData} />
        </div>
      </div>

      {/* Footer */}
      <footer className="bg-white border-t border-gray-200 mt-16">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
          <div className="flex items-center justify-center space-x-4">
            <img 
              src="https://taximonterrico.com/assets/logo_variante-CoJ5dU2i.png" 
              alt="Taxi Monterrico" 
              className="h-8 w-auto"
            />
            <div className="text-center text-gray-600">
              <p>Sistema de Carga de Abonos - Taxi Monterrico</p>
              <p className="text-sm">Procesamiento inteligente de archivos Excel</p>
            </div>
          </div>
        </div>
      </footer>
    </div>
  );
};