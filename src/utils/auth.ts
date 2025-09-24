/**
 * Utilidades para autenticación - Taxi Monterrico
 */
import { LoginRequest, LoginResponse, UserSession } from '../types/auth';

const API_BASE_URL = 'https://api.taximonterrico.com/api';
const SESSION_KEY = 'taxi_monterrico_session';

export const login = async (agente: string, contrasena: string): Promise<LoginResponse> => {
  const loginData: LoginRequest = {
    agente,
    contrasena,
    idempresas: 0
  };

  try {
    const response = await fetch(`${API_BASE_URL}/Eventos/Login`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(loginData)
    });

    const data: LoginResponse = await response.json();
    return data;
  } catch (error) {
    throw new Error(error instanceof Error ? error.message : 'Error de conexión');
  }
};

export const saveSession = (loginResponse: LoginResponse): void => {
  const session: UserSession = {
    idusuario: loginResponse.idusuario,
    idacceso: loginResponse.idacceso,
    fotop: loginResponse.fotop,
    isAuthenticated: true
  };
  
  localStorage.setItem(SESSION_KEY, JSON.stringify(session));
};

export const getSession = (): UserSession | null => {
  try {
    const sessionData = localStorage.getItem(SESSION_KEY);
    if (!sessionData) return null;
    
    return JSON.parse(sessionData);
  } catch {
    return null;
  }
};

export const clearSession = (): void => {
  localStorage.removeItem(SESSION_KEY);
};

export const isAuthenticated = (): boolean => {
  const session = getSession();
  return session?.isAuthenticated || false;
};