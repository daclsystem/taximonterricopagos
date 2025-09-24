/**
 * Tipos para el sistema de autenticación - Taxi Monterrico
 */

export interface LoginRequest {
  agente: string;
  contrasena: string;
  idempresas: number;
}

export interface LoginResponse {
  idusuario: number;
  idacceso: string;
  fotop: string;
  estatus: number;
  message: string;
  msystem: string;
}

export interface UserSession {
  idusuario: number;
  idacceso: string;
  fotop: string;
  isAuthenticated: boolean;
}