import { state } from "./state.js";

export async function loadConfig() {
  const response = await fetch("/.netlify/functions/config");
  if (!response.ok) {
    throw new Error("No se pudo cargar la configuracion publica.");
  }
  state.config = await response.json();
  return state.config;
}
