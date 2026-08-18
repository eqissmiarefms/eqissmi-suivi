const API_URL = import.meta.env.VITE_SHEETS_API_URL; // URL .../exec du déploiement Apps Script
const API_TOKEN = import.meta.env.VITE_SHEETS_API_TOKEN; // doit correspondre à API_TOKEN côté script

async function getRequest(action, params = {}) {
  const query = new URLSearchParams({ action, token: API_TOKEN, ...params }).toString();
  const res = await fetch(`${API_URL}?${query}`);
  const data = await res.json();
  if (data.error) throw new Error(data.error);
  return data;
}

async function postRequest(action, payload = {}) {
  const res = await fetch(API_URL, {
    method: "POST",
    // "text/plain" évite le pré-vol CORS (OPTIONS) qu'Apps Script ne gère pas bien
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action, token: API_TOKEN, ...payload }),
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error);
  return data;
}

export const api = {
  getState: () => getRequest("getState"),

  createSession: (name) => postRequest("createSession", { name }),
  deleteSession: (id) => postRequest("deleteSession", { id }),

  createModule: (name) => postRequest("createModule", { name }),
  deleteModule: (id) => postRequest("deleteModule", { id }),

  createEntry: (entry) => postRequest("createEntry", { entry }),
  updateEntry: (id, patch) => postRequest("updateEntry", { id, patch }),
  deleteEntry: (id) => postRequest("deleteEntry", { id }),

  importBulk: (rows) => postRequest("importBulk", { rows }),
};