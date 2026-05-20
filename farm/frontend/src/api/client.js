/**
 * api/client.js — Axios instance with JWT interceptor.
 * VITE_API_URL defaults to '/api' which the Vite dev proxy rewrites to http://backend:8000
 */

import axios from 'axios';

const BASE_URL = import.meta.env.VITE_API_URL || '/api';

const client = axios.create({
  baseURL: BASE_URL,
  timeout: 30000,
});

// Attach JWT on every request
client.interceptors.request.use((config) => {
  const token = localStorage.getItem('access_token');
  if (token) {
    config.headers['Authorization'] = `Bearer ${token}`;
  }
  return config;
});

// On 401 → clear session and redirect to login
client.interceptors.response.use(
  (res) => res,
  (err) => {
    if (err.response?.status === 401) {
      localStorage.removeItem('access_token');
      window.location.href = '/login';
    }
    return Promise.reject(err);
  }
);

/**
 * Download a generated report with JWT auth (plain <a href> cannot send Bearer tokens).
 */
export async function downloadReport(taskId) {
  const response = await client.get(`/download/${taskId}`, {
    responseType: 'blob',
    timeout: 120000,
  });

  const disposition = response.headers['content-disposition'] || '';
  let filename = `report_${taskId.slice(0, 8)}`;
  const utf8Match = disposition.match(/filename\*=UTF-8''([^;]+)/i);
  const plainMatch = disposition.match(/filename="?([^";\n]+)"?/i);
  if (utf8Match) {
    filename = decodeURIComponent(utf8Match[1]);
  } else if (plainMatch) {
    filename = plainMatch[1];
  }

  const blobUrl = window.URL.createObjectURL(response.data);
  const anchor = document.createElement('a');
  anchor.href = blobUrl;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.URL.revokeObjectURL(blobUrl);
}

export async function downloadProofs(taskId) {
  const response = await client.get(`/download-proofs/${taskId}`, {
    responseType: 'blob',
    timeout: 120000,
  });

  const disposition = response.headers['content-disposition'] || '';
  let filename = `proofs_${taskId.slice(0, 8)}.zip`;
  const utf8Match = disposition.match(/filename\*=UTF-8''([^;]+)/i);
  const plainMatch = disposition.match(/filename="?([^";\n]+)"?/i);
  if (utf8Match) {
    filename = decodeURIComponent(utf8Match[1]);
  } else if (plainMatch) {
    filename = plainMatch[1];
  }

  const blobUrl = window.URL.createObjectURL(response.data);
  const anchor = document.createElement('a');
  anchor.href = blobUrl;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.URL.revokeObjectURL(blobUrl);
}

export default client;
