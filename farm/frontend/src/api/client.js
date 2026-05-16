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

export default client;
