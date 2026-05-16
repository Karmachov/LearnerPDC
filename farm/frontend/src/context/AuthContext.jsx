/**
 * context/AuthContext.jsx — Global auth state.
 */

import { createContext, useContext, useEffect, useState } from 'react';
import client from '../api/client';

const AuthContext = createContext(null);

export function AuthProvider({ children }) {
  const [faculty, setFaculty] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const token = localStorage.getItem('access_token');
    if (!token) { setLoading(false); return; }
    client.get('/auth/me')
      .then(r => setFaculty(r.data))
      .catch(() => localStorage.removeItem('access_token'))
      .finally(() => setLoading(false));
  }, []);

  const login = async (email, password) => {
    const params = new URLSearchParams({ username: email, password });
    const r = await client.post('/auth/token', params, {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    });
    localStorage.setItem('access_token', r.data.access_token);
    const me = await client.get('/auth/me');
    setFaculty(me.data);
  };

  const logout = () => {
    localStorage.removeItem('access_token');
    setFaculty(null);
  };

  return (
    <AuthContext.Provider value={{ faculty, loading, login, logout }}>
      {children}
    </AuthContext.Provider>
  );
}

export const useAuth = () => useContext(AuthContext);
