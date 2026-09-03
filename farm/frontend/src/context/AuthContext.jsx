/**
 * context/AuthContext.jsx — Global auth state.
 */

import { createContext, useContext, useEffect, useState } from 'react';
import client from '../api/client';

const AuthContext = createContext(null);

export function AuthProvider({ children }) {
  const [faculty, setFaculty] = useState(null);
  const [loading, setLoading] = useState(true);
  // Bumped on every faculty refresh so <img src="/api/faculty/{id}/photo"> tags
  // (whose URL is otherwise identical before/after a re-upload) are forced to
  // re-fetch instead of showing the browser's cached copy of the old photo.
  const [photoVersion, setPhotoVersion] = useState(0);

  useEffect(() => {
    const token = localStorage.getItem('access_token');
    if (!token) { setLoading(false); return; }
    client.get('/auth/me')
      .then(r => { setFaculty(r.data); setPhotoVersion(v => v + 1); })
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
    setPhotoVersion(v => v + 1);
  };

  const logout = () => {
    localStorage.removeItem('access_token');
    setFaculty(null);
  };

  /**
   * refreshFaculty — re-fetches /auth/me from the live database and updates
   * the faculty state. Call this after any profile mutation (photo/signature/key
   * upload) so the status badges — and the photo <img>, via photoVersion — update
   * immediately without requiring a page reload.
   */
  const refreshFaculty = async () => {
    try {
      const me = await client.get('/auth/me');
      setFaculty(me.data);
      setPhotoVersion(v => v + 1);
    } catch {
      // Token may have expired — force logout
      localStorage.removeItem('access_token');
      setFaculty(null);
    }
  };

  return (
    <AuthContext.Provider value={{ faculty, loading, login, logout, refreshFaculty, photoVersion }}>
      {children}
    </AuthContext.Provider>
  );
}

export const useAuth = () => useContext(AuthContext);
