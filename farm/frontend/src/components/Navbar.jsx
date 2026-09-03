/**
 * components/Navbar.jsx
 */

import { Link, useLocation } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { LayoutDashboard, Settings, LogOut, GraduationCap, FileText } from 'lucide-react';

export default function Navbar() {
  const { faculty, logout, photoVersion } = useAuth();
  const { pathname } = useLocation();

  const navLink = (to, icon, label) => {
    const active = pathname === to;
    return (
      <Link
        to={to}
        style={{
          display: 'flex',
          alignItems: 'center',
          gap: '6px',
          padding: '6px 14px',
          borderRadius: '8px',
          fontSize: '14px',
          fontWeight: '500',
          transition: 'all 0.15s',
          background: active ? 'rgba(132, 169, 140, 0.15)' : 'transparent', // Light sage
          color: active ? 'var(--color-primary-hover)' : 'var(--color-text-muted)',
          border: active ? '1px solid rgba(132, 169, 140, 0.3)' : '1px solid transparent',
          textDecoration: 'none',
        }}
      >
        {icon}
        {label}
      </Link>
    );
  };

  return (
    <nav style={{
      background: 'var(--color-surface)',
      borderBottom: '1px solid var(--color-border)',
      padding: '0 24px',
      height: '60px',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'space-between',
      position: 'sticky',
      top: 0,
      zIndex: 100,
    }}>
      {/* Brand */}
      <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
        <div style={{
          width: 34, height: 34, borderRadius: '8px',
          background: 'linear-gradient(135deg, var(--color-primary), var(--color-primary-hover))',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
        }}>
          <GraduationCap size={18} color="white" />
        </div>
        <span style={{ fontWeight: '700', fontSize: '16px', letterSpacing: '-0.3px', color: 'var(--color-text)' }}>
          LearnerPDC
        </span>
      </div>

      {/* Links */}
      <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
        {navLink('/dashboard', <LayoutDashboard size={15} />, 'Dashboard')}
        {navLink('/reports', <FileText size={15} />, 'Reports')}
        {navLink('/profile', <Settings size={15} />, 'Profile')}
      </div>

      {/* User + Logout */}
      <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
        {faculty && (
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <div style={{ textAlign: 'right' }}>
              <div style={{ fontSize: '13px', fontWeight: '600' }}>{faculty.name}</div>
              <div style={{ fontSize: '11px', color: 'var(--color-text-muted)' }}>{faculty.role}</div>
            </div>
            {faculty.has_photo ? (
              <img 
                src={`/api/faculty/${faculty._id}/photo?v=${photoVersion}`}
                alt={faculty.name}
                style={{ width: 36, height: 36, borderRadius: '50%', objectFit: 'cover', border: '1px solid var(--color-border)' }}
              />
            ) : (
              <div style={{
                width: 36, height: 36, borderRadius: '50%',
                background: 'var(--color-surface-2)', border: '1px solid var(--color-border)',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                fontSize: '14px', fontWeight: '700', color: 'var(--color-text-muted)',
              }}>
                {faculty.name?.[0]?.toUpperCase() || '?'}
              </div>
            )}
          </div>
        )}
        <button
          onClick={logout}
          title="Logout"
          style={{
            background: 'none',
            border: '1px solid var(--color-border)',
            borderRadius: '8px',
            padding: '6px 10px',
            cursor: 'pointer',
            color: 'var(--color-text-muted)',
            display: 'flex',
            alignItems: 'center',
            transition: 'all 0.15s',
          }}
          onMouseEnter={e => { e.target.style.color = 'var(--color-danger)'; e.target.style.borderColor = 'var(--color-danger)'; }}
          onMouseLeave={e => { e.target.style.color = 'var(--color-text-muted)'; e.target.style.borderColor = 'var(--color-border)'; }}
        >
          <LogOut size={15} />
        </button>
      </div>
    </nav>
  );
}
