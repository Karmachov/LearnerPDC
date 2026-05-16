/**
 * pages/Login.jsx — email/password login + registration toggle.
 */

import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import client from '../api/client';
import { GraduationCap, Mail, Lock, User, Building2, Eye, EyeOff } from 'lucide-react';

/**
 * parseApiError — always returns a plain string safe for React rendering.
 *
 * FastAPI error shapes we must handle:
 *  • 401 / 409 / 500  → detail is a string   → use directly
 *  • 422 Unprocessable → detail is an array   → e.g. [{loc, msg, type}, ...]
 *                        React Error #31 fires if you put an array into {error}
 *  • Network error     → no response at all   → use err.message
 */
function parseApiError(err) {
  const detail = err?.response?.data?.detail;

  if (!detail) {
    // Network timeout, CORS block, or unknown JS error
    return err?.message || 'An unexpected error occurred. Please try again.';
  }

  if (typeof detail === 'string') {
    return detail;
  }

  if (Array.isArray(detail)) {
    // FastAPI 422: each item is { loc: string[], msg: string, type: string }
    return detail
      .map((d) => {
        const field = Array.isArray(d.loc) ? d.loc.filter(p => p !== 'body').join(' → ') : '';
        const msg   = d.msg ?? 'Validation error';
        return field ? `${field}: ${msg}` : msg;
      })
      .join('  •  ');
  }

  if (typeof detail === 'object') {
    // Unexpected object shape — JSON-stringify as last resort
    try { return JSON.stringify(detail); } catch { return 'An unexpected error occurred.'; }
  }

  return String(detail);
}

const inputStyle = {
  width: '100%',
  background: 'var(--color-surface-2)',
  border: '1px solid var(--color-border)',
  borderRadius: '10px',
  padding: '11px 14px 11px 40px',
  color: 'var(--color-text)',
  fontSize: '14px',
  outline: 'none',
  transition: 'border-color 0.15s',
  boxSizing: 'border-box',
};

function Field({ icon, label, ...props }) {
  const [show, setShow] = useState(false);
  const isPassword = props.type === 'password';
  return (
    <div style={{ marginBottom: '14px' }}>
      {/* htmlFor links the label to its input for accessibility and click-focus */}
      <label
        htmlFor={props.id}
        style={{ display: 'block', fontSize: '12px', color: 'var(--color-text-muted)', marginBottom: '6px', fontWeight: '500' }}
      >
        {label}
      </label>
      <div style={{ position: 'relative' }}>
        <span style={{ position: 'absolute', left: '12px', top: '50%', transform: 'translateY(-50%)', color: 'var(--color-text-muted)' }}>
          {icon}
        </span>
        <input
          {...props}
          type={isPassword && show ? 'text' : props.type}
          style={inputStyle}
          onFocus={e => e.target.style.borderColor = 'var(--color-primary)'}
          onBlur={e => e.target.style.borderColor = 'var(--color-border)'}
        />
        {isPassword && (
          <button
            type="button"
            onClick={() => setShow(s => !s)}
            style={{ position: 'absolute', right: '12px', top: '50%', transform: 'translateY(-50%)', background: 'none', border: 'none', cursor: 'pointer', color: 'var(--color-text-muted)', padding: 0 }}
          >
            {show ? <EyeOff size={15} /> : <Eye size={15} />}
          </button>
        )}
      </div>
    </div>
  );
}

export default function Login() {
  const { login } = useAuth();
  const navigate = useNavigate();
  const [isRegister, setIsRegister] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const [form, setForm] = useState({
    email: '', password: '', name: '', role: 'Faculty', department: 'Computer Science and Engineering',
  });

  const set = (k) => (e) => setForm(f => ({ ...f, [k]: e.target.value }));

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');
    setLoading(true);
    try {
      // ── Registration path ───────────────────────────────────────────────
      // Caught in its own block so a 422/409 from /auth/register never
      // cascades into a spurious /auth/token call.
      if (isRegister) {
        try {
          await client.post('/auth/register', form);
        } catch (regErr) {
          setError(parseApiError(regErr));
          return;  // stop — do not attempt login after a failed register
        }
      }

      // ── Login path (always runs after a successful register too) ────────
      await login(form.email, form.password);
      navigate('/dashboard');
    } catch (err) {
      // Catches login-specific errors (wrong password, 401, network, etc.)
      setError(parseApiError(err));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{
      minHeight: '100vh',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      background: 'radial-gradient(ellipse 80% 50% at 50% -20%, rgba(124,58,237,0.15), transparent)',
      padding: '20px',
    }}>
      <div style={{
        width: '100%',
        maxWidth: '420px',
        background: 'var(--color-surface)',
        border: '1px solid var(--color-border)',
        borderRadius: '20px',
        padding: '40px',
        boxShadow: '0 25px 50px rgba(0,0,0,0.4)',
      }}>
        {/* Logo */}
        <div style={{ textAlign: 'center', marginBottom: '32px' }}>
          <div style={{
            width: 56, height: 56, borderRadius: '16px',
            background: 'linear-gradient(135deg, #7c3aed, #06b6d4)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            margin: '0 auto 16px',
            boxShadow: '0 8px 24px rgba(124,58,237,0.4)',
          }}>
            <GraduationCap size={28} color="white" />
          </div>
          <h1 style={{ margin: 0, fontSize: '22px', fontWeight: '700' }}>LearnerPDC</h1>
          <p style={{ margin: '6px 0 0', color: 'var(--color-text-muted)', fontSize: '14px' }}>
            {isRegister ? 'Create your faculty account' : 'Sign in to your account'}
          </p>
        </div>

        <form onSubmit={handleSubmit}>
          {isRegister && (
            <>
              {/*
                id/name match FacultyCreate field names exactly.
                The register POST sends React state as JSON — HTML name attrs
                are for accessibility/autofill only, but they must be consistent.
              */}
              <Field
                label="Full Name"
                icon={<User size={14} />}
                type="text"
                id="name"
                name="name"
                value={form.name}
                onChange={set('name')}
                required
                placeholder="Dr. Jane Smith"
                autoComplete="name"
              />
              <Field
                label="Role"
                icon={<User size={14} />}
                type="text"
                id="role"
                name="role"
                value={form.role}
                onChange={set('role')}
                placeholder="Faculty / HOD / Lab Instructor"
              />
              <Field
                label="Department"
                icon={<Building2 size={14} />}
                type="text"
                id="department"
                name="department"
                value={form.department}
                onChange={set('department')}
                autoComplete="organization"
              />
            </>
          )}
          {/*
            Email field dual-mode name/id:
            • Register mode → id="email" name="email"  (matches FacultyCreate JSON key)
            • Login mode    → id="username" name="username"  (matches OAuth2PasswordRequestForm)
            The actual wire format is always driven by React state + Axios, not these
            HTML attributes — but they must be semantically correct for autofill,
            password managers, and accessibility tools to work properly in both modes.
          */}
          <Field
            label="Email Address"
            icon={<Mail size={14} />}
            type="email"
            id={isRegister ? 'email' : 'username'}
            name={isRegister ? 'email' : 'username'}
            value={form.email}
            onChange={set('email')}
            required
            placeholder="faculty@manipal.edu"
            autoComplete={isRegister ? 'email' : 'username'}
          />
          <Field
            label="Password"
            icon={<Lock size={14} />}
            type="password"
            id="password"
            name="password"
            value={form.password}
            onChange={set('password')}
            required
            placeholder="••••••••"
            autoComplete={isRegister ? 'new-password' : 'current-password'}
          />

          {error && (
            <div style={{
              background: 'rgba(239,68,68,0.1)', border: '1px solid rgba(239,68,68,0.3)',
              borderRadius: '8px', padding: '10px 14px', fontSize: '13px',
              color: '#fca5a5', marginBottom: '14px',
            }}>
              {error}
            </div>
          )}

          <button
            type="submit"
            disabled={loading}
            style={{
              width: '100%',
              background: loading ? 'var(--color-surface-2)' : 'linear-gradient(135deg, #7c3aed, #6d28d9)',
              color: loading ? 'var(--color-text-muted)' : 'white',
              border: 'none',
              borderRadius: '10px',
              padding: '12px',
              fontWeight: '700',
              fontSize: '15px',
              cursor: loading ? 'not-allowed' : 'pointer',
              transition: 'all 0.15s',
              marginTop: '4px',
              boxShadow: loading ? 'none' : '0 4px 15px rgba(124,58,237,0.4)',
            }}
          >
            {loading ? 'Please wait…' : isRegister ? 'Create Account' : 'Sign In'}
          </button>
        </form>

        <div style={{ textAlign: 'center', marginTop: '20px', fontSize: '13px', color: 'var(--color-text-muted)' }}>
          {isRegister ? 'Already have an account?' : "Don't have an account?"}
          {' '}
          <button
            onClick={() => { setIsRegister(r => !r); setError(''); }}
            style={{ background: 'none', border: 'none', color: '#a78bfa', cursor: 'pointer', fontWeight: '600', fontSize: '13px' }}
          >
            {isRegister ? 'Sign In' : 'Register'}
          </button>
        </div>
      </div>
    </div>
  );
}
