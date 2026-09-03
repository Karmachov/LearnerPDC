/**
 * pages/Profile.jsx — Upload and persist signature image + signing keys to MongoDB vault.
 */

import { useState } from 'react';
import { useAuth } from '../context/AuthContext';
import client from '../api/client';

import {
  ShieldCheck, Key, ImageIcon, CheckCircle, XCircle,
  Upload, Lock, AlertTriangle, RefreshCw,
} from 'lucide-react';

/**
 * parseUploadError — converts any Axios error into a safe plain string.
 * Handles:
 *  • FastAPI 400/422 where detail is a string or array of validation objects
 *  • 413 Payload Too Large (thrown by Nginx before FastAPI sees the request)
 *  • Network errors (no response object)
 */
function parseUploadError(err) {
  const status = err?.response?.status;
  const detail = err?.response?.data?.detail;

  if (status === 413) {
    return 'File is too large. The server rejected it (413 Payload Too Large). Try a smaller file.';
  }

  if (!detail) {
    return err?.message || 'Upload failed. Check your network connection and try again.';
  }

  if (typeof detail === 'string') return detail;

  if (Array.isArray(detail)) {
    return detail
      .map(d => {
        const field = Array.isArray(d.loc) ? d.loc.filter(p => p !== 'body').join(' → ') : '';
        return field ? `${field}: ${d.msg}` : d.msg;
      })
      .join('  •  ');
  }

  try { return JSON.stringify(detail); } catch { return 'Upload failed.'; }
}

function Card({ children, style }) {
  return (
    <div style={{
      background: 'var(--color-surface)',
      border: '1px solid var(--color-border)',
      borderRadius: '16px',
      padding: '28px',
      ...style,
    }}>{children}</div>
  );
}

function StatusBadge({ ok, label }) {
  return (
    <span style={{
      display: 'inline-flex', alignItems: 'center', gap: 5,
      fontSize: 12, fontWeight: 600,
      color: ok ? 'var(--color-success)' : 'var(--color-text-muted)',
      background: ok ? 'rgba(16,185,129,0.1)' : 'rgba(139,148,158,0.1)',
      border: `1px solid ${ok ? 'rgba(16,185,129,0.3)' : 'var(--color-border)'}`,
      borderRadius: 99, padding: '3px 10px',
    }}>
      {ok ? <CheckCircle size={11} /> : <XCircle size={11} />}
      {label}
    </span>
  );
}

function FileDropZone({ label, accept, file, setFile, hint }) {
  const [drag, setDrag] = useState(false);
  return (
    <div style={{ marginBottom: 16 }}>
      <label style={{ fontSize: 12, fontWeight: 600, color: 'var(--color-text-muted)', display: 'block', marginBottom: 6 }}>
        {label}
      </label>
      <label
        style={{
          display: 'flex', alignItems: 'center', gap: 10,
          background: drag ? 'rgba(124,58,237,0.08)' : 'var(--color-surface-2)',
          border: `2px dashed ${drag ? 'var(--color-primary)' : file ? 'var(--color-success)' : 'var(--color-border)'}`,
          borderRadius: 10, padding: '14px 18px', cursor: 'pointer',
          transition: 'all 0.15s',
        }}
        onDragEnter={() => setDrag(true)}
        onDragLeave={() => setDrag(false)}
        onDrop={e => { e.preventDefault(); setDrag(false); setFile(e.dataTransfer.files[0]); }}
        onDragOver={e => e.preventDefault()}
      >
        <Upload size={16} color={file ? 'var(--color-success)' : 'var(--color-text-muted)'} />
        <span style={{ fontSize: 13, color: file ? 'var(--color-success)' : 'var(--color-text-muted)' }}>
          {file ? `✓ ${file.name}` : hint}
        </span>
        <input type="file" accept={accept} style={{ display: 'none' }} onChange={e => setFile(e.target.files[0])} />
      </label>
    </div>
  );
}

function UploadSection({ title, icon, children, onSubmit, loading, success, error }) {
  return (
    <Card>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 20 }}>
        <div style={{
          width: 36, height: 36, borderRadius: 10,
          background: 'linear-gradient(135deg, rgba(132,169,140,0.2), rgba(132,169,140,0.05))',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
        }}>
          {icon}
        </div>
        <h2 style={{ margin: 0, fontSize: 16, fontWeight: 700 }}>{title}</h2>
      </div>

      {children}

      {error && (
        <div style={{
          display: 'flex', alignItems: 'center', gap: 8,
          background: 'rgba(239,68,68,0.1)', border: '1px solid rgba(239,68,68,0.3)',
          borderRadius: 8, padding: '10px 14px', fontSize: 13, color: '#fca5a5', marginBottom: 14,
        }}>
          <AlertTriangle size={13} />
          {error}
        </div>
      )}

      {success && (
        <div style={{
          display: 'flex', alignItems: 'center', gap: 8,
          background: 'rgba(16,185,129,0.1)', border: '1px solid rgba(16,185,129,0.3)',
          borderRadius: 8, padding: '10px 14px', fontSize: 13, color: '#6ee7b7', marginBottom: 14,
        }}>
          <CheckCircle size={13} />
          {success}
        </div>
      )}

      <button
        onClick={onSubmit}
        disabled={loading}
        style={{
          background: loading ? 'var(--color-surface-2)' : 'linear-gradient(135deg, var(--color-primary), var(--color-primary-hover))',
          color: loading ? 'var(--color-text-muted)' : 'white',
          border: 'none', borderRadius: 10, padding: '10px 22px',
          fontWeight: 700, fontSize: 14, cursor: loading ? 'not-allowed' : 'pointer',
          display: 'flex', alignItems: 'center', gap: 7, transition: 'all 0.15s',
          boxShadow: loading ? 'none' : '0 4px 14px rgba(132,169,140,0.35)',
        }}
      >
        {loading ? <><RefreshCw size={14} style={{ animation: 'spin 1s linear infinite' }} /> Saving…</> : 'Save to Vault'}
      </button>
    </Card>
  );
}

export default function Profile() {
  const { faculty, refreshFaculty, photoVersion } = useAuth();

  // Signature upload state
  const [sigFile, setSigFile] = useState(null);
  const [sigLoading, setSigLoading] = useState(false);
  const [sigSuccess, setSigSuccess] = useState('');
  const [sigError, setSigError] = useState('');

  // Photo upload state
  const [photoFile, setPhotoFile] = useState(null);
  const [photoLoading, setPhotoLoading] = useState(false);
  const [photoSuccess, setPhotoSuccess] = useState('');
  const [photoError, setPhotoError] = useState('');

  // Key upload state
  const [keyFile, setKeyFile] = useState(null);
  const [certFile, setCertFile] = useState(null);
  const [keyPassword, setKeyPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [keyLoading, setKeyLoading] = useState(false);
  const [keySuccess, setKeySuccess] = useState('');
  const [keyError, setKeyError] = useState('');

  const uploadSignature = async () => {
    if (!sigFile) { setSigError('Please select an image file.'); return; }
    setSigError(''); setSigSuccess(''); setSigLoading(true);
    try {
      const fd = new FormData();
      fd.append('image', sigFile);
      await client.put('/profile/signature', fd);
      setSigSuccess('Initial image encrypted and saved to the secure vault.');
      await refreshFaculty();
    } catch (err) {
      setSigError(parseUploadError(err));
    } finally {
      setSigLoading(false);
    }
  };

  const uploadPhoto = async () => {
    if (!photoFile) { setPhotoError('Please select a photo file.'); return; }
    setPhotoError(''); setPhotoSuccess(''); setPhotoLoading(true);
    try {
      const fd = new FormData();
      fd.append('image', photoFile);
      await client.put('/profile/photo', fd);
      setPhotoSuccess('Profile photo updated successfully.');
      await refreshFaculty();
    } catch (err) {
      setPhotoError(parseUploadError(err));
    } finally {
      setPhotoLoading(false);
    }
  };

  const uploadKeys = async () => {
    if (!keyFile || !certFile) { setKeyError('Please select both the private key and certificate.'); return; }
    if (!keyPassword) { setKeyError('Please enter the private key passphrase.'); return; }
    setKeyError(''); setKeySuccess(''); setKeyLoading(true);
    try {
      const fd = new FormData();
      fd.append('private_key', keyFile);
      fd.append('certificate', certFile);
      fd.append('key_password', keyPassword);
      await client.put('/profile/keys', fd);
      setKeySuccess(
        'Private key, certificate, and passphrase saved. Generate a PDF with "Enable digital signing" to verify.'
      );
      setKeyPassword('');
      // Refresh the faculty profile so the status badges update immediately.
      await refreshFaculty();
    } catch (err) {
      setKeyError(parseUploadError(err));
    } finally {
      setKeyLoading(false);
    }
  };

  return (
    <>
      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
      <div style={{ maxWidth: 800, margin: '0 auto', padding: '32px 20px' }}>
        {/* Page header */}
        <div style={{ marginBottom: 28 }}>
          <h1 style={{ margin: 0, fontSize: 26, fontWeight: 800, letterSpacing: '-0.5px' }}>Profile Settings</h1>
          <p style={{ margin: '6px 0 0', color: 'var(--color-text-muted)', fontSize: 14 }}>
            Manage your faculty profile and securely store signing credentials.
          </p>
        </div>

        {/* Faculty info card */}
        <Card style={{ marginBottom: 20, display: 'flex', alignItems: 'center', gap: 20, flexWrap: 'wrap' }}>
          {faculty?.has_photo ? (
            <img 
              src={`/api/faculty/${faculty._id}/photo?v=${photoVersion}`}
              alt={faculty.name}
              style={{
                width: 54, height: 54, borderRadius: '50%', objectFit: 'cover',
                border: '1px solid var(--color-border)', flexShrink: 0
              }}
            />
          ) : (
            <div style={{
              width: 54, height: 54, borderRadius: '50%',
              background: 'linear-gradient(135deg, var(--color-primary), var(--color-primary-hover))',
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              fontSize: 22, fontWeight: 700, color: 'white', flexShrink: 0,
            }}>
              {faculty?.name?.[0]?.toUpperCase() || '?'}
            </div>
          )}
          <div style={{ flex: 1 }}>
            <div style={{ fontWeight: 700, fontSize: 18 }}>{faculty?.name}</div>
            <div style={{ color: 'var(--color-text-muted)', fontSize: 13 }}>{faculty?.role} · {faculty?.department}</div>
            <div style={{ color: 'var(--color-text-muted)', fontSize: 13 }}>{faculty?.email}</div>
          </div>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            <StatusBadge ok={faculty?.has_signature} label="Initial" />
            {/* <StatusBadge ok={faculty?.has_private_key} label="Private Key" /> */}
            {/* <StatusBadge ok={faculty?.has_certificate} label="Certificate" /> */}
          </div>
        </Card>

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20 }}>
          {/* Profile Photo upload */}
          <UploadSection
            title="Profile Photo"
            icon={<ImageIcon size={18} color="var(--color-primary)" />}
            onSubmit={uploadPhoto}
            loading={photoLoading}
            success={photoSuccess}
            error={photoError}
          >
            <p style={{ fontSize: 13, color: 'var(--color-text-muted)', marginTop: 0, marginBottom: 16 }}>
              Update your profile photo. JPEG or PNG format.
            </p>
            <FileDropZone
              label="Profile Photo"
              accept="image/png,image/jpeg,image/webp"
              file={photoFile}
              setFile={setPhotoFile}
              hint="Drop PNG/JPEG here or click to browse…"
            />
          </UploadSection>

          {/* Signature upload */}
          <UploadSection
            title="Initial"
            icon={<ImageIcon size={18} color="var(--color-primary)" />}
            onSubmit={uploadSignature}
            loading={sigLoading}
            success={sigSuccess}
            error={sigError}
          >
            <p style={{ fontSize: 13, color: 'var(--color-text-muted)', marginTop: 0, marginBottom: 16 }}>
              This image will appear on all report pages. PNG with transparent background recommended.
            </p>
            <FileDropZone
              label="Initial"
              accept="image/png,image/jpeg"
              file={sigFile}
              setFile={setSigFile}
              hint="Drop PNG/JPEG here or click to browse…"
            />
          </UploadSection>

          {/* Key upload - commented out for current release */}
          {/* <UploadSection
            title="Signing Credentials"
            icon={<Key size={18} color="var(--color-primary)" />}
            onSubmit={uploadKeys}
            loading={keyLoading}
            success={keySuccess}
            error={keyError}
          >
            <p style={{ fontSize: 13, color: 'var(--color-text-muted)', marginTop: 0, marginBottom: 16 }}>
              Your private key and certificate are Fernet-encrypted before being stored. They are never stored in plaintext.
            </p>
            <FileDropZone
              label="Private Key (.pem)"
              accept=".pem"
              file={keyFile}
              setFile={setKeyFile}
              hint="Drop private_key.pem here…"
            />
            <FileDropZone
              label="Certificate (.pem)"
              accept=".pem"
              file={certFile}
              setFile={setCertFile}
              hint="Drop certificate.pem here…"
            />

            <div style={{ marginBottom: 16 }}>
              <label style={{ fontSize: 12, fontWeight: 600, color: 'var(--color-text-muted)', display: 'block', marginBottom: 6 }}>
                Private Key Passphrase <span style={{ color: 'var(--color-danger)' }}>*</span>
              </label>
              <div style={{ position: 'relative' }}>
                <Lock size={13} style={{ position: 'absolute', left: 12, top: '50%', transform: 'translateY(-50%)', color: 'var(--color-text-muted)' }} />
                <input
                  type={showPassword ? 'text' : 'password'}
                  value={keyPassword}
                  onChange={e => setKeyPassword(e.target.value)}
                  placeholder="Key passphrase…"
                  style={{
                    width: '100%', background: 'var(--color-surface-2)',
                    border: '1px solid var(--color-border)', borderRadius: 8,
                    padding: '9px 38px 9px 34px', color: 'var(--color-text)', fontSize: 14,
                    outline: 'none', boxSizing: 'border-box',
                  }}
                  onFocus={e => e.target.style.borderColor = 'var(--color-primary)'}
                  onBlur={e => e.target.style.borderColor = 'var(--color-border)'}
                />
                <button
                  type="button"
                  onClick={() => setShowPassword(s => !s)}
                  style={{ position: 'absolute', right: 10, top: '50%', transform: 'translateY(-50%)', background: 'none', border: 'none', cursor: 'pointer', color: 'var(--color-text-muted)', padding: 0 }}
                >
                  <ShieldCheck size={14} color={showPassword ? 'var(--color-primary)' : undefined} />
                </button>
              </div>
              <p style={{ fontSize: 11, color: 'var(--color-text-muted)', margin: '4px 0 0' }}>
                Required every time you click Save to Vault. Use the same passphrase that unlocks your
                private key file (test with OpenSSL if unsure).
              </p>
            </div>
          </UploadSection> */}
        </div>
      </div>
    </>
  );
}
