/**
 * components/TaskStatusCard.jsx — animated progress indicator.
 */

import { useState } from 'react';
import { CheckCircle, XCircle, Loader2, Clock, Download } from 'lucide-react';
import { downloadReport } from '../api/client';

const STATUS_CONFIG = {
  PENDING: {
    icon: <Clock size={22} color="#f59e0b" />,
    label: 'Queued',
    color: '#f59e0b',
    barColor: 'linear-gradient(90deg, #f59e0b, #fbbf24)',
  },
  STARTED: {
    icon: <Loader2 size={22} color="#06b6d4" style={{ animation: 'spin 1s linear infinite' }} />,
    label: 'Processing',
    color: '#06b6d4',
    barColor: 'linear-gradient(90deg, #06b6d4, #7c3aed)',
  },
  SUCCESS: {
    icon: <CheckCircle size={22} color="#10b981" />,
    label: 'Completed',
    color: '#10b981',
    barColor: 'linear-gradient(90deg, #10b981, #06b6d4)',
  },
  FAILURE: {
    icon: <XCircle size={22} color="#ef4444" />,
    label: 'Failed',
    color: '#ef4444',
    barColor: 'linear-gradient(90deg, #ef4444, #b91c1c)',
  },
  REVOKED: {
    icon: <XCircle size={22} color="#94a3b8" />,
    label: 'Cancelled',
    color: '#94a3b8',
    barColor: 'linear-gradient(90deg, #64748b, #94a3b8)',
  },
};

export default function TaskStatusCard({
  taskId, status, message, progress, downloadToken, error, pollError, onReset,
}) {
  const cfg = STATUS_CONFIG[status] || STATUS_CONFIG.PENDING;
  const [downloading, setDownloading] = useState(false);
  const [downloadError, setDownloadError] = useState('');

  const handleDownload = async () => {
    const id = downloadToken || taskId;
    if (!id) return;
    setDownloadError('');
    setDownloading(true);
    try {
      await downloadReport(id);
    } catch (err) {
      let detail = err.response?.data?.detail;
      if (err.response?.data instanceof Blob) {
        try {
          const parsed = JSON.parse(await err.response.data.text());
          detail = parsed.detail;
        } catch {
          /* ignore */
        }
      }
      if (typeof detail === 'string') {
        setDownloadError(detail);
      } else if (err.response?.status === 401) {
        setDownloadError('Session expired. Please sign in again.');
      } else {
        setDownloadError('Download failed. Please try again.');
      }
    } finally {
      setDownloading(false);
    }
  };

  return (
    <>
      <style>{`
        @keyframes spin { to { transform: rotate(360deg); } }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        .task-card { animation: fadeIn 0.3s ease; }
        .pulse-bar { animation: pulseBar 1.5s ease-in-out infinite; }
        @keyframes pulseBar { 0%,100% { opacity: 1; } 50% { opacity: 0.7; } }
      `}</style>

      <div className="task-card" style={{
        background: 'var(--color-surface)',
        border: `1px solid ${cfg.color}33`,
        borderRadius: '16px',
        padding: '28px',
        marginTop: '24px',
        boxShadow: `0 0 30px ${cfg.color}18`,
      }}>
        {/* Header */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '20px' }}>
          {cfg.icon}
          <div>
            <div style={{ fontWeight: '700', fontSize: '16px', color: cfg.color }}>{cfg.label}</div>
            <div style={{ fontSize: '12px', color: 'var(--color-text-muted)', marginTop: '2px' }}>
              Task: <code style={{ fontFamily: 'monospace' }}>{taskId?.slice(0, 8)}…</code>
            </div>
          </div>
        </div>

        {/* Progress bar */}
        <div style={{
          background: 'var(--color-surface-2)',
          borderRadius: '99px',
          height: '8px',
          overflow: 'hidden',
          marginBottom: '12px',
        }}>
          <div
            className={status === 'STARTED' ? 'pulse-bar' : ''}
            style={{
              width: `${progress ?? 0}%`,
              height: '100%',
              background: cfg.barColor,
              borderRadius: '99px',
              transition: 'width 0.5s ease',
            }}
          />
        </div>

        {/* Message */}
        <p style={{ fontSize: '14px', color: 'var(--color-text-muted)', margin: '0 0 16px' }}>
          {message || 'Waiting for status update…'}
        </p>

        {pollError && status !== 'FAILURE' && status !== 'SUCCESS' && (
          <div style={{
            background: 'rgba(245,158,11,0.1)',
            border: '1px solid rgba(245,158,11,0.35)',
            borderRadius: '8px',
            padding: '10px 14px',
            fontSize: '13px',
            color: '#fcd34d',
            marginBottom: '16px',
          }}>
            {pollError}
          </div>
        )}

        {/* Error */}
        {error && (
          <div style={{
            background: 'rgba(239,68,68,0.1)',
            border: '1px solid rgba(239,68,68,0.3)',
            borderRadius: '8px',
            padding: '10px 14px',
            fontSize: '13px',
            color: '#fca5a5',
            marginBottom: '16px',
            fontFamily: 'monospace',
          }}>
            {error}
          </div>
        )}

        {downloadError && (
          <div style={{
            background: 'rgba(239,68,68,0.1)',
            border: '1px solid rgba(239,68,68,0.3)',
            borderRadius: '8px',
            padding: '10px 14px',
            fontSize: '13px',
            color: '#fca5a5',
            marginBottom: '16px',
          }}>
            {downloadError}
          </div>
        )}

        {/* Actions */}
        <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
          {status === 'SUCCESS' && (downloadToken || taskId) && (
            <button
              type="button"
              onClick={handleDownload}
              disabled={downloading}
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: '7px',
                background: 'linear-gradient(135deg, #10b981, #06b6d4)',
                color: 'white',
                border: 'none',
                borderRadius: '10px',
                padding: '10px 20px',
                fontWeight: '600',
                fontSize: '14px',
                cursor: downloading ? 'wait' : 'pointer',
                opacity: downloading ? 0.7 : 1,
              }}
            >
              {downloading ? (
                <Loader2 size={15} style={{ animation: 'spin 1s linear infinite' }} />
              ) : (
                <Download size={15} />
              )}
              {downloading ? 'Downloading…' : 'Download Report'}
            </button>
          )}
          {(status === 'SUCCESS' || status === 'FAILURE' || status === 'REVOKED') && (
            <button
              onClick={onReset}
              style={{
                background: 'var(--color-surface-2)',
                color: 'var(--color-text-muted)',
                border: '1px solid var(--color-border)',
                borderRadius: '10px',
                padding: '10px 20px',
                fontWeight: '600',
                fontSize: '14px',
                cursor: 'pointer',
                transition: 'all 0.15s',
              }}
            >
              Generate Another
            </button>
          )}
        </div>
      </div>
    </>
  );
}
