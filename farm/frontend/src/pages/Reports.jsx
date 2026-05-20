import { useState, useEffect } from 'react';
import client, { downloadReport, downloadProofs } from '../api/client';
import { History, Download, RefreshCw } from 'lucide-react';

function Card({ children, style }) {
  return (
    <div style={{
      background: 'var(--color-surface)',
      border: '1px solid var(--color-border)',
      borderRadius: '16px',
      padding: '24px',
      ...style,
    }}>
      {children}
    </div>
  );
}

export default function Reports() {
  const [history, setHistory] = useState([]);
  const [loadingHistory, setLoadingHistory] = useState(false);
  const [showAllReports, setShowAllReports] = useState(false);

  const fetchHistory = async () => {
    setLoadingHistory(true);
    try {
      const r = await client.get('/reports');
      setHistory(r.data.reports);
    } catch (err) {
      console.error('Failed to fetch history:', err);
    } finally {
      setLoadingHistory(false);
    }
  };

  useEffect(() => {
    fetchHistory();
  }, []);

  return (
    <div style={{ maxWidth: 900, margin: '0 auto', padding: '32px 20px' }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '20px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
          <History size={20} color="var(--color-primary)" />
          <h2 style={{ margin: 0, fontSize: '20px', fontWeight: '700' }}>All Reports</h2>
        </div>
        <button
          onClick={fetchHistory}
          disabled={loadingHistory}
          style={{
            display: 'flex', alignItems: 'center', gap: '6px',
            background: 'var(--color-surface)', border: '1px solid var(--color-border)',
            borderRadius: '6px', padding: '6px 12px', color: 'var(--color-text)',
            fontSize: '12px', fontWeight: '600', cursor: 'pointer',
            transition: 'all 0.15s'
          }}
          onMouseOver={e => {
            e.currentTarget.style.borderColor = 'var(--color-primary)';
            e.currentTarget.style.color = 'var(--color-primary)';
          }}
          onMouseOut={e => {
            e.currentTarget.style.borderColor = 'var(--color-border)';
            e.currentTarget.style.color = 'var(--color-text)';
          }}
        >
          <RefreshCw size={14} className={loadingHistory ? 'animate-spin' : ''} />
          Refresh
        </button>
      </div>

      <Card style={{ padding: 0, overflow: 'hidden' }}>
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '14px' }}>
            <thead>
              <tr style={{ borderBottom: '1px solid var(--color-border)', background: 'var(--color-surface-2)' }}>
                <th style={{ padding: '12px 16px', color: 'var(--color-text-muted)', fontWeight: '600' }}>Date</th>
                <th style={{ padding: '12px 16px', color: 'var(--color-text-muted)', fontWeight: '600' }}>Semester</th>
                <th style={{ padding: '12px 16px', color: 'var(--color-text-muted)', fontWeight: '600' }}>Type</th>
                <th style={{ padding: '12px 16px', color: 'var(--color-text-muted)', fontWeight: '600' }}>Format</th>
                <th style={{ padding: '12px 16px', color: 'var(--color-text-muted)', fontWeight: '600' }}>Status</th>
                <th style={{ padding: '12px 16px', color: 'var(--color-text-muted)', fontWeight: '600' }}>Proofs</th>
                <th style={{ padding: '12px 16px', color: 'var(--color-text-muted)', fontWeight: '600', textAlign: 'right' }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {history.length === 0 && !loadingHistory && (
                <tr>
                  <td colSpan="7" style={{ padding: '40px', textAlign: 'center', color: 'var(--color-text-muted)' }}>
                    No reports generated yet.
                  </td>
                </tr>
              )}
              {(showAllReports ? history : history.slice(0, 5)).map((item) => (
                <tr key={item.task_id} style={{ borderBottom: '1px solid var(--color-border)', transition: 'background 0.1s' }} onMouseOver={e => e.currentTarget.style.background = 'var(--color-surface-2)'} onMouseOut={e => e.currentTarget.style.background = 'none'}>
                  <td style={{ padding: '14px 16px', whiteSpace: 'nowrap' }}>
                    {new Date(item.created_at).toLocaleString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }).replace(',', '')}
                  </td>
                  <td style={{ padding: '14px 16px', fontWeight: '500' }}>{item.semester}</td>
                  <td style={{ padding: '14px 16px' }}>
                    <span style={{
                      padding: '2px 8px', borderRadius: '4px', fontSize: '12px', fontWeight: '600',
                      background: item.learner_type === 'slow' ? 'rgba(239,68,68,0.1)' : 'rgba(34,197,94,0.1)',
                      color: item.learner_type === 'slow' ? 'var(--color-danger)' : 'var(--color-success)',
                      textTransform: 'capitalize'
                    }}>
                      {item.learner_type}
                    </span>
                  </td>
                  <td style={{ padding: '14px 16px', color: 'var(--color-text-muted)' }}>
                    Format {item.format_choice} ({item.output_type.toUpperCase()})
                  </td>
                  <td style={{ padding: '14px 16px' }}>
                    <span style={{
                      fontSize: '12px', fontWeight: '700',
                      color: item.status === 'SUCCESS' ? 'var(--color-success)' :
                             item.status === 'FAILURE' ? 'var(--color-danger)' :
                             'var(--color-primary)'
                    }}>
                      {item.status}
                    </span>
                  </td>
                  <td style={{ padding: '14px 16px' }}>
                    {item.has_proofs && item.status === 'SUCCESS' && (
                      <button
                        onClick={() => downloadProofs(item.task_id)}
                        style={{
                          display: 'inline-flex', alignItems: 'center', gap: '6px',
                          background: 'none', border: '1px solid var(--color-border)',
                          borderRadius: '6px', padding: '6px 10px', color: 'var(--color-text)',
                          fontSize: '12px', fontWeight: '600', cursor: 'pointer',
                          transition: 'all 0.15s'
                        }}
                        onMouseOver={e => {
                          e.currentTarget.style.borderColor = 'var(--color-primary)';
                          e.currentTarget.style.color = 'var(--color-primary)';
                        }}
                        onMouseOut={e => {
                          e.currentTarget.style.borderColor = 'var(--color-border)';
                          e.currentTarget.style.color = 'var(--color-text)';
                        }}
                      >
                        <Download size={14} />
                        ZIP
                      </button>
                    )}
                    {!item.has_proofs && <span style={{ color: 'var(--color-text-muted)', fontSize: '13px' }}>-</span>}
                  </td>
                  <td style={{ padding: '14px 16px', textAlign: 'right' }}>
                    {item.status === 'SUCCESS' && (
                      <button
                        onClick={() => downloadReport(item.task_id)}
                        style={{
                          display: 'inline-flex', alignItems: 'center', gap: '6px',
                          background: 'var(--color-surface-2)', border: '1px solid var(--color-border)',
                          borderRadius: '6px', padding: '6px 10px', color: 'var(--color-text)',
                          fontSize: '12px', fontWeight: '600', cursor: 'pointer',
                          transition: 'all 0.15s'
                        }}
                        onMouseOver={e => {
                          e.currentTarget.style.borderColor = 'var(--color-primary)';
                          e.currentTarget.style.color = 'var(--color-primary)';
                        }}
                        onMouseOut={e => {
                          e.currentTarget.style.borderColor = 'var(--color-border)';
                          e.currentTarget.style.color = 'var(--color-text)';
                        }}
                      >
                        <Download size={14} />
                        Download
                      </button>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          
          {history.length > 5 && (
            <div style={{ textAlign: 'center', padding: '12px', background: 'var(--color-surface)' }}>
              <button
                onClick={() => setShowAllReports(!showAllReports)}
                style={{
                  background: 'none', border: 'none', color: 'var(--color-primary)',
                  fontSize: '13px', fontWeight: '600', cursor: 'pointer',
                }}
              >
                {showAllReports ? 'Collapse All' : `View More (${history.length - 5})`}
              </button>
            </div>
          )}
        </div>
      </Card>
    </div>
  );
}
