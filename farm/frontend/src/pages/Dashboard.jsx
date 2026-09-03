/**
 * pages/Dashboard.jsx — Report generation form with real-time task polling.
 */

import { useState, useEffect, useRef } from 'react';
import { useAuth } from '../context/AuthContext';
import client from '../api/client';
import TaskStatusCard from '../components/TaskStatusCard';
import {
  FileSpreadsheet, Upload, ChevronDown, AlertTriangle, Zap,
  FileText, Table2, BookOpen, Layers, Star, X
} from 'lucide-react';

const POLL_INTERVAL_MS = 2500;
const MAX_POLL_ERRORS = 5;

const SEMESTER_OPTIONS = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII'];
const FORMAT_OPTIONS = [
  { value: '1', label: 'Format 1 — Assessment of learning levels', icon: <FileText size={14} /> },
  { value: '2', label: 'Format 2 — Performance / improvement report', icon: <BookOpen size={14} /> },
  { value: '3', label: 'Format 3 — Tabular summary (all students)', icon: <Table2 size={14} /> },
  { value: '4', label: 'Combined Format 1 & 2', icon: <Layers size={14} /> },
  { value: '5', label: 'All Formats (2 files: Combined + Summary)', icon: <Star size={14} /> },
];

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

function SectionTitle({ children }) {
  return (
    <div style={{
      fontSize: '11px',
      fontWeight: '700',
      letterSpacing: '0.08em',
      textTransform: 'uppercase',
      color: 'var(--color-text-muted)',
      marginBottom: '14px',
      paddingBottom: '10px',
      borderBottom: '1px solid var(--color-border)',
    }}>
      {children}
    </div>
  );
}

function Label({ children, required }) {
  return (
    <label style={{ display: 'block', fontSize: '12px', fontWeight: '600', color: 'var(--color-text-muted)', marginBottom: '6px' }}>
      {children}{required && <span style={{ color: 'var(--color-danger)', marginLeft: 3 }}>*</span>}
    </label>
  );
}

const inputCls = {
  width: '100%',
  background: 'var(--color-surface-2)',
  border: '1px solid var(--color-border)',
  borderRadius: '8px',
  padding: '9px 12px',
  color: 'var(--color-text)',
  fontSize: '14px',
  outline: 'none',
  transition: 'border-color 0.15s',
  boxSizing: 'border-box',
  marginBottom: '14px',
};

function Select({ label, required, children, ...props }) {
  return (
    <div>
      <Label required={required}>{label}</Label>
      <div style={{ position: 'relative', marginBottom: '14px' }}>
        <select
          {...props}
          style={{
            ...inputCls,
            marginBottom: 0,
            appearance: 'none',
            paddingRight: '32px',
            cursor: 'pointer',
          }}
          onFocus={e => e.target.style.borderColor = 'var(--color-primary)'}
          onBlur={e => e.target.style.borderColor = 'var(--color-border)'}
        >
          {children}
        </select>
        <ChevronDown size={14} style={{ position: 'absolute', right: 10, top: '50%', transform: 'translateY(-50%)', color: 'var(--color-text-muted)', pointerEvents: 'none' }} />
      </div>
    </div>
  );
}

function FileInput({ label, accept, name, required, file, setFile, hint }) {
  return (
    <div style={{ marginBottom: '14px' }}>
      <Label required={required}>{label}</Label>
      <label style={{
        display: 'flex',
        alignItems: 'center',
        gap: '10px',
        background: 'var(--color-surface-2)',
        border: `1px dashed ${file ? 'var(--color-success)' : 'var(--color-border)'}`,
        borderRadius: '8px',
        padding: '10px 14px',
        cursor: 'pointer',
        transition: 'border-color 0.15s',
      }}>
        <Upload size={14} color={file ? 'var(--color-success)' : 'var(--color-text-muted)'} />
        <span style={{ fontSize: '13px', color: file ? 'var(--color-success)' : 'var(--color-text-muted)' }}>
          {file ? file.name : hint || `Choose ${label.toLowerCase()}…`}
        </span>
        <input type="file" name={name} accept={accept} style={{ display: 'none' }} onChange={e => setFile(e.target.files[0])} />
      </label>
    </div>
  );
}

function TypedMultiFileInput({ proofs, setProofs, maxFiles = 5 }) {
  const [currentType, setCurrentType] = useState('Attendance for extra classes');

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (proofs.length >= maxFiles) {
      alert(`You can only attach a maximum of ${maxFiles} files.`);
      return;
    }
    setProofs([...proofs, { type: currentType, file }]);
    e.target.value = '';
  };

  const removeProof = (index) => {
    setProofs(proofs.filter((_, i) => i !== index));
  };

  return (
    <div style={{ marginBottom: '14px' }}>
      <Label>Proof of Remediation {proofs.length > 0 && `(${proofs.length}/${maxFiles})`}</Label>
      
      <div style={{ display: 'flex', gap: '10px', marginBottom: '10px' }}>
        <div style={{ flex: 1, position: 'relative' }}>
          <select
            value={currentType}
            onChange={(e) => setCurrentType(e.target.value)}
            style={{
              width: '100%', background: 'var(--color-surface-2)', border: '1px solid var(--color-border)',
              borderRadius: '8px', padding: '9px 12px', color: 'var(--color-text)', fontSize: '14px',
              outline: 'none', appearance: 'none', paddingRight: '32px', cursor: 'pointer',
            }}
          >
            <option value="Attendance for extra classes">Attendance for extra classes</option>
            <option value="Assignments">Assignments</option>
            <option value="Others">Others</option>
          </select>
          <ChevronDown size={14} style={{ position: 'absolute', right: 10, top: '50%', transform: 'translateY(-50%)', color: 'var(--color-text-muted)', pointerEvents: 'none' }} />
        </div>
        
        <label style={{
          display: 'flex', alignItems: 'center', gap: '6px',
          background: 'var(--color-surface-2)', border: '1px dashed var(--color-primary)',
          borderRadius: '8px', padding: '0 14px', cursor: 'pointer',
          color: 'var(--color-primary)', fontSize: '13px', fontWeight: '500',
        }}>
          <Upload size={14} /> Attach
          <input type="file" accept=".pdf,.png,.jpg,.jpeg" style={{ display: 'none' }} onChange={handleFileChange} />
        </label>
      </div>

      {proofs.length > 0 && (
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
          {proofs.map((proof, idx) => (
            <div key={idx} style={{
              display: 'flex', alignItems: 'center', gap: '6px',
              background: 'var(--color-surface)', border: '1px solid var(--color-border)',
              borderRadius: '999px', padding: '4px 10px', fontSize: '12px',
              color: 'var(--color-text)'
            }}>
              <span style={{ fontWeight: 600, color: 'var(--color-primary)' }}>{proof.type}:</span>
              <span style={{ maxWidth: '100px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }} title={proof.file.name}>
                {proof.file.name}
              </span>
              <X size={12} style={{ cursor: 'pointer', color: 'var(--color-text-muted)' }} onClick={() => removeProof(idx)} />
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function Toggle({ label, checked, onChange }) {
  return (
    <label style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer', marginBottom: '14px' }}>
      <div
        onClick={() => onChange(!checked)}
        style={{
          width: 40, height: 22, borderRadius: 99,
          background: checked ? 'var(--color-primary)' : 'var(--color-surface-2)',
          border: `1px solid ${checked ? 'var(--color-primary)' : 'var(--color-border)'}`,
          position: 'relative', transition: 'all 0.2s', flexShrink: 0,
        }}
      >
        <div style={{
          width: 16, height: 16, borderRadius: '50%', background: 'white',
          position: 'absolute', top: 2, left: checked ? 20 : 2, transition: 'left 0.2s',
        }} />
      </div>
      <span style={{ fontSize: '14px', fontWeight: '500' }}>{label}</span>
    </label>
  );
}

export default function Dashboard() {
  const { faculty } = useAuth();

  // Form state
  const [excelFile, setExcelFile] = useState(null);
  const [cgpaFile, setCgpaFile] = useState(null);
  const [gradeFile, setGradeFile] = useState(null);
  const [semester, setSemester] = useState('III');
  const [learnerType, setLearnerType] = useState('slow');
  const [formatChoice, setFormatChoice] = useState('4');
  const [outputType, setOutputType] = useState('pdf');
  const [slowThreshold, setSlowThreshold] = useState(40);
  const [advancedThreshold, setAdvancedThreshold] = useState(90);
  const [comment, setComment] = useState('');
  const [enableSigning, setEnableSigning] = useState(false);
  const [proofs, setProofs] = useState([]);

  const [showAllReports, setShowAllReports] = useState(false);

  // Task state
  const [taskId, setTaskId] = useState(null);
  const [taskData, setTaskData] = useState(null);
  const [submitting, setSubmitting] = useState(false);
  const [formError, setFormError] = useState('');
  const [pollError, setPollError] = useState('');
  const pollRef = useRef(null);
  const pollFailCountRef = useRef(0);

  // Signing warning
  const signingMissing = enableSigning && (!faculty?.has_private_key || !faculty?.has_certificate);

  // Polling
  useEffect(() => {
    if (!taskId) return;
    pollFailCountRef.current = 0;
    setPollError('');

    const pollStatusMessage = (err) => {
      const status = err.response?.status;
      const detail = err.response?.data?.detail;
      if (status === 403) return typeof detail === 'string' ? detail : 'You do not have access to this task.';
      if (status === 404) return typeof detail === 'string' ? detail : 'Task not found.';
      if (!err.response) return 'Lost connection while checking report status.';
      return 'Could not refresh report status. Please try again.';
    };

    const stopPollingAsFailed = (message) => {
      clearInterval(pollRef.current);
      setPollError(message);
      setTaskData((prev) => ({
        ...(prev || {}),
        status: 'FAILURE',
        progress: 0,
        message,
        error: message,
      }));
    };

    const poll = async () => {
      try {
        const r = await client.get(`/task-status/${taskId}`);
        pollFailCountRef.current = 0;
        setPollError('');
        setTaskData(r.data);
        if (['SUCCESS', 'FAILURE', 'REVOKED'].includes(r.data.status)) {
          clearInterval(pollRef.current);
        }
      } catch (err) {
        const message = pollStatusMessage(err);
        pollFailCountRef.current += 1;
        setPollError(message);
        if (
          pollFailCountRef.current >= MAX_POLL_ERRORS
          || err.response?.status === 403
          || err.response?.status === 404
        ) {
          stopPollingAsFailed(message);
        }
      }
    };
    poll();
    pollRef.current = setInterval(poll, POLL_INTERVAL_MS);
    return () => clearInterval(pollRef.current);
  }, [taskId]);

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!excelFile) { setFormError('Please select an Excel file.'); return; }
    setFormError('');
    setSubmitting(true);

    const fd = new FormData();
    fd.append('excel_file', excelFile);
    if (cgpaFile) fd.append('cgpa_file', cgpaFile);
    if (gradeFile) fd.append('grade_file', gradeFile);
    fd.append('semester', semester);
    fd.append('learner_type', learnerType);
    fd.append('format_choice', formatChoice);
    fd.append('output_type', outputType);
    fd.append('slow_threshold', slowThreshold);
    fd.append('advanced_threshold', advancedThreshold);
    fd.append('common_comment', comment);
    fd.append('faculty_name', faculty?.name || '');
    fd.append('enable_signing', enableSigning ? 'true' : 'false');
    
    proofs.forEach(proof => {
      fd.append('proof_types', proof.type);
      fd.append('proof_files', proof.file);
    });

    try {
      const r = await client.post('/generate-report', fd, {
        headers: { 'Content-Type': 'multipart/form-data' },
      });
      setTaskId(r.data.task_id);
      setPollError('');
      pollFailCountRef.current = 0;
      setTaskData({ status: 'PENDING', progress: 0, message: 'Report queued.' });
    } catch (err) {
      setFormError(err.response?.data?.detail || 'Submission failed. Please try again.');
    } finally {
      setSubmitting(false);
    }
  };

  const reset = () => {
    setTaskId(null);
    setTaskData(null);
    setPollError('');
    pollFailCountRef.current = 0;
    setExcelFile(null);
    setCgpaFile(null);
    setGradeFile(null);
    setProofs([]);
    setFormError('');
  };

  return (
    <div style={{ maxWidth: 900, margin: '0 auto', padding: '32px 20px' }}>
      {/* Page header */}
      <div style={{ marginBottom: '28px' }}>
        <h1 style={{ margin: 0, fontSize: '26px', fontWeight: '800', letterSpacing: '-0.5px' }}>
          Report Dashboard
        </h1>
        <p style={{ margin: '6px 0 0', color: 'var(--color-text-muted)', fontSize: '14px' }}>
          Upload your mid-term Excel file and configure report parameters.
        </p>
      </div>

      {/* Task card (while task is running) */}
      {taskData && (
        <TaskStatusCard
          taskId={taskId}
          {...taskData}
          pollError={pollError}
          onReset={reset}
        />
      )}

      {/* Form (hidden while task is running) */}
      {!taskId && (
        <form onSubmit={handleSubmit}>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>

            {/* ── LEFT COLUMN ── */}
            <div>
              <Card>
                <SectionTitle>Upload Files</SectionTitle>
                <FileInput
                  label="Mid-Term Marks Excel"
                  accept=".xls,.xlsx"
                  name="excel_file"
                  required
                  file={excelFile}
                  setFile={setExcelFile}
                  hint="Choose .xls or .xlsx…"
                />
                <FileInput
                  label="CGPA File"
                  accept=".xls,.xlsx,.csv"
                  name="cgpa_file"
                  file={cgpaFile}
                  setFile={setCgpaFile}
                  hint="CGPA data for previous semester…"
                />
                <FileInput
                  label="Grade Sheet"
                  accept=".xls,.xlsx,.csv"
                  name="grade_file"
                  file={gradeFile}
                  setFile={setGradeFile}
                  hint="End-semester grade data…"
                />
              </Card>

              {/* <Card style={{ marginTop: 20 }}>
                <SectionTitle>Signing Options</SectionTitle>
                <Toggle
                  label="Enable Digital Signature"
                  checked={enableSigning}
                  onChange={setEnableSigning}
                />
                {enableSigning && signingMissing && (
                  <div style={{
                    display: 'flex', alignItems: 'flex-start', gap: 10,
                    background: 'rgba(245,158,11,0.1)',
                    border: '1px solid rgba(245,158,11,0.3)',
                    borderRadius: 8, padding: '10px 14px', fontSize: 13, color: 'var(--color-warning)',
                  }}>
                    <AlertTriangle size={15} style={{ flexShrink: 0, marginTop: 1 }} />
                    <span>
                      Private key or certificate not found in your profile.{' '}
                      <a href="/profile" style={{ color: 'var(--color-warning)', fontWeight: 600 }}>Upload them →</a>
                    </span>
                  </div>
                )}
                {enableSigning && !signingMissing && (
                  <div style={{ fontSize: 13, color: 'var(--color-success)' }}>
                    ✅ Signing credentials loaded from secure vault.
                  </div>
                )}
                {!enableSigning && (
                  <p style={{ fontSize: 12, color: 'var(--color-text-muted)', margin: 0 }}>
                    {faculty?.has_signature ? '✅ Signature image stored.' : '⚠️ No signature image. Add one in Profile.'}
                  </p>
                )}
              </Card> */}
            </div>

            {/* ── RIGHT COLUMN ── */}
            <div>
              <Card>
                <SectionTitle>Report Parameters</SectionTitle>

                <Select label="Semester" required value={semester} onChange={e => setSemester(e.target.value)}>
                  {SEMESTER_OPTIONS.map(s => <option key={s} value={s}>{s}</option>)}
                </Select>

                <Select label="Learner Type" required value={learnerType} onChange={e => setLearnerType(e.target.value)}>
                  <option value="slow">Slow Learners</option>
                  <option value="advanced">Advanced Learners</option>
                </Select>

                {learnerType === 'slow' && (
                  <div>
                    <Label>Slow threshold (%)</Label>
                    <input
                      type="number" min={0} max={100} step={1}
                      value={slowThreshold}
                      onChange={e => setSlowThreshold(Number(e.target.value))}
                      style={inputCls}
                      onFocus={e => e.target.style.borderColor = 'var(--color-primary)'}
                      onBlur={e => e.target.style.borderColor = 'var(--color-border)'}
                    />
                  </div>
                )}

                {learnerType === 'advanced' && (
                  <div>
                    <Label>Advanced threshold (%)</Label>
                    <input
                      type="number" min={0} max={100} step={1}
                      value={advancedThreshold}
                      onChange={e => setAdvancedThreshold(Number(e.target.value))}
                      style={inputCls}
                      onFocus={e => e.target.style.borderColor = 'var(--color-primary)'}
                      onBlur={e => e.target.style.borderColor = 'var(--color-border)'}
                    />
                  </div>
                )}

                <Select label="Report Format" required value={formatChoice} onChange={e => setFormatChoice(e.target.value)}>
                  {FORMAT_OPTIONS.map(f => (
                    <option key={f.value} value={f.value}>{f.label}</option>
                  ))}
                </Select>

                <Select label="Output Type" required value={outputType} onChange={e => setOutputType(e.target.value)}>
                  <option value="pdf">PDF</option>
                  <option value="word">Word (.docx)</option>
                </Select>

                <div>
                  <Label>Common Action / Comment</Label>
                  <textarea
                    value={comment}
                    onChange={e => setComment(e.target.value)}
                    rows={3}
                    placeholder="e.g. Remedial classes conducted; Extra assignments given…"
                    style={{ ...inputCls, resize: 'vertical', fontFamily: 'inherit' }}
                    onFocus={e => e.target.style.borderColor = 'var(--color-primary)'}
                    onBlur={e => e.target.style.borderColor = 'var(--color-border)'}
                  />
                </div>

                <TypedMultiFileInput
                  proofs={proofs}
                  setProofs={setProofs}
                  maxFiles={5}
                />
              </Card>
            </div>
          </div>

          {/* Error */}
          {formError && (
            <div style={{
              display: 'flex', alignItems: 'center', gap: 10,
              background: 'rgba(239,68,68,0.1)', border: '1px solid rgba(239,68,68,0.3)',
              borderRadius: 10, padding: '12px 16px', fontSize: 14, color: 'var(--color-danger)',
              marginTop: 16,
            }}>
              <AlertTriangle size={15} />
              {formError}
            </div>
          )}

          {/* Submit */}
          <button
            type="submit"
            disabled={submitting || (enableSigning && signingMissing)}
            style={{
              marginTop: '20px',
              width: '100%',
              background: submitting ? 'var(--color-surface-2)' : 'linear-gradient(135deg, var(--color-primary), var(--color-primary-hover))',
              color: submitting ? 'var(--color-text-muted)' : 'white',
              border: 'none',
              borderRadius: '12px',
              padding: '15px',
              fontWeight: '700',
              fontSize: '15px',
              cursor: submitting ? 'not-allowed' : 'pointer',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              gap: '8px',
              transition: 'all 0.15s',
              boxShadow: submitting ? 'none' : '0 4px 20px rgba(124,58,237,0.35)',
            }}
          >
            <Zap size={16} />
            {submitting ? 'Submitting…' : 'Generate Report'}
          </button>
        </form>
      )}

    </div>
  );
}
