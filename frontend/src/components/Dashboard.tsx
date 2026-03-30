import React, { useState } from 'react';
import type { PMOState, GenerationLogEntry } from '../types';
import GlassCard from './GlassCard';
import { Target, FileCheck, Wallet, Download, CheckCircle, AlertCircle, Zap, AlertTriangle, Server } from 'lucide-react';

import axios from 'axios';

interface DashboardProps {
  result: PMOState;
}

const API_BASE = import.meta.env.DEV ? 'http://127.0.0.1:8000' : window.location.origin;

const Dashboard: React.FC<DashboardProps> = ({ result }) => {
  const [isExportingPPTX, setIsExportingPPTX] = useState(false);
  const { project, decision, docs, gates, summary } = result;
  
  const isApprove = decision?.toUpperCase().includes('APPROV') || decision?.toUpperCase().includes('PASS');
  const isReject = decision?.toUpperCase().includes('REJECT') || decision?.toUpperCase().includes('FAIL');
  
  const statusColor = isApprove ? '#10b981' : isReject ? '#ef4444' : '#f59e0b';
  
  const docStats = Object.values(docs).reduce((acc, d) => {
    if (d.status === 'SUFFICIENT') acc.ok++;
    acc.total++;
    return acc;
  }, { ok: 0, total: 0 });

  const downloadFile = (blob: Blob, filename: string) => {
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', filename);
    document.body.appendChild(link);
    link.click();
    link.parentNode?.removeChild(link);
  };

  const handleDownloadMD = (docType: string) => {
    const doc = docs[docType];
    if (!doc || !doc.content_markdown) return;
    const blob = new Blob([doc.content_markdown], { type: 'text/markdown' });
    downloadFile(blob, `${project.project_id}_${docType}.md`);
  };

  const handleDownloadDOCX = async (docType: string) => {
    try {
      const response = await axios.post(`${API_BASE}/export/docx?doc_type=${docType}`, result, {
        responseType: 'blob'
      });
      downloadFile(response.data, `${project.project_id}_${docType}.docx`);
    } catch (err) {
      console.error("DOCX export failed", err);
      alert("Failed to export DOCX. please check server.");
    }
  };

  const handleDownloadPPTX = async () => {
    setIsExportingPPTX(true);
    try {
      const response = await axios.post(`${API_BASE}/export/pptx`, result, {
        responseType: 'blob'
      });
      downloadFile(response.data, `${project.project_id}_Client_Status.pptx`);
    } catch (err) {
      console.error("PPTX export failed", err);
      alert("Failed to export PPTX. Please check server logs.");
    } finally {
      setIsExportingPPTX(false);
    }
  };

  const handleDownloadReport = async () => {
    try {
      const response = await axios.post(`${API_BASE}/export/report`, result, {
        responseType: 'blob'
      });
      downloadFile(response.data, `${project.project_id}_PMO_Decision_Report.docx`);
    } catch (err) {
      console.error("Report export failed", err);
      alert("Failed to export Report. please check server.");
    }
  };

  return (
    <div className="fade-in" style={{ display: 'grid', gap: '2rem' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-end', marginBottom: '1rem' }}>
        <div>
          <h1 style={{ fontSize: '2rem', marginBottom: '0.5rem' }}>Analysis Results</h1>
          <p style={{ color: 'var(--text-muted)' }}>Project: <strong>{project.project_name}</strong> ({project.project_id})</p>
        </div>
        <div style={{ textAlign: 'right', display: 'flex', flexDirection: 'column', gap: '0.5rem', alignItems: 'flex-end' }}>
          <div style={{ display: 'flex', gap: '0.5rem', marginBottom: '0.5rem', flexWrap: 'wrap', justifyContent: 'flex-end' }}>
            <button 
              onClick={handleDownloadPPTX}
              className="btn-secondary" 
              disabled={isExportingPPTX}
              style={{ 
                display: 'flex', alignItems: 'center', gap: '0.5rem', 
                padding: '0.5rem 1rem', fontSize: '0.85rem' 
              }}
            >
              <Download size={16} /> {isExportingPPTX ? 'Generating PPTX...' : 'Export Client Status (PPTX)'}
            </button>
            <button 
              onClick={handleDownloadReport}
              className="btn-primary" 
              style={{ 
                display: 'flex', alignItems: 'center', gap: '0.5rem', 
                padding: '0.5rem 1rem', fontSize: '0.85rem' 
              }}
            >
              <Download size={16} /> Export Decision Report (DOCX)
            </button>
          </div>
          <div>
            <span style={{ fontSize: '0.75rem', textTransform: 'uppercase', color: 'var(--text-muted)', letterSpacing: '0.1em' }}>Verdict</span>
            <div style={{ fontSize: '1.5rem', fontWeight: '800', color: statusColor }}>{decision || 'PENDING'}</div>
          </div>
        </div>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '1.25rem' }}>
        <GlassCard>
            <div style={{ display: 'flex', gap: '1rem', alignItems: 'center' }}>
                <div style={{ padding: '0.75rem', background: 'rgba(59,130,246,0.1)', borderRadius: '12px' }}><Target color="#3b82f6" /></div>
                <div>
                    <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)', textTransform: 'uppercase' }}>Documents</div>
                    <div style={{ fontSize: '1.5rem', fontWeight: '700' }}>{docStats.ok}/{docStats.total}</div>
                </div>
            </div>
        </GlassCard>
        <GlassCard>
            <div style={{ display: 'flex', gap: '1rem', alignItems: 'center' }}>
                <div style={{ padding: '0.75rem', background: 'rgba(16,185,129,0.1)', borderRadius: '12px' }}><FileCheck color="#10b981" /></div>
                <div>
                    <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)', textTransform: 'uppercase' }}>Pass Rate</div>
                    <div style={{ fontSize: '1.5rem', fontWeight: '700' }}>{Math.round((docStats.ok/docStats.total)*100)}%</div>
                </div>
            </div>
        </GlassCard>
        <GlassCard>
            <div style={{ display: 'flex', gap: '1rem', alignItems: 'center' }}>
                <div style={{ padding: '0.75rem', background: 'rgba(245,158,11,0.1)', borderRadius: '12px' }}><Wallet color="#f59e0b" /></div>
                <div>
                    <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)', textTransform: 'uppercase' }}>Actual Spend</div>
                    <div style={{ fontSize: '1.5rem', fontWeight: '700' }}>${project.actual_budget_consumed?.toLocaleString()}</div>
                </div>
            </div>
        </GlassCard>
      </div>

      {/* Generation Status Banner */}
      {result.audit?.generation_log && result.audit.generation_log.length > 0 && (() => {
        const genLog = result.audit.generation_log as GenerationLogEntry[];
        const hasLocalTemplate = genLog.some((e: GenerationLogEntry) => e.provider === 'local_template' || e.provider === 'Local Template');
        const llmEntry = genLog.find((e: GenerationLogEntry) => e.status === 'ok' && e.provider !== 'local_template' && e.provider !== 'Local Template');
        const llmDesc = llmEntry ? ('Documents generated by ' + llmEntry.provider + ' (' + llmEntry.model + '). AI reasoning is active.') : '';
        return (
          <div style={{
            padding: '1rem 1.25rem',
            borderRadius: '10px',
            border: '1px solid ' + (hasLocalTemplate ? '#f59e0b' : '#10b981'),
            background: hasLocalTemplate ? 'rgba(245,158,11,0.07)' : 'rgba(16,185,129,0.07)',
            display: 'flex', alignItems: 'center', gap: '0.75rem',
            marginBottom: '0.5rem'
          }}>
            {hasLocalTemplate
              ? <AlertTriangle size={20} color="#f59e0b" />
              : <Zap size={20} color="#10b981" />}
            <div>
              <div style={{ fontWeight: '700', fontSize: '0.9rem', color: hasLocalTemplate ? '#92400e' : '#065f46' }}>
                {hasLocalTemplate ? 'Local Template Mode (Offline)' : 'LLM Generation Active'}
              </div>
              <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)', marginTop: '2px' }}>
                {hasLocalTemplate
                  ? 'Documents were generated using the offline Local Template (LLM provider unreachable). Content is based on your form inputs only, not AI reasoning.'
                  : llmDesc}
              </div>
            </div>
          </div>
        );
      })()}

      <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: '2rem' }}>
        <div style={{ display: 'grid', gap: '1.5rem' }}>
        <GlassCard title="Document Status">
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.9rem' }}>
            <thead>
              <tr style={{ textAlign: 'left', borderBottom: '1px solid #e2e8f0', color: 'var(--text-muted)' }}>
                <th style={{ padding: '0.75rem 0' }}>Document</th>
                <th style={{ padding: '0.75rem 0' }}>Status</th>
                <th style={{ padding: '0.75rem 0', textAlign: 'right' }}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {Object.entries(docs).map(([key, d]) => (
                <tr key={key} style={{ borderBottom: '1px solid #f1f5f9' }}>
                  <td style={{ padding: '1rem 0', fontWeight: '600' }}>{d.title}</td>
                  <td style={{ padding: '1rem 0' }}>
                    <span style={{ 
                      padding: '0.25rem 0.6rem', 
                      borderRadius: '4px', 
                      fontSize: '0.75rem', 
                      background: d.status === 'SUFFICIENT' ? '#d1fae5' : '#fee2e2',
                      color: d.status === 'SUFFICIENT' ? '#065f46' : '#991b1b'
                    }}>{d.status}</span>
                  </td>
                  <td style={{ padding: '1rem 0', textAlign: 'right', display: 'flex', gap: '0.5rem', justifyContent: 'flex-end' }}>
                    <button 
                      onClick={() => handleDownloadDOCX(key)}
                      className="btn-secondary" 
                      title="Download DOCX"
                      style={{ padding: '0.4rem 0.8rem', fontSize: '0.8rem', display: 'flex', alignItems: 'center', gap: '0.3rem' }}
                    >
                      <Download size={14} /> DOCX
                    </button>
                    <button 
                      onClick={() => handleDownloadMD(key)}
                      className="btn-secondary" 
                      title="Download Markdown"
                      style={{ padding: '0.4rem 0.8rem', fontSize: '0.8rem', display: 'flex', alignItems: 'center', gap: '0.3rem' }}
                    >
                      <Download size={14} /> MD
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </GlassCard>

        {/* Generation Log Detail Table */}
        {result.audit?.generation_log && result.audit.generation_log.length > 0 && (
          <GlassCard title="AI Generation Log">
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.82rem' }}>
              <thead>
                <tr style={{ textAlign: 'left', borderBottom: '1px solid #e2e8f0', color: 'var(--text-muted)' }}>
                  <th style={{ padding: '0.5rem 0' }}>Document</th>
                  <th style={{ padding: '0.5rem 0' }}>Provider</th>
                  <th style={{ padding: '0.5rem 0' }}>Model</th>
                  <th style={{ padding: '0.5rem 0' }}>Status</th>
                </tr>
              </thead>
              <tbody>
                {(result.audit.generation_log as GenerationLogEntry[]).map((entry, i) => {
                  const isLLM = entry.status === 'ok' && entry.provider !== 'local_template' && entry.provider !== 'Local Template';
                  const isTemplate = entry.provider === 'local_template' || entry.provider === 'Local Template';
                  const statusColor = isLLM ? '#065f46' : isTemplate ? '#92400e' : '#991b1b';
                  const statusBg   = isLLM ? '#d1fae5' : isTemplate ? '#fef3c7' : '#fee2e2';
                  const icon = isLLM ? <Zap size={12} /> : isTemplate ? <Server size={12} /> : <AlertTriangle size={12} />;
                  return (
                    <tr key={i} style={{ borderBottom: '1px solid #f1f5f9' }}>
                      <td style={{ padding: '0.6rem 0', fontWeight: '500' }}>{entry.doc || '-'}</td>
                      <td style={{ padding: '0.6rem 0', color: 'var(--text-muted)' }}>{entry.provider}</td>
                      <td style={{ padding: '0.6rem 0', color: 'var(--text-muted)', fontFamily: 'monospace', fontSize: '0.75rem' }}>{entry.model}</td>
                      <td style={{ padding: '0.6rem 0' }}>
                        <span style={{
                          display: 'inline-flex', alignItems: 'center', gap: '0.3rem',
                          padding: '0.2rem 0.55rem', borderRadius: '4px', fontSize: '0.72rem', fontWeight: '600',
                          background: statusBg, color: statusColor
                        }}>
                          {icon}
                          {isLLM ? 'LLM (ok)' : isTemplate ? 'Local Template' : entry.status}
                        </span>
                        {entry.note && <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)', marginTop: '2px' }}>{entry.note.slice(0, 80)}</div>}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </GlassCard>
        )}
        </div>

        <div style={{ display: 'grid', gap: '1.5rem' }}>
          <GlassCard title="Decision Summary">
            <p style={{ fontSize: '0.9rem', color: 'var(--text-muted)', lineHeight: '1.6' }}>{summary || 'No summary available.'}</p>
          </GlassCard>
          
          <GlassCard title="Gate Checks">
            <div style={{ display: 'grid', gap: '0.75rem' }}>
              {gates.map((g, i) => (
                <div key={i} style={{ display: 'flex', gap: '0.75rem', alignItems: 'flex-start' }}>
                  {g.passed ? <CheckCircle size={18} color="#10b981" /> : <AlertCircle size={18} color="#ef4444" />}
                  <div>
                    <div style={{ fontSize: '0.85rem', fontWeight: '600' }}>{g.gate}</div>
                    <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>{g.findings[0] || 'No issues found.'}</div>
                  </div>
                </div>
              ))}
            </div>
          </GlassCard>

          <GlassCard title="Est. Budget Breakdown">
            <div style={{ fontSize: '0.85rem', color: 'var(--text-main)', display: 'grid', gap: '0.5rem' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                <span>Labour:</span> <strong>${project.labour_cost?.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) || '0.00'}</strong>
              </div>
              <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                <span>Development:</span> <strong>${project.development_cost?.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) || '0.00'}</strong>
              </div>
              <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                <span>Testing:</span> <strong>${project.test_cost?.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) || '0.00'}</strong>
              </div>
              <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                <span>Software:</span> <strong>${project.software_cost?.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) || '0.00'}</strong>
              </div>
              <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                <span>Infrastructure:</span> <strong>${project.infrastructure_cost?.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) || '0.00'}</strong>
              </div>
              <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                <span>Overhead (15%):</span> <strong>${project.overhead_cost?.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) || '0.00'}</strong>
              </div>
              <hr style={{ border: 'none', borderTop: '1px solid var(--border-light)', margin: '0.5rem 0' }} />
              <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: '700', color: 'var(--primary-blue)' }}>
                <span>Total Estimate:</span> <span>${project.estimated_budget?.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) || '0.00'}</span>
              </div>
            </div>
          </GlassCard>
        </div>
      </div>
    </div>
  );
};

export default Dashboard;
