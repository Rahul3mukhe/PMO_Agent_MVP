import React from 'react';
import type { PMOState } from '../types';
import GlassCard from './GlassCard';
import { Target, FileCheck, Wallet, Download, CheckCircle, AlertCircle } from 'lucide-react';

import axios from 'axios';

interface DashboardProps {
  result: PMOState;
}

const API_BASE = window.location.origin;

const Dashboard: React.FC<DashboardProps> = ({ result }) => {
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
          <button 
            onClick={handleDownloadReport}
            className="btn-primary" 
            style={{ 
              display: 'flex', alignItems: 'center', gap: '0.5rem', 
              padding: '0.5rem 1rem', fontSize: '0.85rem', marginBottom: '0.5rem' 
            }}
          >
            <Download size={16} /> Export Decision Report (DOCX)
          </button>
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

      <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: '2rem' }}>
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
        </div>
      </div>
    </div>
  );
};

export default Dashboard;
