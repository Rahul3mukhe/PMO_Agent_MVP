import React from 'react';
import GlassCard from './GlassCard';
import { Upload, FileText, X } from 'lucide-react';

interface DocUploadProps {
  title?: string;
  subtitle?: string;
  files: File[];
  onFilesChange: (files: File[]) => void;
  mapping: Record<string, string>;
  onMappingChange: (mapping: Record<string, string>) => void;
  docTypes: Record<string, { title: string }>;
  id: string;
}

const DocUpload: React.FC<DocUploadProps> = ({ 
    title = "Documents", 
    subtitle = "Upload existing files", 
    files, onFilesChange, mapping, onMappingChange, docTypes, id 
}) => {
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      onFilesChange([...files, ...Array.from(e.target.files)]);
    }
  };

  const removeFile = (index: number) => {
    const fileToRemove = files[index];
    onFilesChange(files.filter((_, i) => i !== index));
    const newMapping = { ...mapping };
    delete newMapping[fileToRemove.name];
    onMappingChange(newMapping);
  };

  const updateMapping = (fileName: string, type: string) => {
    onMappingChange({ ...mapping, [fileName]: type });
  };

  return (
    <GlassCard title={title} subtitle={subtitle}>
      <div style={{
        border: '2px dashed rgba(59, 130, 246, 0.2)',
        borderRadius: '12px',
        padding: '2rem',
        textAlign: 'center',
        background: 'rgba(239, 246, 255, 0.4)',
        cursor: 'pointer',
        marginBottom: '1.5rem'
      }} onClick={() => document.getElementById(id)?.click()}>
        <Upload size={32} color="var(--primary-blue)" style={{ marginBottom: '0.75rem' }} />
        <p style={{ fontSize: '0.9rem', color: 'var(--text-main)', fontWeight: '500' }}>Click to upload files</p>
        <p style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>PDF, DOCX, TXT or Markdown</p>
        <input id={id} type="file" multiple style={{ display: 'none' }} onChange={handleFileChange} />
      </div>

      {files.length > 0 && (
        <div style={{ display: 'grid', gap: '0.75rem' }}>
          {files.map((file, i) => (
            <div key={i} style={{ 
              display: 'flex', 
              alignItems: 'center', 
              gap: '1rem', 
              padding: '0.75rem', 
              background: 'white', 
              borderRadius: '8px',
              border: '1px solid #e2e8f0'
            }}>
              <FileText size={20} color="var(--primary-blue)" />
              <div style={{ flex: 1 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span style={{ fontSize: '0.85rem', fontWeight: '600', maxWidth: '180px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{file.name}</span>
                  <button onClick={() => removeFile(i)} style={{ background: 'none', border: 'none', color: '#94a3b8', cursor: 'pointer' }}><X size={16}/></button>
                </div>
                <select 
                  className="glass-input" 
                  style={{ marginTop: '0.4rem', padding: '0.3rem 0.5rem', fontSize: '0.8rem' }}
                  value={mapping[file.name] || ''}
                  onChange={(e) => updateMapping(file.name, e.target.value)}
                >
                  <option value="">Select Document Type...</option>
                  {Object.entries(docTypes).map(([key, info]) => (
                    <option key={key} value={key}>{info.title}</option>
                  ))}
                </select>
              </div>
            </div>
          ))}
        </div>
      )}
    </GlassCard>
  );
};

export default DocUpload;
