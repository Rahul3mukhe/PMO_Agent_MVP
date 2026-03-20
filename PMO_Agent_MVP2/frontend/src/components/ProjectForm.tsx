import React, { useState } from 'react';
import type { Project } from '../types';
import GlassCard from './GlassCard';
import { ChevronRight, ChevronLeft, Check, Plus, Trash2 } from 'lucide-react';

interface ProjectFormProps {
  onSubmit: (data: Project) => void;
  initialData?: Partial<Project>;
}

const ProjectForm: React.FC<ProjectFormProps> = ({ onSubmit, initialData }) => {
  const [step, setStep] = useState(1);
  const [formData, setFormData] = useState<Project>({
    project_id: initialData?.project_id || '',
    project_name: initialData?.project_name || '',
    project_type: initialData?.project_type || 'Development',
    sponsor: initialData?.sponsor || '',
    estimated_budget: initialData?.estimated_budget || 0,
    actual_budget_consumed: initialData?.actual_budget_consumed || 0,
    total_time_taken_days: initialData?.total_time_taken_days || 0,
    timeline_summary: initialData?.timeline_summary || '',
    scope_summary: initialData?.scope_summary || '',
    key_deliverables: initialData?.key_deliverables || [],
    known_risks: initialData?.known_risks || [],
  });

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => {
    const { name, value } = e.target;
    setFormData(prev => ({
      ...prev,
      [name]: name.includes('budget') || name === 'total_time_taken_days' ? Number(value) : value
    }));
  };

  const handleListChange = (key: 'key_deliverables' | 'known_risks', index: number, value: string) => {
    const newList = [...formData[key]];
    newList[index] = value;
    setFormData(prev => ({ ...prev, [key]: newList }));
  };

  const addListItem = (key: 'key_deliverables' | 'known_risks') => {
    setFormData(prev => ({ ...prev, [key]: [...prev[key], ''] }));
  };

  const removeListItem = (key: 'key_deliverables' | 'known_risks', index: number) => {
    setFormData(prev => ({ ...prev, [key]: prev[key].filter((_, i) => i !== index) }));
  };

  const nextStep = () => setStep(prev => Math.min(prev + 1, 3));
  const prevStep = () => setStep(prev => Math.max(prev - 1, 1));

  return (
    <div className="project-form">
      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '2rem' }}>
        {[1, 2, 3].map(s => (
          <div key={s} style={{ 
            display: 'flex', 
            alignItems: 'center', 
            gap: '0.5rem',
            color: step >= s ? 'var(--primary-blue)' : 'var(--text-muted)',
            fontWeight: step === s ? '700' : '400'
          }}>
            <div style={{
              width: '24px', height: '24px', borderRadius: '50%', 
              background: step >= s ? 'var(--primary-blue)' : 'var(--accent-blue)',
              color: step >= s ? 'white' : 'var(--primary-blue)',
              display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.75rem'
            }}>
              {step > s ? <Check size={14} /> : s}
            </div>
            <span style={{ fontSize: '0.85rem' }}>{s === 1 ? 'Core' : s === 2 ? 'Finance' : 'Scope'}</span>
          </div>
        ))}
      </div>

      <GlassCard title={step === 1 ? 'Project Foundation' : step === 2 ? 'Timeline & Budget' : 'Scope & Risks'}>
        {step === 1 && (
          <div style={{ display: 'grid', gap: '1.25rem' }}>
            <div>
              <label style={{ display: 'block', fontSize: '0.75rem', color: 'var(--text-muted)', marginBottom: '0.4rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Project ID</label>
              <input className="glass-input" name="project_id" value={formData.project_id} onChange={handleChange} placeholder="e.g. PRJ-771" />
            </div>
            <div>
              <label style={{ display: 'block', fontSize: '0.75rem', color: 'var(--text-muted)', marginBottom: '0.4rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Project Name</label>
              <input className="glass-input" name="project_name" value={formData.project_name} onChange={handleChange} placeholder="Enter official name..." />
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem' }}>
              <div>
                <label style={{ display: 'block', fontSize: '0.75rem', color: 'var(--text-muted)', marginBottom: '0.4rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Type</label>
                <select className="glass-input" name="project_type" value={formData.project_type} onChange={handleChange}>
                  <option>Development</option>
                  <option>Research</option>
                  <option>Operations</option>
                  <option>Regulated</option>
                </select>
              </div>
              <div>
                <label style={{ display: 'block', fontSize: '0.75rem', color: 'var(--text-muted)', marginBottom: '0.4rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Sponsor</label>
                <input className="glass-input" name="sponsor" value={formData.sponsor} onChange={handleChange} placeholder="Executive name..." />
              </div>
            </div>
          </div>
        )}

        {step === 2 && (
          <div style={{ display: 'grid', gap: '1.25rem' }}>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem' }}>
              <div>
                <label style={{ display: 'block', fontSize: '0.75rem', color: 'var(--text-muted)', marginBottom: '0.4rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Est. Budget ($)</label>
                <input type="number" className="glass-input" name="estimated_budget" value={formData.estimated_budget} onChange={handleChange} />
              </div>
              <div>
                <label style={{ display: 'block', fontSize: '0.75rem', color: 'var(--text-muted)', marginBottom: '0.4rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Actual Consumed ($)</label>
                <input type="number" className="glass-input" name="actual_budget_consumed" value={formData.actual_budget_consumed} onChange={handleChange} />
              </div>
            </div>
            <div>
              <label style={{ display: 'block', fontSize: '0.75rem', color: 'var(--text-muted)', marginBottom: '0.4rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Total Days</label>
              <input type="number" className="glass-input" name="total_time_taken_days" value={formData.total_time_taken_days} onChange={handleChange} />
            </div>
            <div>
              <label style={{ display: 'block', fontSize: '0.75rem', color: 'var(--text-muted)', marginBottom: '0.4rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Timeline Summary</label>
              <textarea className="glass-input" name="timeline_summary" rows={3} value={formData.timeline_summary} onChange={handleChange} placeholder="Briefly describe the schedule..." />
            </div>
          </div>
        )}

        {step === 3 && (
          <div style={{ display: 'grid', gap: '1.25rem' }}>
            <div>
              <label style={{ display: 'block', fontSize: '0.75rem', color: 'var(--text-muted)', marginBottom: '0.4rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Scope Summary</label>
              <textarea className="glass-input" name="scope_summary" rows={3} value={formData.scope_summary} onChange={handleChange} placeholder="What does this project cover?" />
            </div>
            
            <div>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.5rem' }}>
                <label style={{ fontSize: '0.75rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Key Deliverables</label>
                <button className="btn-secondary" onClick={() => addListItem('key_deliverables')} style={{ padding: '0.2rem 0.5rem', fontSize: '0.7rem' }}><Plus size={12} /></button>
              </div>
              <div style={{ display: 'grid', gap: '0.5rem' }}>
                {formData.key_deliverables.map((item, i) => (
                  <div key={i} style={{ display: 'flex', gap: '0.5rem' }}>
                    <input className="glass-input" value={item} onChange={(e) => handleListChange('key_deliverables', i, e.target.value)} />
                    <button onClick={() => removeListItem('key_deliverables', i)} style={{ background: 'none', border: 'none', color: '#ef4444', cursor: 'pointer' }}><Trash2 size={16}/></button>
                  </div>
                ))}
              </div>
            </div>

            <div>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.5rem' }}>
                <label style={{ fontSize: '0.75rem', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Known Risks</label>
                <button className="btn-secondary" onClick={() => addListItem('known_risks')} style={{ padding: '0.2rem 0.5rem', fontSize: '0.7rem' }}><Plus size={12} /></button>
              </div>
              <div style={{ display: 'grid', gap: '0.5rem' }}>
                {formData.known_risks.map((item, i) => (
                  <div key={i} style={{ display: 'flex', gap: '0.5rem' }}>
                    <input className="glass-input" value={item} onChange={(e) => handleListChange('known_risks', i, e.target.value)} />
                    <button onClick={() => removeListItem('known_risks', i)} style={{ background: 'none', border: 'none', color: '#ef4444', cursor: 'pointer' }}><Trash2 size={16}/></button>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '2rem' }}>
          {step > 1 ? (
            <button className="btn-secondary" onClick={prevStep} style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
              <ChevronLeft size={18} /> Previous
            </button>
          ) : <div />}
          
          {step < 3 ? (
            <button className="btn-primary" onClick={nextStep} style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
              Next <ChevronRight size={18} />
            </button>
          ) : (
            <button className="btn-primary" onClick={() => onSubmit(formData)} style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
              Launch AI Analysis <Check size={18} />
            </button>
          )}
        </div>
      </GlassCard>
    </div>
  );
};

export default ProjectForm;
