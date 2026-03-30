import React from 'react';

interface GlassCardProps {
  children: React.ReactNode;
  className?: string;
  title?: string;
  subtitle?: string;
}

const GlassCard: React.FC<GlassCardProps> = ({ children, className = '', title, subtitle }) => {
  return (
    <div className={`glass-card fade-in ${className}`}>
      {(title || subtitle) && (
        <div style={{ marginBottom: '1.5rem' }}>
          {title && <h2 style={{ fontSize: '1.25rem', color: 'var(--text-main)' }}>{title}</h2>}
          {subtitle && <p style={{ fontSize: '0.875rem', color: 'var(--text-muted)' }}>{subtitle}</p>}
        </div>
      )}
      {children}
    </div>
  );
};

export default GlassCard;
