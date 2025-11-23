import React from 'react';
import '../Theme.css';

const DCATAssistant: React.FC = () => {
  return (
    <div className="main-content fade-in">
      <div className="vso-form-container glow" style={{ width: '80%', maxWidth: 1000 }}>
        <div className="banner-title">
          <span className="title-text">DCAT Assistant</span>
          <span className="title-sub">Coming Soon</span>
        </div>
        <div style={{ padding: 24, textAlign: 'center' }}>
          <h1 className="dcat-coming-title" style={{ fontSize: 36 }}>COMING SOON...</h1>
          <p style={{ color: '#a6b7c6', fontSize: 16, maxWidth: 820, margin: '12px auto' }}>
            The DCAT Assistant will give an overview of DCATs per project to ensure accuracy and efficiency.
          </p>
        </div>
      </div>
    </div>
  );
};

export default DCATAssistant;
