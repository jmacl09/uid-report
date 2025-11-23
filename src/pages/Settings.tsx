import React, { useEffect, useState } from 'react';
import { Dropdown, IDropdownOption } from '@fluentui/react';
import '../Theme.css';

const themeOptions: IDropdownOption[] = [
  { key: 'dark', text: 'Dark' },
  { key: 'light', text: 'Light' },
];

const SettingsPage: React.FC = () => {
  const [theme, setTheme] = useState<string>(localStorage.getItem('appTheme') || 'dark');

  useEffect(() => {
    try {
      const root = document.documentElement;
      if (theme === 'light') root.classList.add('light-theme'); else root.classList.remove('light-theme');
    } catch (e) {}
  }, [theme]);

  const handleThemeChange = (t: string) => {
    try {
      setTheme(t);
      localStorage.setItem('appTheme', t);
    } catch (e) {}
  };

  const dropdownStyles = {
    title: {
      backgroundColor: theme === 'light' ? '#ffffff' : '#141414',
      color: theme === 'light' ? 'var(--vso-dropdown-text, #0f172a)' : '#ffffff',
      border: theme === 'light' ? '1px solid rgba(195, 218, 244, 0.6)' : '1px solid #333',
    },
    caretDownWrapper: {
      color: theme === 'light' ? 'var(--vso-dropdown-text, #0f172a)' : '#ffffff',
    },
  };

  return (
    <div className="main-content fade-in">
      <div className="vso-form-container glow" style={{ width: '60%', maxWidth: 700 }}>
        <div className="banner-title">
          <span className="title-text">Settings</span>
          <span className="title-sub">Site preferences</span>
        </div>
        <div style={{ padding: 8 }}>
          <div style={{ marginBottom: 8 }}>
            <div style={{ marginBottom: 4, color: '#ccc', fontWeight: 600 }}>Theme</div>
            <Dropdown
              className={theme === 'light' ? 'dropdown-light' : 'dropdown-dark'}
              options={themeOptions}
              selectedKey={theme}
              onChange={(_, opt) => opt && handleThemeChange(opt.key as string)}
              styles={dropdownStyles as any}
            />
          </div>
        </div>
      </div>
    </div>
  );
};

export default SettingsPage;
