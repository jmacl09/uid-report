import React, { useEffect, useState } from 'react';
import { Toggle, Dropdown, IDropdownOption, PrimaryButton } from '@fluentui/react';
import '../Theme.css';

const themeOptions: IDropdownOption[] = [
  { key: 'dark', text: 'Dark' },
  { key: 'light', text: 'Light' },
];

const SettingsPage: React.FC = () => {
  const [theme, setTheme] = useState<string>(localStorage.getItem('appTheme') || 'dark');
  const [animations, setAnimations] = useState<boolean>(localStorage.getItem('appAnimations') !== 'false');
  const [compact, setCompact] = useState<boolean>(localStorage.getItem('appCompact') === 'true');

  useEffect(() => {
    applyTheme(theme, animations, compact);
  }, []);

  const applyTheme = (t: string, anim: boolean, comp: boolean) => {
    try {
      const root = document.documentElement;
      if (t === 'light') root.classList.add('light-theme'); else root.classList.remove('light-theme');
      if (!anim) root.classList.add('no-animations'); else root.classList.remove('no-animations');
      if (comp) root.classList.add('compact-mode'); else root.classList.remove('compact-mode');
    } catch (e) {}
  };

  const handleSave = () => {
    localStorage.setItem('appTheme', theme);
    localStorage.setItem('appAnimations', animations ? 'true' : 'false');
    localStorage.setItem('appCompact', compact ? 'true' : 'false');
    applyTheme(theme, animations, compact);
    alert('Settings saved.');
  };

  return (
    <div className="main-content fade-in">
      <div className="vso-form-container glow" style={{ width: '80%', maxWidth: 900 }}>
        <div className="banner-title">
          <span className="title-text">Settings</span>
          <span className="title-sub">Site preferences</span>
        </div>

        <div style={{ padding: 20 }}>
          <div style={{ marginBottom: 14 }}>
            <div style={{ marginBottom: 6, color: '#ccc', fontWeight: 600 }}>Theme</div>
            <Dropdown
              options={themeOptions}
              selectedKey={theme}
              onChange={(_, opt) => opt && setTheme(opt.key as string)}
            />
          </div>

          <div style={{ marginBottom: 14 }}>
            <Toggle
              label="Enable Animations"
              checked={animations}
              onChange={(_, v) => setAnimations(!!v)}
            />
          </div>

          <div style={{ marginBottom: 14 }}>
            <Toggle
              label="Compact Mode (reduce paddings)"
              checked={compact}
              onChange={(_, v) => setCompact(!!v)}
            />
          </div>

          <div style={{ display: 'flex', gap: 8, marginTop: 12 }}>
            <PrimaryButton text="Save" onClick={handleSave} />
          </div>
        </div>
      </div>
    </div>
  );
};

export default SettingsPage;
