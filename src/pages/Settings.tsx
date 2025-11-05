import React, { useEffect, useState } from 'react';
import { Toggle, Dropdown, IDropdownOption, PrimaryButton, TextField, MessageBar, MessageBarType } from '@fluentui/react';
import { saveToStorage, SaveError } from '../api/saveToStorage';
import '../Theme.css';

const themeOptions: IDropdownOption[] = [
  { key: 'dark', text: 'Dark' },
  { key: 'light', text: 'Light' },
];

const SettingsPage: React.FC = () => {
  const [theme, setTheme] = useState<string>(localStorage.getItem('appTheme') || 'dark');
  const [animations, setAnimations] = useState<boolean>(localStorage.getItem('appAnimations') !== 'false');
  const [compact, setCompact] = useState<boolean>(localStorage.getItem('appCompact') === 'true');
  // Simple storage test state
  const [testUid, setTestUid] = useState<string>('99999999999');
  const [testLoading, setTestLoading] = useState<boolean>(false);
  const [testOk, setTestOk] = useState<string | null>(null);
  const [testErr, setTestErr] = useState<string | null>(null);

  useEffect(() => {
    applyTheme(theme, animations, compact);
  }, [theme, animations, compact]);

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

  const handleStorageTest = async () => {
    setTestOk(null); setTestErr(null); setTestLoading(true);
    try {
      const uid = (testUid || '').trim() || '99999999999';
      const email = (() => { try { return localStorage.getItem('loggedInEmail') || ''; } catch { return ''; } })();
      const resp = await saveToStorage({
        category: 'Notes',
        uid,
        title: 'Test Save',
        description: `SWA storage test at ${new Date().toISOString()}`,
        owner: email || 'tester',
      });
      setTestOk(resp || 'OK');
    } catch (e: any) {
      const msg = e instanceof SaveError ? (e.body || e.message) : (e?.body || e?.message || 'Failed to save');
      setTestErr(String(msg));
    } finally { setTestLoading(false); }
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

          {/* Storage save test */}
          <div style={{ marginTop: 28 }}>
            <div className="section-title" style={{ margin: '8px 0' }}>Storage save test</div>
            <div style={{ display: 'flex', gap: 10, alignItems: 'center', maxWidth: 540 }}>
              <TextField label="UID (for test)" value={testUid} onChange={(_, v) => setTestUid((v||'').replace(/\D/g, '').slice(0, 11))} placeholder="11-digit UID" />
              <PrimaryButton text={testLoading ? 'Testing…' : 'Run test'} disabled={testLoading} onClick={handleStorageTest} />
            </div>
            {testOk && (<div style={{ marginTop: 8 }}><MessageBar messageBarType={MessageBarType.success} isMultiline={false}>Saved successfully: {testOk.slice(0, 200)}</MessageBar></div>)}
            {testErr && (<div style={{ marginTop: 8 }}><MessageBar messageBarType={MessageBarType.error} isMultiline={false}>{testErr}</MessageBar></div>)}
          </div>
        </div>
      </div>
    </div>
  );
};

export default SettingsPage;
