import React, { useEffect, useState } from 'react';
import { PrimaryButton, Dialog, DialogType, DialogFooter, DefaultButton, MessageBar, MessageBarType } from '@fluentui/react';
import '../Theme.css';

const WirecheckAutomation: React.FC = () => {
  const [description] = useState<string>(() => {
    try {
      return localStorage.getItem('wirecheckDescription') || 'Add a short description about Wirecheck Automation here.';
    } catch (e) {
      return 'Add a short description about Wirecheck Automation here.';
    }
  });

  // The bookmarklet contains many escaped characters and a script: URL. Disable related lint rules for this block.
  /* eslint-disable no-script-url, no-useless-escape */
  const [code] = useState<string>(() => {
    try {
      return (
        localStorage.getItem('wirecheckCode') ||
        `javascript:(async()=>{const showMsg=(t,c,d=2000)=>{const m=document.createElement('div');m.textContent=t;m.style.position='fixed';m.style.top='20px';m.style.left='50%';m.style.transform='translateX(-50%)';m.style.background=c;m.style.color='white';m.style.padding='20px 30px';m.style.borderRadius='16px';m.style.fontSize='24px';m.style.fontWeight='bold';m.style.boxShadow='0 4px 8px rgba(0,0,0,0.3)';m.style.zIndex='9999';m.id='auto-msg';document.body.appendChild(m);setTimeout(()=>{if(m)m.remove()},d);};const createLiveLog=()=>{let live=document.getElementById(\'live-log\');if(!live){live=document.createElement(\'div\');live.id=\'live-log\';live.style.position=\'fixed\';live.style.bottom=\'20px\';live.style.left=\'20px\';live.style.width=\'350px\';live.style.maxHeight=\'400px\';live.style.overflowY=\'auto\';live.style.background=\'rgba(0,120,215,0.9)\';live.style.color=\'white\';live.style.padding=\'20px\';live.style.borderRadius=\'16px\';live.style.fontSize=\'18px\';live.style.fontWeight=\'bold\';live.style.boxShadow=\'0 4px 8px rgba(0,0,0,0.3)\';live.style.zIndex=\'9999\';live.innerHTML=\'<b>📋 Live Log:</b><br>\';document.body.appendChild(live);}};const addLog=(text)=>{let live=document.getElementById(\'live-log\');if(live){live.innerHTML+=text+\'<br>\';live.scrollTop=live.scrollHeight;}};const removeLiveLog=()=>{let live=document.getElementById(\'live-log\');if(live)live.remove();};createLiveLog();showMsg(\'✅ Automation Started\',\'#4CAF50\');const e=ms=>new Promise(r=>setTimeout(r,ms)),t=new Set;function n(e){return Date.parse(e.trim())}const o=[...document.querySelectorAll(\"td.table-warning\")].filter(e=>\"pending\"===e.textContent.trim().toLowerCase());console.log(\`🟡 Found \${o.length} outer pending rows\`);for(let r=0;r<o.length;r++){const n=o[r];t.has(n)||(n.click(),t.add(n),console.log(\`🔽 Opened outer pending #\${r+1}\`),addLog(\`🔽 Opened outer pending #\${r+1}\`),await e(300))}showMsg(\'🔽 Pending Rows Expanded\',\'#03A9F4\');console.log(\"⏳ Letting all nested rows expand...\");await e(1000);const a=[...document.querySelectorAll(\"tr\")].filter(e=>{const t=e.querySelectorAll(\"td\");return t.length>=4&&\"wait\"===t[0].textContent.trim().toLowerCase()&&\"pending\"===t[2].textContent.trim().toLowerCase()});console.log(\`📥 Found \${a.length} inner wait+pending rows\`);a.forEach(e=>{const t=e.querySelectorAll(\"td\");if(t.length>=4){const o=n(t[1].textContent);if(!isNaN(o))for(let e of[0,1,2])if(t[e]&&null!==t[e].offsetParent){t[e].click();console.log(\`📥 Clicked nested row: \${t[1].textContent.trim()}\`);addLog(\`📥 Clicked nested row: \${t[1].textContent.trim()}\`);break}}});showMsg(\'📥 Nested Rows Clicked\',\'#2196F3\');await e(1000);const l=[...document.querySelectorAll(\"button.btn.btn-primary\")].filter(e=>\"respond\"===e.textContent?.trim().toLowerCase()&&null!==e.offsetParent);console.log(\`🚀 Clicking \${l.length} Respond buttons!\`);showMsg(\`🚀 Clicking \${l.length} Respond buttons...\`,\'#FFC107\');let clicked=0;l.forEach((e,t)=>{setTimeout(()=>{try{e.click();clicked++;console.log(\`✅ Respond button #\${clicked} clicked\`);addLog(\`🚀 Respond clicked: \${clicked}/\${l.length}\`);}catch(e){console.error(\`❌ Failed to click Respond #\${clicked}\`,e);addLog(\`❌ Error clicking Respond \${clicked}\`);}},50*t)});setTimeout(()=>{showMsg(\'✅ Automation Completed\',\'#4CAF50\',4000);addLog(\'✅ Automation Completed Successfully!\');},l.length*50+2000);})();`
      );
    } catch (e) {
      return `// Add your example commands or code snippets here`;
    }
  });
  /* eslint-enable no-script-url, no-useless-escape */

  const [installOpen, setInstallOpen] = useState<boolean>(false);

  useEffect(() => {
    // ensure theme classes applied if existing settings present
  }, []);

  const [showCopied, setShowCopied] = useState<boolean>(false);
  const copyToClipboard = async (text: string) => {
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(text);
      } else {
        const ta = document.createElement('textarea');
        ta.value = text;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        ta.remove();
      }
      setShowCopied(true);
      setTimeout(() => setShowCopied(false), 3000);
    } catch (e) {
      setShowCopied(false);
    }
  };

  return (
    <div className="main-content fade-in">
  <div className="vso-form-container glow" style={{ width: '95%', maxWidth: 1700 }}>
        <div className="banner-title">
          <span className="title-text">Wirecheck Automation</span>
          <span className="title-sub">Tools and automation for Wirecheck</span>
        </div>

        {/* Professional green status banner (below title, above image) */}
        <div style={{ marginTop: 12, display: 'flex', justifyContent: 'center' }}>
          <div
            role="status"
            aria-live="polite"
            style={{
              width: '100%',
              maxWidth: 1200,
              background: 'linear-gradient(180deg,#e6fff1,#d7fbe8)',
              border: '1px solid #9fe9b8',
              color: '#033a16',
              padding: '14px 18px',
              borderRadius: 8,
              display: 'flex',
              alignItems: 'center',
              gap: 14,
              boxShadow: '0 4px 12px rgba(2,72,35,0.08)'
            }}
          >
            {/* Microsoft-blue circular info icon */}
            <div style={{ width: 36, height: 36, borderRadius: 18, background: '#0078D4', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#ffffff', fontSize: 16, fontWeight: 600 }}>
              i
            </div>
            <div style={{ lineHeight: 1.2 }}>
              <div style={{ fontWeight: 700 }}>Method still working — native integration coming soon</div>
              <div style={{ fontSize: 13, marginTop: 6 }}>We're working closely with the ClockWarp team to deliver a native integration which should be in effect before the end of the year.</div>
            </div>
          </div>
        </div>

        <div style={{ padding: 20, textAlign: 'center' }}>
          {/* much larger image for easier viewing - wrapped so it can scroll horizontally */}
          <div style={{ overflowX: 'hidden', paddingBottom: 8 }}>
            {/* image removed - placeholder area kept for layout */}
            <div style={{ width: '100%', height: 240, margin: '12px auto', background: '#f3f3f3', borderRadius: 6, border: '1px dashed #ddd', color: '#666', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <span>Wirecheck image removed (asset not tracked)</span>
            </div>
          </div>

          <div style={{ marginTop: 18, textAlign: 'left', lineHeight: 1.5 }}>
            {/* SharePoint-style rich description */}
            <div className="wirecheck-description" style={{ background: 'transparent', padding: 12 }}>
              <h2 style={{ marginTop: 0 }}>Wirecheck Automation Video Tutorial: Streamlining Link Validation in Clock Warp</h2>

              <p style={{ marginTop: 6 }}>
                In this video, I'll walk you through my custom-built <strong>Wirecheck Automation</strong>—a JavaScript-based browser automation tool
                designed to streamline the validation and processing of Wirecheck results within Clock Warp workflows.
              </p>

              <p>
                {description}
              </p>

              <h3 style={{ marginBottom: 6 }}>🔍 How It Works:</h3>
              <ul style={{ marginTop: 0 }}>
                <li><strong>Workflow Detection:</strong> Automatically identifies if the Clock Warp workflow is open.</li>
                <li><strong>Pending Row Identification:</strong> Scans the page for all rows marked as 'Pending', including nested ones.</li>
                <li><strong>Automatic Expansion:</strong> Expands all collapsed rows to expose hidden 'Respond' buttons.</li>
                <li><strong>Action Automation:</strong> Clicks all 'Respond' buttons in rapid succession—completing the process before the server-side refresh.</li>
                <li><strong>Status Monitoring:</strong> Confirms successful responses and moves the workflow forward efficiently.</li>
              </ul>

              <h4 style={{ marginTop: 10 }}>Strategic Benefits</h4>
              <ul>
                <li><strong>95% Reduction in Manual Effort:</strong> Hundreds of clicks condensed into a single action.</li>
                <li><strong>Accelerated Workflow Clearance:</strong> Links validated in seconds, not hours.</li>
                <li><strong>Boosted System Throughput:</strong> Minimizes bottlenecks and repetitive steps.</li>
                <li><strong>Significant Time Savings:</strong> Early testing shows major improvements in completion times.</li>
              </ul>
            </div>
            {/* Single button to open instructions dialog */}
            <div style={{ marginTop: 12, display: 'flex', justifyContent: 'flex-start', gap: 8 }}>
              <PrimaryButton text="Open instructions" onClick={() => setInstallOpen(true)} />
            </div>
          </div>
        </div>
      </div>
      {/* Install dialog with steps and full code snippet */}
      <Dialog
        hidden={!installOpen}
        onDismiss={() => setInstallOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'How to install — QC Auto Retry',
          closeButtonAriaLabel: 'Close',
        }}
        modalProps={{ isBlocking: false }}
      >
        <div style={{ paddingTop: 6 }}>
          <ol>
            <li>📋 Copy the Code Below.</li>
            <li>⭐️ Press <strong>CTRL + D</strong> → Click "More..." → Edit the URL with the code you copied.</li>
            <li>🏷️ Name it: <strong>✅ QC Auto Retry</strong>.</li>
            <li>📌 Show Favorites Bar: Navigate to your Favorites page, click the three dots ⋮ → Select "Show Favorites Bar" → Choose "Always".</li>
          </ol>

          <div style={{ marginTop: 12 }}>
            <div style={{ background: '#0b0b0b', color: '#e6eef6', padding: 12, borderRadius: 6, fontFamily: 'Consolas, Menlo, monospace', fontSize: 12, overflowX: 'auto', whiteSpace: 'pre' }}>
              {code}
            </div>
          </div>
          {showCopied && (
            <div style={{ marginTop: 10 }}>
              <MessageBar messageBarType={MessageBarType.success}>Bookmarklet copied to clipboard</MessageBar>
            </div>
          )}
        </div>

        <DialogFooter>
          <PrimaryButton onClick={() => copyToClipboard(code)} text="Copy Code" />
          <DefaultButton onClick={() => setInstallOpen(false)} text="Close" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default WirecheckAutomation;
