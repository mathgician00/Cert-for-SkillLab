'use client';

import { useEffect, useState } from 'react';
import { User } from 'firebase/auth';
import { initAuth, googleSignIn, getAccessToken, logout } from '@/lib/firebase';
import { LogOut, User as UserIcon, RefreshCw, FileText, CheckCircle2, XCircle, Loader2 } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

type RowData = {
  rowIndex: number;
  linkColLetter: string;
  name: string;
  rawDate: string;
  course: string;
  displayDate?: string;
  yymmdd?: string;
};

export default function App() {
  const [needsAuth, setNeedsAuth] = useState(true);
  const [user, setUser] = useState<User | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [loading, setLoading] = useState(true);

  // App state
  const [checkingRecap, setCheckingRecap] = useState(false);
  const [sheetUrl, setSheetUrl] = useState('');
  const [pendingRows, setPendingRows] = useState<RowData[]>([]);
  const [totalPending, setTotalPending] = useState(-1); // -1 = hasn't checked
  const [errorMsg, setErrorMsg] = useState('');
  
  // Override URL
  const [overrideUrl, setOverrideUrl] = useState('');
  const [showOverride, setShowOverride] = useState(false);

  // Generation state
  const [generating, setGenerating] = useState(false);
  const [progress, setProgress] = useState(0);
  const [logs, setLogs] = useState<{msg: string; type: 'ok'|'err'|'info'}[]>([]);
  const [rowStatuses, setRowStatuses] = useState<Record<number, 'pending'|'processing'|'done'|'error'>>({});

  useEffect(() => {
    const unsubscribe = initAuth(
      (u) => {
        setUser(u);
        setNeedsAuth(false);
        setLoading(false);
      },
      () => {
        setUser(null);
        setNeedsAuth(true);
        setLoading(false);
      }
    );
    return () => unsubscribe();
  }, []);

  const handleLogin = async () => {
    setIsLoggingIn(true);
    setErrorMsg('');
    try {
      const result = await googleSignIn();
      if (result) {
        setUser(result.user);
        setNeedsAuth(false);
      }
    } catch (err: any) {
      console.error('Login failed:', err);
      setErrorMsg('Failed to sign in. Make sure popups are allowed and Firebase config is correct.');
    } finally {
      setIsLoggingIn(false);
    }
  };

  const handleLogout = async () => {
    await logout();
    // Reset state
    setSheetUrl('');
    setPendingRows([]);
    setTotalPending(-1);
    setLogs([]);
    setRowStatuses({});
  };

  const checkRecapSheet = async (override?: string) => {
    setProgress(0);
    setErrorMsg('');
    setCheckingRecap(true);
    setLogs([]);
    setRowStatuses({});
    
    try {
      const token = await getAccessToken();
      const res = await fetch('/api/certificates/check-recap', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`
        },
        body: JSON.stringify({ sheetUrl: override || '' })
      });
      const data = await res.json();
      
      if (!res.ok) throw new Error(data.error || 'Failed to check recap sheet.');

      setSheetUrl(data.url);
      setPendingRows(data.rows);
      setTotalPending(data.total);
      
      // Initialize row statuses
      const initialStatuses: any = {};
      data.rows.forEach((r: RowData) => {
        initialStatuses[r.rowIndex] = 'pending';
      });
      setRowStatuses(initialStatuses);
      
    } catch (err: any) {
      setErrorMsg(err.message);
    } finally {
      setCheckingRecap(false);
      setShowOverride(false);
    }
  };

  const handleGenerate = async () => {
    if (pendingRows.length === 0) return;
    
    // Explicit User Confirmation per Guidelines
    const confirmMessage = `Please ensure the capitalization for ${pendingRows.length} student(s) for the certificates are correct. This will appear as inputted in the sheet.`;
    if (!window.confirm(confirmMessage)) return;

    setGenerating(true);
    setProgress(0);
    setLogs([{ msg: 'Starting generation process...', type: 'info' }]);
    setErrorMsg('');

    try {
      const token = await getAccessToken();
      
      // Process in small batches (e.g., 3 at a time) to avoid timeout/rate limits
      const BATCH_SIZE = 3;
      let completedCount = 0;
      
      for (let i = 0; i < pendingRows.length; i += BATCH_SIZE) {
        const batch = pendingRows.slice(i, i + BATCH_SIZE);
        
        batch.forEach(r => {
          setRowStatuses(prev => ({ ...prev, [r.rowIndex]: 'processing' }));
        });

        const res = await fetch('/api/certificates/generate', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${token}`
          },
          body: JSON.stringify({
            sheetUrl,
            rows: batch,
            userEmail: user?.email
          })
        });
        
        const data = await res.json();
        
        if (!res.ok) {
           batch.forEach(r => {
             setRowStatuses(prev => ({ ...prev, [r.rowIndex]: 'error' }));
             setLogs(prev => [...prev, { msg: `✗ Row ${r.rowIndex}: Failed - ${data.error}`, type: 'err' }]);
           });
           completedCount += batch.length;
           setProgress(Math.round((completedCount / pendingRows.length) * 100));
           continue;
        }

        data.results.forEach((resItem: any) => {
          if (resItem.error) {
            setRowStatuses(prev => ({ ...prev, [resItem.rowIndex]: 'error' }));
            setLogs(prev => [...prev, { msg: `✗ Row ${resItem.rowIndex}: ${resItem.error}`, type: 'err' }]);
          } else {
            setRowStatuses(prev => ({ ...prev, [resItem.rowIndex]: 'done' }));
            setLogs(prev => [...prev, { msg: `✓ Row ${resItem.rowIndex}: Generated successfully`, type: 'ok' }]);
          }
        });
        
        completedCount += batch.length;
        setProgress(Math.round((completedCount / pendingRows.length) * 100));
      }
      
      setLogs(prev => [...prev, { msg: 'Generation process completed.', type: 'info' }]);

    } catch (err: any) {
      setErrorMsg(err.message || 'An error occurred during generation.');
      setLogs(prev => [...prev, { msg: 'Process failed completely.', type: 'err' }]);
    } finally {
      setGenerating(false);
    }
  };

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-50">
        <Loader2 className="w-8 h-8 text-blue-600 animate-spin" />
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-900 flex flex-col overflow-x-hidden">
      <header className="flex items-center justify-between px-8 py-5 bg-white border-b border-slate-200">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-indigo-600 rounded flex items-center justify-center">
            <svg className="w-5 h-5 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M8 9l3 3-3 3m5 0h3M5 20h14a2 2 0 002-2V6a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" /></svg>
          </div>
          <h1 className="text-xl font-semibold tracking-tight text-slate-800 hidden sm:block">Certificate Generator <span className="text-xs font-normal text-slate-400 ml-2">SkillLab</span></h1>
        </div>
        {user && (
          <div className="flex items-center gap-6">
            <div className="flex items-center gap-2 px-3 py-1 bg-emerald-50 text-emerald-700 border border-emerald-100 rounded-full text-xs font-medium">
              <div className="w-1.5 h-1.5 bg-emerald-500 rounded-full"></div>
              Authenticated
            </div>
            <div className="flex items-center gap-3 border-l border-slate-200 pl-6">
              <div className="text-right hidden sm:block">
                <p className="text-xs font-semibold">Active Session</p>
                <p className="text-[10px] text-slate-400 max-w-[150px] truncate">{user.email}</p>
              </div>
              <div className="w-9 h-9 bg-slate-200 rounded-full flex items-center justify-center">
                <span className="text-slate-600 text-xs font-bold">{user.email?.[0].toUpperCase()}</span>
              </div>
              <button onClick={handleLogout} className="ml-2 text-slate-400 hover:text-slate-700 transition-colors">
                <LogOut className="w-4 h-4" />
              </button>
            </div>
          </div>
        )}
      </header>

      <main className="flex-1 p-8 overflow-x-hidden flex flex-col gap-8 max-w-5xl mx-auto w-full">
        
        {/* Error Alert */}
        <AnimatePresence>
          {errorMsg && (
            <motion.div initial={{ opacity: 0, y: -10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, height: 0 }} 
              className="bg-red-50 text-red-700 border border-red-200 p-4 rounded-lg text-sm flex gap-3 items-start">
              <XCircle className="w-5 h-5 shrink-0 mt-0.5 text-red-600" />
              <div className="leading-relaxed">{errorMsg}</div>
            </motion.div>
          )}
        </AnimatePresence>

        {needsAuth ? (
          // Login State
          <motion.div initial={{ opacity: 0, scale: 0.95 }} animate={{ opacity: 1, scale: 1 }} className="flex flex-col items-center justify-center text-center mt-12 gap-6 bg-white p-10 rounded-xl border border-slate-200 shadow-sm max-w-xl mx-auto w-full">
            <div className="w-16 h-16 bg-slate-100 text-slate-500 rounded-xl flex items-center justify-center border border-slate-200">
              <UserIcon className="w-8 h-8" />
            </div>
            <div>
              <h2 className="text-2xl font-bold text-slate-800 mb-2">Use Branch Google Account</h2>
              <p className="text-sm text-slate-500 max-w-sm mx-auto leading-relaxed">
                Sign in to generate certificates, recaped and saved automatically to your Google Drive.
              </p>
            </div>
            <button
              onClick={handleLogin}
              disabled={isLoggingIn}
              className="gsi-material-button relative overflow-hidden bg-white text-slate-700 font-medium text-sm border border-slate-300 rounded-md hover:bg-slate-50 transition-colors shadow-sm disabled:opacity-70 flex items-center pr-4"
            >
              <div className="p-3 bg-white">
                <svg version="1.1" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48" className="w-5 h-5">
                  <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"></path>
                  <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"></path>
                  <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"></path>
                  <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"></path>
                </svg>
              </div>
              <span className="pl-3 font-roboto">{isLoggingIn ? 'Connecting...' : 'Sign in with Google'}</span>
            </button>
          </motion.div>
        ) : (
          // Authenticated State
          <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="flex flex-col gap-6">
            
            {/* Step 1: Check Recap Sheet */}
            {totalPending === -1 && (
              <div className="bg-white p-10 rounded-xl border border-slate-200 shadow-sm flex flex-col items-center text-center gap-6 max-w-xl mx-auto w-full">
                 <div className="w-16 h-16 bg-indigo-50 text-indigo-600 rounded-xl flex items-center justify-center border border-indigo-100">
                  <FileText className="w-8 h-8" />
                </div>
                <div>
                  <h2 className="text-2xl font-bold text-slate-800 mb-2">Check Recap Sheet</h2>
                  <p className="text-sm text-slate-500 max-w-sm mx-auto leading-relaxed">
                    We will scan your Drive for a Recap Sheet and find any rows that need certificates generated.
                  </p>
                </div>
                
                {showOverride ? (
                   <div className="w-full max-w-md flex flex-col gap-3">
                     <label className="text-xs font-semibold text-slate-700 text-left uppercase tracking-wider">Custom Sheet URL</label>
                     <div className="flex gap-2">
                       <input 
                         type="text" 
                         className="flex-1 bg-slate-50 border border-slate-300 text-slate-900 text-sm rounded-md focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 block w-full px-3 py-2" 
                         placeholder="https://docs.google.com/spreadsheets/d/..."
                         value={overrideUrl}
                         onChange={e => setOverrideUrl(e.target.value)}
                       />
                       <button 
                         onClick={() => checkRecapSheet(overrideUrl)}
                         disabled={checkingRecap || !overrideUrl}
                         className="bg-slate-900 text-white hover:bg-slate-800 font-medium rounded-md text-sm px-5 py-2 disabled:opacity-50 transition-colors"
                       >
                         {checkingRecap ? <Loader2 className="w-4 h-4 animate-spin" /> : 'Load'}
                       </button>
                     </div>
                     <button onClick={() => setShowOverride(false)} className="text-xs text-slate-500 hover:text-slate-800 self-start transition-colors">Cancel</button>
                   </div>
                ) : (
                  <div className="flex flex-col items-center gap-4">
                    <button 
                      onClick={() => checkRecapSheet()}
                      disabled={checkingRecap}
                      className="bg-slate-900 text-white hover:bg-slate-800 font-medium rounded-md text-sm px-8 py-3 disabled:opacity-50 transition-colors shadow-sm flex items-center gap-2"
                    >
                      {checkingRecap && <Loader2 className="w-4 h-4 animate-spin" />}
                      {checkingRecap ? 'Checking...' : 'Check Recap Sheet'}
                    </button>
                    <button onClick={() => setShowOverride(true)} className="text-xs text-indigo-600 hover:text-indigo-700 font-medium transition-colors">
                      Use a specific sheet URL instead
                    </button>
                  </div>
                )}
              </div>
            )}

            {/* Step 2: Show Results and Generate */}
            {totalPending !== -1 && (
              <>
                <div className="flex items-center justify-between mb-2">
                   <div>
                     <h2 className="text-lg font-semibold text-slate-800">Recap Sheet Loaded</h2>
                     {sheetUrl && (
                       <a href={sheetUrl} target="_blank" rel="noreferrer" className="text-xs text-indigo-600 hover:text-indigo-700 transition-colors inline-flex items-center mt-1">
                         Open in Google Sheets ↗
                       </a>
                     )}
                   </div>
                   <button 
                     onClick={() => { setTotalPending(-1); setSheetUrl(''); setPendingRows([]); }}
                     className="text-xs bg-slate-100 border border-slate-200 text-slate-700 px-3 py-1.5 rounded-md hover:bg-slate-200 transition-colors font-medium"
                   >
                     Reset
                   </button>
                </div>

                {totalPending === 0 ? (
                  <div className="bg-emerald-50 border border-emerald-100 p-6 rounded-xl flex items-center gap-4 shadow-sm">
                    <CheckCircle2 className="w-8 h-8 text-emerald-600 shrink-0" />
                    <div>
                      <h3 className="font-semibold text-emerald-900 text-lg">All caught up! Ccertificates for your participants are ready.</h3>
                      <p className="text-sm text-emerald-700 mt-1">Refresh if new participants just added.</p>
                    </div>
                  </div>
                ) : (
                  <div className="bg-white border border-slate-200 rounded-xl shadow-sm flex flex-col overflow-hidden">
                    <div className="px-6 py-4 border-b border-slate-100 flex justify-between items-center bg-white">
                      <div className="flex items-center gap-3">
                        <h3 className="font-semibold text-slate-700 text-sm">Pending Certificates</h3>
                        <span className="bg-indigo-50 text-indigo-600 text-[10px] font-bold px-2 py-0.5 rounded-full">{pendingRows.length} total</span>
                      </div>
                      <button 
                        onClick={() => checkRecapSheet(sheetUrl)} 
                        disabled={checkingRecap || generating}
                        className="text-slate-400 hover:text-slate-600 transition-colors disabled:opacity-50"
                        title="Refresh"
                      >
                        <RefreshCw className={`w-4 h-4 ${checkingRecap ? 'animate-spin' : ''}`} />
                      </button>
                    </div>
                    
                    <div className="overflow-x-auto flex-1">
                      <table className="w-full text-left">
                        <thead className="text-[11px] text-slate-400 uppercase tracking-widest border-b border-slate-50 bg-slate-50/50 sticky top-0 z-10">
                          <tr>
                            <th className="px-6 py-3 font-semibold">Row</th>
                            <th className="px-6 py-3 font-semibold">Name</th>
                            <th className="px-6 py-3 font-semibold">Date</th>
                            <th className="px-6 py-3 font-semibold">Course</th>
                            <th className="px-6 py-3 font-semibold">Status</th>
                          </tr>
                        </thead>
                        <tbody className="text-sm text-slate-600 divide-y divide-slate-50">
                          {pendingRows.map((row) => (
                            <tr key={row.rowIndex} className="hover:bg-slate-50/50 transition-colors">
                              <td className="px-6 py-4 font-medium text-slate-400">{row.rowIndex}</td>
                              <td className="px-6 py-4 text-slate-800 font-medium">{row.name || <span className="text-red-500 text-[10px] uppercase font-bold bg-red-50 px-2 py-0.5 rounded-full">Missing</span>}</td>
                              <td className="px-6 py-4">{row.rawDate || <span className="text-red-500 text-[10px] uppercase font-bold bg-red-50 px-2 py-0.5 rounded-full">Missing</span>}</td>
                              <td className="px-6 py-4">{row.course || <span className="text-red-500 text-[10px] uppercase font-bold bg-red-50 px-2 py-0.5 rounded-full">Missing</span>}</td>
                              <td className="px-6 py-4">
                                {rowStatuses[row.rowIndex] === 'pending' && <span className="px-2 py-0.5 bg-slate-100 text-slate-500 rounded-full text-[10px] font-bold inline-flex items-center gap-1.5"><span className="w-1 h-1 bg-slate-400 rounded-full"></span>STANDBY</span>}
                                {rowStatuses[row.rowIndex] === 'processing' && <span className="px-2 py-0.5 bg-indigo-50 text-indigo-600 rounded-full text-[10px] font-bold inline-flex items-center gap-1.5"><span className="w-1 h-1 bg-indigo-500 rounded-full animate-pulse"></span>PROCESSING</span>}
                                {rowStatuses[row.rowIndex] === 'done' && <span className="px-2 py-0.5 bg-emerald-50 text-emerald-600 rounded-full text-[10px] font-bold inline-flex items-center gap-1.5"><span className="w-1 h-1 bg-emerald-500 rounded-full"></span>DONE</span>}
                                {rowStatuses[row.rowIndex] === 'error' && <span className="px-2 py-0.5 bg-red-50 text-red-600 rounded-full text-[10px] font-bold inline-flex items-center gap-1.5"><span className="w-1 h-1 bg-red-500 rounded-full"></span>ERROR</span>}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    {(generating || progress > 0) && (
                      <div className="px-6 py-5 border-t border-slate-100 bg-white flex flex-col gap-3">
                         <div className="flex justify-between text-[11px] font-semibold text-slate-500 uppercase tracking-wider">
                           <span>{progress === 100 ? 'Completed' : 'Generating Certificates...'}</span>
                           <span>{progress}%</span>
                         </div>
                         <div className="w-full bg-slate-100 rounded-full h-1.5 overflow-hidden">
                           <div className="bg-indigo-500 h-full rounded-full transition-all duration-500 ease-out" style={{ width: `${progress}%` }}></div>
                         </div>
                      </div>
                    )}

                    {logs.length > 0 && (
                      <div className="px-6 py-4 border-t border-slate-100 bg-slate-900 text-slate-300 text-[11px] font-mono max-h-40 overflow-y-auto flex flex-col gap-1.5">
                        {logs.map((log, idx) => (
                          <div key={idx} className={`${log.type === 'err' ? 'text-rose-400' : log.type === 'ok' ? 'text-emerald-400' : 'text-slate-400'}`}>
                            {log.msg}
                          </div>
                        ))}
                      </div>
                    )}

                    <div className="px-6 py-4 border-t border-slate-100 bg-slate-50 flex justify-end">
                      <button 
                        onClick={handleGenerate}
                        disabled={generating || progress === 100}
                        className="bg-slate-900 text-white hover:bg-slate-800 font-medium rounded-md text-sm px-5 py-2 disabled:opacity-50 disabled:cursor-not-allowed transition-all shadow-sm flex items-center gap-2"
                      >
                        {generating && <Loader2 className="w-4 h-4 animate-spin" />}
                        {progress === 100 ? 'Generated' : 'Generate Certificates'}
                      </button>
                    </div>
                  </div>
                )}
              </>
            )}
          </motion.div>
        )}
      </main>
    </div>
  );
}
