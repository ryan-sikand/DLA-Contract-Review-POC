import { useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertTriangle,
  ChevronLeft,
  ChevronRight,
  CheckCircle2,
  ClipboardCheck,
  FileSearch,
  FileSpreadsheet,
  FileText,
  ListChecks,
  Maximize2,
  Play,
  RefreshCw,
  ShieldCheck,
  Trash2,
  Upload,
  X,
} from 'lucide-react';
import dlaLogo from '../DLA.jpg';
import { buildMaestroInputs, testSolutionConfig } from './config';
import { useAuth } from './hooks/useAuth';
import { completedRun, defaultSteps, initialDocumentSlots, mockFindings, runHistory } from './mockData';
import { buildMaestroInstanceUrl, loadRecentSolutionRuns, loadSolutionRunById, startReviewProcess } from './services/uipathSolution';
import type { AgentFinding, DocumentSlot, DocumentSlotId, ReviewRun, RuleResult, RunStep } from './types';

const runsStorageKey = 'dla-cqr-runs';
const pendingRunsStorageKey = 'dla-cqr-pending-runs';
const deletedRunsStorageKey = 'dla-cqr-deleted-runs';
const historyPageSize = 5;

const slotSampleNames: Record<DocumentSlotId, string[]> = {
  dd2579: ['DD2579-TE1.pdf', 'DD2579-TE2.pdf', 'DD2579-TE3.pdf', 'DD2579-TE4.pdf'],
  sf1449Award: ['SF1449-TE1-A.pdf', 'SF1449-TE2-A.pdf', 'SF1449-TE3-A.pdf', 'SF1449-TE4-A.pdf'],
  sf1449Solicitation: ['SF1449-TE1-S.pdf', 'SF1449-TE2-S.pdf', 'SF1449-TE3-S.pdf', 'SF1449-TE4-S.pdf'],
  saad: ['SAAD-TE1.pdf', 'SAAD-TE2.pdf', 'SAAD-TE3.pdf', 'SAAD-TE4.pdf'],
  dftr: ['DF-TE3.pdf', 'DF-TE4.pdf'],
};

function resultClass(result: RuleResult) {
  if (result === 'Pass') return 'result-pass';
  if (result === 'Flag') return 'result-flag';
  return 'result-na';
}

function stepIcon(status: RunStep['status']) {
  if (status === 'Complete') return <CheckCircle2 size={16} />;
  if (status === 'Running') return <RefreshCw className="spin" size={16} />;
  if (status === 'Error') return <AlertTriangle size={16} />;
  return <span className="pending-dot" />;
}

function makeRunningSteps(activeIndex: number): RunStep[] {
  return defaultSteps.map((step, index) => ({
    ...step,
    status: index < activeIndex ? 'Complete' : index === activeIndex ? 'Running' : 'Pending',
  }));
}

function loadStoredRuns(): ReviewRun[] {
  try {
    const raw = window.localStorage.getItem(runsStorageKey) ?? window.localStorage.getItem(pendingRunsStorageKey);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed)
      ? parsed.filter(isReviewRun).map((run) => ({ ...run, displayName: run.displayName || 'Contract Review Package' }))
      : [];
  } catch {
    return [];
  }
}

function saveStoredRuns(runItems: ReviewRun[]) {
  try {
    if (!runItems.length) {
      window.localStorage.removeItem(runsStorageKey);
      return;
    }
    window.localStorage.setItem(runsStorageKey, JSON.stringify(sortRuns(runItems).slice(0, 50)));
    window.localStorage.removeItem(pendingRunsStorageKey);
  } catch {
    // Ignore storage failures; live Maestro refresh remains the source of truth.
  }
}

function loadDeletedRunIds(): string[] {
  try {
    const raw = window.localStorage.getItem(deletedRunsStorageKey);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.filter((value) => typeof value === 'string') : [];
  } catch {
    return [];
  }
}

function saveDeletedRunIds(runIds: string[]) {
  try {
    const uniqueRunIds = [...new Set(runIds)];
    if (!uniqueRunIds.length) {
      window.localStorage.removeItem(deletedRunsStorageKey);
      return;
    }
    window.localStorage.setItem(deletedRunsStorageKey, JSON.stringify(uniqueRunIds));
  } catch {
    // Ignore storage failures; the current session state will still be updated.
  }
}

function mergeLiveRuns(liveRuns: ReviewRun[], currentRuns: ReviewRun[], startedJobId?: string) {
  const byKey = new Map<string, ReviewRun>();
  for (const run of currentRuns) {
    byKey.set(getRunMergeKey(run), run);
  }

  for (const liveRun of liveRuns) {
    const key = getRunMergeKey(liveRun);
    const existing = byKey.get(key) ?? findMatchingRun(liveRun, [...byKey.values()], startedJobId);
    if (existing) {
      byKey.delete(getRunMergeKey(existing));
      byKey.set(key, mergeRunData(existing, liveRun));
    } else {
      byKey.set(key, liveRun);
    }
  }

  return sortRuns([...byKey.values()]).slice(0, 50);
}

function getRunMergeKey(run: ReviewRun) {
  return run.latestRunId || run.instanceId || run.id;
}

function findMatchingRun(run: ReviewRun, candidates: ReviewRun[], startedJobId?: string) {
  const keys = new Set([run.id, run.instanceId, run.latestRunId].filter(Boolean));
  if (startedJobId && keys.has(startedJobId)) keys.add(startedJobId);
  return candidates.find((candidate) => [candidate.id, candidate.instanceId, candidate.latestRunId].some((value) => value && keys.has(value)));
}

function mergeRunData(existing: ReviewRun, liveRun: ReviewRun) {
  return {
    ...existing,
    ...liveRun,
    displayName: liveRun.displayName || existing.displayName,
    createdAt: liveRun.createdAt === 'Unknown time' ? existing.createdAt : liveRun.createdAt,
    documentsProcessed: liveRun.documentsProcessed || existing.documentsProcessed,
    findings: liveRun.findings.length ? liveRun.findings : existing.findings,
    reportPath: liveRun.reportPath === 'Pending' ? existing.reportPath : liveRun.reportPath,
    sharePointItem: liveRun.sharePointItem === 'Pending' ? existing.sharePointItem : liveRun.sharePointItem,
    reportUrl: liveRun.reportUrl || existing.reportUrl,
  };
}

function sortRuns(runItems: ReviewRun[]) {
  return [...runItems].sort((a, b) => getRunSortValue(b) - getRunSortValue(a));
}

function getRunSortValue(runItem: ReviewRun) {
  if (runItem.status === 'Running' && !runItem.instanceId) return Date.now() + 1;
  if (runItem.createdAt === 'Running now') return Date.now();
  const parsed = Date.parse(runItem.createdAt);
  return Number.isNaN(parsed) ? 0 : parsed;
}

function pickSelectedRun(runs: ReviewRun[], selectedRunId: string, preserveSelection?: boolean) {
  if (preserveSelection) {
    const current = runs.find((run) => run.id === selectedRunId || run.latestRunId === selectedRunId);
    if (current) return current;
  }
  return runs[0];
}

function isLikelyStartedRun(run: ReviewRun, startedAt: Date) {
  const createdAt = new Date(run.createdAt);
  if (Number.isNaN(createdAt.getTime())) return false;
  return createdAt.getTime() >= startedAt.getTime() - 30_000;
}

function isReviewRun(value: unknown): value is ReviewRun {
  return Boolean(value && typeof value === 'object' && typeof (value as ReviewRun).id === 'string');
}

function isDeletedRun(run: ReviewRun, deletedRunIds: string[]) {
  return [run.id, run.instanceId, run.latestRunId].some((value) => value && deletedRunIds.includes(value));
}

function buildRunDisplayName(documents: DocumentSlot[]) {
  const tokenCounts = new Map<string, number>();
  for (const document of documents) {
    const match = document.fileName.match(/\bTE\s*[-_ ]?(\d+)\b/i) ?? document.fileName.match(/\bTE(\d+)\b/i);
    if (!match) continue;
    const token = `TE${match[1]}`;
    tokenCounts.set(token, (tokenCounts.get(token) ?? 0) + 1);
  }
  const [bestToken] =
    [...tokenCounts.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))[0] ?? [];
  return bestToken ? `Contract Review Package ${bestToken}` : 'Contract Review Package';
}

function App() {
  const { sdk, isAuthenticated, isConfigured, isLoading: isAuthLoading, error: authError, login, logout } = useAuth();
  const [documents, setDocuments] = useState<DocumentSlot[]>(initialDocumentSlots);
  const [runs, setRuns] = useState<ReviewRun[]>(() => loadStoredRuns());
  const [deletedRunIds, setDeletedRunIds] = useState<string[]>(() => loadDeletedRunIds());
  const [selectedRunId, setSelectedRunId] = useState(completedRun.id);
  const [selectedCheck, setSelectedCheck] = useState(mockFindings[1].check);
  const [activeFilter, setActiveFilter] = useState<'All' | RuleResult>('All');
  const [historyPage, setHistoryPage] = useState(0);
  const [isRunning, setIsRunning] = useState(false);
  const [isIntakeOpen, setIsIntakeOpen] = useState(false);
  const [liveStatus, setLiveStatus] = useState<string>('Connected to mock data. Sign in to query the solution folder.');
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [maestroModalUrl, setMaestroModalUrl] = useState('');
  const pollingRef = useRef<number | undefined>(undefined);
  const runningRefreshRef = useRef<number | undefined>(undefined);
  const selectedRunIdRef = useRef(selectedRunId);

  const requiredDocsReady = documents.filter((doc) => doc.required).every((doc) => doc.fileName.trim().length > 0);
  const visibleRuns = (runs.length ? runs : runHistory).filter((item) => !isDeletedRun(item, deletedRunIds));
  const historyPageCount = Math.max(1, Math.ceil(visibleRuns.length / historyPageSize));
  const currentHistoryPage = Math.min(historyPage, historyPageCount - 1);
  const historyStart = currentHistoryPage * historyPageSize;
  const pagedRuns = visibleRuns.slice(historyStart, historyStart + historyPageSize);
  const run = visibleRuns.find((item) => item.id === selectedRunId) ?? visibleRuns[0] ?? completedRun;
  const selectedFinding =
    run.findings.find((finding) => finding.check === selectedCheck) ??
    run.findings[0] ??
    ({
      check: 'Pending agent output',
      result: 'Not Applicable',
      recommendation: 'Waiting for automation',
      fieldsReviewed: 'Pending',
      valuesCompared: 'Pending',
      summary: 'The agent findings will appear after the automation completes and the run is refreshed.',
      action: 'Refresh the solution run after completion.',
      notes: '',
      sources: 'Pending',
    } satisfies AgentFinding);
  const filteredFindings = activeFilter === 'All' ? run.findings : run.findings.filter((finding) => finding.result === activeFilter);

  const summary = useMemo(() => {
    return run.findings.reduce(
      (counts, finding) => {
        counts[finding.result] += 1;
        return counts;
      },
      { Pass: 0, Flag: 0, 'Not Applicable': 0 } as Record<RuleResult, number>,
    );
  }, [run.findings]);

  useEffect(() => {
    if (!isAuthenticated) return;
    void refreshLiveData({ preserveSelection: true });
  }, [isAuthenticated]);

  useEffect(() => {
    selectedRunIdRef.current = selectedRunId;
  }, [selectedRunId]);

  useEffect(() => {
    saveStoredRuns(runs.filter((item) => !isDeletedRun(item, deletedRunIds)));
  }, [runs, deletedRunIds]);

  useEffect(() => {
    saveDeletedRunIds(deletedRunIds);
  }, [deletedRunIds]);

  useEffect(() => {
    if (!maestroModalUrl) return;

    const scrollY = window.scrollY;
    const previousBodyStyles = {
      position: document.body.style.position,
      top: document.body.style.top,
      width: document.body.style.width,
      overflow: document.body.style.overflow,
    };

    document.body.style.position = 'fixed';
    document.body.style.top = `-${scrollY}px`;
    document.body.style.width = '100%';
    document.body.style.overflow = 'hidden';

    return () => {
      document.body.style.position = previousBodyStyles.position;
      document.body.style.top = previousBodyStyles.top;
      document.body.style.width = previousBodyStyles.width;
      document.body.style.overflow = previousBodyStyles.overflow;
      window.scrollTo(0, scrollY);
    };
  }, [maestroModalUrl]);

  useEffect(() => {
    setHistoryPage((current) => Math.min(current, Math.max(0, Math.ceil(visibleRuns.length / historyPageSize) - 1)));
  }, [visibleRuns.length]);

  useEffect(() => {
    return () => {
      if (pollingRef.current) window.clearInterval(pollingRef.current);
      if (runningRefreshRef.current) window.clearInterval(runningRefreshRef.current);
    };
  }, []);

  useEffect(() => {
    if (!isAuthenticated || !runs.some((item) => item.status === 'Running')) {
      if (runningRefreshRef.current) {
        window.clearInterval(runningRefreshRef.current);
        runningRefreshRef.current = undefined;
      }
      return;
    }

    if (runningRefreshRef.current) return;
    runningRefreshRef.current = window.setInterval(() => {
      void refreshLiveData({ preserveSelection: true, quiet: true });
    }, 15000);
  }, [isAuthenticated, runs]);

  if (isConfigured && isAuthLoading) {
    return (
      <main className="auth-gate">
        <div className="auth-gate-panel">
          <div className="brand-mark">
            <ShieldCheck size={22} />
          </div>
          <p className="eyebrow">DLA Contract</p>
          <h1>Signing in to UiPath</h1>
          <p>Connecting to the solution folder before opening the review workspace.</p>
          <RefreshCw className="spin" size={22} />
        </div>
      </main>
    );
  }

  function updateDocument(id: DocumentSlotId, fileName: string) {
    setDocuments((current) => current.map((doc) => (doc.id === id ? { ...doc, fileName } : doc)));
  }

  function handleFileChange(id: DocumentSlotId, file: File | undefined) {
    if (!file) return;
    updateDocument(id, file.name);
  }

  function clearDocument(id: DocumentSlotId) {
    updateDocument(id, '');
  }

  async function refreshLiveData(options: { preserveSelection?: boolean; quiet?: boolean } = {}) {
    if (!isAuthenticated) {
      setLiveStatus('Sign in with UiPath before refreshing live runs and Excel reports.');
      return;
    }

    if (!options.quiet) {
      setIsRefreshing(true);
      setLiveStatus(`Reading jobs and reports from ${testSolutionConfig.folderPath}.`);
    }
    try {
      const liveRuns = await loadRecentSolutionRuns(sdk);
      if (liveRuns.length) {
        setRuns((current) => {
          const mergedRuns = mergeLiveRuns(
            liveRuns.filter((item) => !isDeletedRun(item, deletedRunIds)),
            current.filter((item) => !isDeletedRun(item, deletedRunIds)),
          );
          const nextRun = pickSelectedRun(mergedRuns, selectedRunIdRef.current, options.preserveSelection);
          setSelectedRunId(nextRun.id);
          setSelectedCheck(nextRun.findings[0]?.check ?? mockFindings[0].check);
          return mergedRuns;
        });
      }
      setLiveStatus(`Loaded ${liveRuns.length} live Maestro process instances from the solution folder.`);
    } catch (error) {
      setLiveStatus(error instanceof Error ? error.message : 'Unable to refresh live UiPath data.');
    } finally {
      if (!options.quiet) setIsRefreshing(false);
    }
  }

  async function runAutomation() {
    if (!requiredDocsReady || isRunning) return;

    if (isConfigured && !isAuthenticated) {
      setLiveStatus('Sign in with UiPath before starting a live solution run.');
      await login();
      return;
    }

    setIsRunning(true);
    setLiveStatus('Starting the solution automation in UiPath.');

    if (isAuthenticated) {
      try {
        const inputArguments = buildMaestroInputs(documents);
        const startedAt = new Date();
        const job = await startReviewProcess(sdk, inputArguments);
        const jobId = String((job as { key?: string }).key ?? `DLA-CQR-${new Date().toISOString().replace(/\D/g, '').slice(0, 14)}`);
        const runningRun: ReviewRun = {
          ...completedRun,
          id: jobId,
          displayName: buildRunDisplayName(documents),
          latestRunId: jobId,
          status: 'Running',
          createdAt: 'Running now',
          documentsProcessed: documents.filter((doc) => doc.fileName).length,
          steps: makeRunningSteps(0),
          findings: [],
          overallStatus: 'Not Applicable',
          reportPath: 'Pending',
          sharePointItem: 'Pending',
          packageName: `${testSolutionConfig.processName} ${testSolutionConfig.processVersion}`,
          instanceId: jobId,
          maestroUrl: buildMaestroInstanceUrl(jobId),
        };

        setLiveStatus(`Started solution job ${jobId}. Refresh in a minute to pull the latest status and report.`);
        setIsIntakeOpen(false);
        setRuns((current) => sortRuns([runningRun, ...current]));
        setHistoryPage(0);
        setSelectedRunId(jobId);
        setIsRunning(false);
        startPolling(jobId, startedAt);
        return;
      } catch (error) {
        setLiveStatus(`Live start failed: ${error instanceof Error ? error.message : 'unknown error'}`);
        setIsRunning(false);
        return;
      }
    } else {
      setLiveStatus('UiPath is not signed in. The automation was not started.');
      setIsRunning(false);
      return;
    }
  }

  function openMaestro(url: string) {
    window.open(url, '_blank', 'noopener,noreferrer');
  }

  function startPolling(startedJobId: string, startedAt: Date) {
    if (pollingRef.current) window.clearInterval(pollingRef.current);
    let attempts = 0;
    pollingRef.current = window.setInterval(() => {
      attempts += 1;
      void Promise.all([loadSolutionRunById(sdk, startedJobId), loadRecentSolutionRuns(sdk)])
        .then(([exactRun, recentRuns]) => {
          const liveRuns = exactRun ? [exactRun, ...recentRuns.filter((runItem) => runItem.id !== exactRun.id)] : recentRuns;
          if (!liveRuns.length) return;
          const selectedLiveRun =
            liveRuns.find((liveRun) => liveRun.latestRunId === startedJobId) ??
            liveRuns.find((liveRun) => liveRun.id === startedJobId) ??
            liveRuns.find((liveRun) => isLikelyStartedRun(liveRun, startedAt)) ??
            liveRuns[0];
          setRuns((current) =>
            mergeLiveRuns(
              liveRuns.filter((item) => !isDeletedRun(item, deletedRunIds)),
              current.filter((item) => !isDeletedRun(item, deletedRunIds)),
              startedJobId,
            ),
          );
          if (selectedRunIdRef.current === startedJobId || selectedRunIdRef.current === selectedLiveRun.id) {
            setSelectedRunId(selectedLiveRun.id);
            setSelectedCheck(selectedLiveRun.findings[0]?.check ?? mockFindings[0].check);
          }
          setLiveStatus(`Updated from Maestro instance ${selectedLiveRun.id}.`);
          if (selectedLiveRun.status === 'Complete' || selectedLiveRun.status === 'Failed' || attempts >= 360) {
            if (pollingRef.current) window.clearInterval(pollingRef.current);
            pollingRef.current = undefined;
          }
        })
        .catch((error) => {
          setLiveStatus(error instanceof Error ? error.message : 'Unable to poll Maestro process instance.');
        });
    }, 5000);
  }

  function selectRun(id: string) {
    setSelectedRunId(id);
    const nextRun = visibleRuns.find((item) => item.id === id);
    setSelectedCheck(nextRun?.findings[0]?.check ?? mockFindings[0].check);
  }

  function deleteRun(id: string) {
    const runToDelete = visibleRuns.find((item) => item.id === id);
    const idsToDelete = [runToDelete?.id, runToDelete?.instanceId, runToDelete?.latestRunId].filter(Boolean) as string[];
    setDeletedRunIds((current) => [...new Set([...current, ...idsToDelete])]);
    setRuns((current) => current.filter((item) => !idsToDelete.includes(item.id) && !idsToDelete.includes(item.instanceId ?? '') && !idsToDelete.includes(item.latestRunId ?? '')));
    if (selectedRunId === id) {
      const nextRun = visibleRuns.find((item) => item.id !== id);
      if (nextRun) {
        setSelectedRunId(nextRun.id);
        setSelectedCheck(nextRun.findings[0]?.check ?? mockFindings[0].check);
      }
    }
  }

  function openSelectedMaestro(runItem: ReviewRun) {
    if (!runItem.instanceId) {
      setLiveStatus('Waiting for Maestro to expose the process instance. The app will enable this link as soon as polling finds it.');
      return;
    }
    setMaestroModalUrl(runItem.maestroUrl);
  }

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div className="brand-lockup">
          <img className="dla-logo" src={dlaLogo} alt="Defense Logistics Agency" />
          <div>
            <p className="eyebrow">DLA Contract</p>
            <h1>Quality Review</h1>
          </div>
        </div>

        <nav className="nav-stack" aria-label="Application views">
          <button className="nav-item active" type="button">
            <ClipboardCheck size={18} />
            Review Workspace
          </button>
        </nav>

        <section className="side-panel">
          <p className="panel-label">Environment</p>
          <strong>{testSolutionConfig.label}</strong>
          <span>{testSolutionConfig.folderPath}</span>
          <span>{isAuthenticated ? 'Signed in to UiPath' : 'Mock fallback available'}</span>
        </section>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <div>
            <p className="eyebrow">Analyst console</p>
            <h2>Contract Quality Review Results</h2>
          </div>
          <div className="topbar-actions">
            {isAuthenticated ? (
              <button className="secondary-button" type="button" onClick={logout}>
                Sign Out
              </button>
            ) : (
              <button className="secondary-button" disabled={!isConfigured || isAuthLoading} type="button" onClick={login}>
                {isAuthLoading ? <RefreshCw className="spin" size={17} /> : null}
                Sign In
              </button>
            )}
            <button className="secondary-button" disabled={isRefreshing} type="button" onClick={() => void refreshLiveData()}>
              {isRefreshing ? <RefreshCw className="spin" size={17} /> : <RefreshCw size={17} />}
              Refresh
            </button>
            <button className="primary-button" onClick={() => setIsIntakeOpen(true)} type="button">
              <Play size={17} />
              Run Review
            </button>
          </div>
        </header>

        <section className="live-banner">
          <strong>Main solution data</strong>
          <span>{liveStatus}</span>
          {authError ? <span className="banner-error">{authError}</span> : null}
        </section>

        <section className="metrics-grid">
          <div className="metric">
            <span>Overall Status</span>
            <strong className={resultClass(run.overallStatus)}>{run.overallStatus}</strong>
          </div>
          <div className="metric">
            <span>Documents</span>
            <strong>{run.documentsProcessed}</strong>
          </div>
          <div className="metric">
            <span>Pass</span>
            <strong>{summary.Pass}</strong>
          </div>
          <div className="metric">
            <span>Flag</span>
            <strong>{summary.Flag}</strong>
          </div>
          <div className="metric">
            <span>Review Package</span>
            <strong className="run-id">{run.displayName}</strong>
          </div>
        </section>

        <div className="content-grid">
          <section className="panel run-history-panel">
            <div className="section-title">
              <div>
                <p className="eyebrow">Automation history</p>
                <h3>Runs</h3>
              </div>
              <span className="history-count">
                {visibleRuns.length ? `${historyStart + 1}-${Math.min(historyStart + historyPageSize, visibleRuns.length)} of ${visibleRuns.length}` : '0 of 0'}
              </span>
            </div>

            <div className="run-list">
              {pagedRuns.map((item) => (
                <article className={`run-list-item ${item.id === selectedRunId ? 'active' : ''}`} key={item.id}>
                  <button className="run-list-select" onClick={() => selectRun(item.id)} type="button">
                  <div className="run-list-icon">
                    {item.status === 'Running' ? <RefreshCw className="spin" size={18} /> : <ListChecks size={18} />}
                  </div>
                  <div>
                    <div className="run-list-heading">
                      <strong>{item.displayName || 'Contract Review Package'}</strong>
                      <span className={resultClass(item.overallStatus)}>{item.status === 'Running' ? 'Running' : item.overallStatus}</span>
                    </div>
                    <p>{item.createdAt} · {item.documentsProcessed} documents · {item.packageName}</p>
                  </div>
                  </button>
                  <button className="run-open-button" type="button" disabled={!item.instanceId} title="Open Maestro instance" onClick={() => openSelectedMaestro(item)}>
                    <Maximize2 size={15} />
                  </button>
                  <button className="run-delete-button" type="button" title="Remove run from this app" onClick={() => deleteRun(item.id)}>
                    <Trash2 size={15} />
                  </button>
                </article>
              ))}
            </div>
            {historyPageCount > 1 ? (
              <div className="history-pagination" aria-label="Automation history pages">
                <button
                  className="pager-button"
                  type="button"
                  disabled={currentHistoryPage === 0}
                  onClick={() => setHistoryPage((page) => Math.max(0, page - 1))}
                >
                  <ChevronLeft size={15} />
                  Previous
                </button>
                <span>Page {currentHistoryPage + 1} of {historyPageCount}</span>
                <button
                  className="pager-button"
                  type="button"
                  disabled={currentHistoryPage >= historyPageCount - 1}
                  onClick={() => setHistoryPage((page) => Math.min(historyPageCount - 1, page + 1))}
                >
                  Next
                  <ChevronRight size={15} />
                </button>
              </div>
            ) : null}
          </section>

          <section className="panel run-panel">
            <div className="section-title">
              <div>
                <p className="eyebrow">Automation</p>
                <h3>Run Status</h3>
              </div>
              <span className={`status-pill ${run.status.toLowerCase()}`}>{run.status}</span>
            </div>

            <button className="maestro-button" type="button" disabled={!run.instanceId} onClick={() => openSelectedMaestro(run)}>
              <Maximize2 size={16} />
              {run.instanceId ? 'Open Maestro Instance' : 'Waiting for Maestro Instance'}
            </button>

            <div className="timeline">
              {run.steps.map((step) => (
                <div className={`timeline-step ${step.status.toLowerCase()}`} key={step.label}>
                  <div className="timeline-icon">{stepIcon(step.status)}</div>
                  <div>
                    <strong>{step.label}</strong>
                    <p>{step.detail}</p>
                  </div>
                </div>
              ))}
            </div>

            <div className="report-card">
              <div className="doc-icon large">
                <FileSpreadsheet size={22} />
              </div>
              <div>
                <span>Latest Excel Report</span>
                <strong>{run.reportPath}</strong>
                <p>SharePoint item: {run.sharePointItem}</p>
                {run.reportUrl ? (
                  <button className="text-link-button" type="button" onClick={() => window.open(run.reportUrl, '_blank', 'noopener,noreferrer')}>
                    Open SharePoint Excel report
                  </button>
                ) : null}
              </div>
            </div>
          </section>
        </div>

        <section className="panel findings-panel">
          <div className="section-title">
            <div>
              <p className="eyebrow">Agent output</p>
              <h3>Business Rule Findings</h3>
            </div>
            <div className="segmented-control" aria-label="Filter findings">
              {(['All', 'Pass', 'Flag', 'Not Applicable'] as const).map((filter) => (
                <button
                  className={activeFilter === filter ? 'selected' : ''}
                  key={filter}
                  onClick={() => setActiveFilter(filter)}
                  type="button"
                >
                  {filter}
                </button>
              ))}
            </div>
          </div>

          <div className="findings-layout">
            <div className="findings-table-wrap">
              <table className="findings-table">
                <thead>
                  <tr>
                    <th>Check</th>
                    <th>Result</th>
                    <th>Recommendation</th>
                    <th>Result Summary</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredFindings.map((finding) => (
                    <tr
                      className={finding.check === selectedFinding.check ? 'selected-row' : ''}
                      key={finding.check}
                      onClick={() => setSelectedCheck(finding.check)}
                    >
                      <td>{finding.check}</td>
                      <td>
                        <span className={`result-chip ${resultClass(finding.result)}`}>
                          {finding.result === 'Flag' ? <AlertTriangle size={14} /> : <CheckCircle2 size={14} />}
                          {finding.result}
                        </span>
                      </td>
                      <td>{finding.recommendation}</td>
                      <td>{finding.summary}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <aside className="finding-detail">
              <div className="detail-header">
                <FileSearch size={20} />
                <div>
                  <span>Selected Finding</span>
                  <h4>{selectedFinding.check}</h4>
                </div>
              </div>
              <dl>
                <div>
                  <dt>Fields Reviewed</dt>
                  <dd>{selectedFinding.fieldsReviewed}</dd>
                </div>
                <div>
                  <dt>Values Compared</dt>
                  <dd>{selectedFinding.valuesCompared}</dd>
                </div>
                <div>
                  <dt>Recommended Action</dt>
                  <dd>{selectedFinding.action}</dd>
                </div>
                <div>
                  <dt>Data Completeness Notes</dt>
                  <dd>{selectedFinding.notes || 'No completeness gaps noted.'}</dd>
                </div>
                <div>
                  <dt>Source Documents</dt>
                  <dd>{selectedFinding.sources}</dd>
                </div>
              </dl>
            </aside>
          </div>
        </section>
      </section>

      {isIntakeOpen && (
        <div className="modal-backdrop" role="dialog" aria-modal="true" aria-label="Run review input package">
          <section className="modal-panel">
            <div className="section-title">
              <div>
                <p className="eyebrow">Input package</p>
                <h3>Source Documents</h3>
              </div>
              <span className={requiredDocsReady ? 'ready-pill' : 'missing-pill'}>
                {requiredDocsReady ? 'Ready' : 'Missing required files'}
              </span>
            </div>

            <div className="document-list">
              {documents.map((doc) => (
                <article className="document-row" key={doc.id}>
                  <div className="doc-icon">
                    <FileText size={18} />
                  </div>
                  <div className="doc-main">
                    <div className="doc-heading">
                      <strong>{doc.label}</strong>
                      {doc.required ? <span>Required</span> : <span>Optional</span>}
                    </div>
                    <p>{doc.helper}</p>
                    <div className="doc-input-row">
                      <input
                        aria-label={`${doc.label} file name`}
                        value={doc.fileName}
                        onChange={(event) => updateDocument(doc.id, event.target.value)}
                        placeholder="Select or type a file name"
                      />
                      <label className="upload-button" title={`Upload ${doc.label}`}>
                        <Upload size={16} />
                        <input type="file" accept="application/pdf" onChange={(event) => handleFileChange(doc.id, event.target.files?.[0])} />
                      </label>
                      <button className="clear-button" type="button" onClick={() => clearDocument(doc.id)}>
                        Clear
                      </button>
                    </div>
                    <div className="sample-list">
                      {slotSampleNames[doc.id].map((name) => (
                        <button key={name} type="button" onClick={() => updateDocument(doc.id, name)}>
                          {name.replace('.pdf', '')}
                        </button>
                      ))}
                    </div>
                  </div>
                </article>
              ))}
            </div>

            <div className="modal-actions">
              <button className="secondary-button" type="button" onClick={() => setIsIntakeOpen(false)}>
                Cancel
              </button>
              <button className="primary-button" disabled={!requiredDocsReady || isRunning} onClick={runAutomation} type="button">
                {isRunning ? <RefreshCw className="spin" size={17} /> : <Play size={17} />}
                {isRunning ? 'Running' : 'Start Automation'}
              </button>
            </div>
          </section>
        </div>
      )}

      {maestroModalUrl && (
        <div
          className="maestro-modal-backdrop"
          role="dialog"
          aria-modal="true"
          aria-label="Maestro process instance"
          onWheel={(event) => event.stopPropagation()}
          onTouchMove={(event) => event.stopPropagation()}
        >
          <section className="maestro-modal-panel" onWheel={(event) => event.stopPropagation()} onTouchMove={(event) => event.stopPropagation()}>
            <header className="maestro-modal-header">
              <div>
                <p className="eyebrow">Maestro</p>
                <h3>Process Instance</h3>
              </div>
              <div className="maestro-modal-actions">
                <button className="secondary-button" type="button" onClick={() => openMaestro(maestroModalUrl)}>
                  <Maximize2 size={16} />
                  Open New Tab
                </button>
                <button className="icon-button" type="button" title="Close Maestro viewer" onClick={() => setMaestroModalUrl('')}>
                  <X size={18} />
                </button>
              </div>
            </header>
            <iframe className="maestro-frame" src={maestroModalUrl} title="Maestro process instance" />
          </section>
        </div>
      )}
    </main>
  );
}

export default App;
