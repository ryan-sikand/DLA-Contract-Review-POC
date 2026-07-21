import { ProcessInstances, type ProcessInstanceExecutionHistoryResponse } from '@uipath/uipath-typescript/maestro-processes';
import { Processes, StartStrategy } from '@uipath/uipath-typescript/processes';
import { Jobs } from '@uipath/uipath-typescript/jobs';
import type { UiPath } from '@uipath/uipath-typescript';
import { defaultSteps } from '../mockData';
import type { AgentFinding, ReviewRun, RuleResult, RunStep, StepStatus } from '../types';
import { testSolutionConfig } from '../config';

type ProcessInstanceShape = {
  instanceId: string;
  packageVersion?: string;
  latestRunId?: string;
  latestRunStatus?: string;
  startedTime?: string;
  completedTime?: string | null;
  instanceDisplayName?: string;
  getVariables?: () => Promise<VariableResponse>;
};

type JobShape = {
  key: string;
  state: string;
  processName?: string | null;
  releaseName?: string | null;
  startTime?: string | null;
  startedTime?: string | null;
  endTime?: string | null;
  createdTime?: string | null;
  creationTime?: string | null;
  processVersion?: string | null;
  inputArguments?: string | null;
  outputArguments?: string | null;
  parentJobKey?: string | null;
  jobError?: unknown;
};

type StepProgress = {
  duStarted: boolean;
  duComplete: boolean;
  duFailed: boolean;
  consolidateStarted: boolean;
  consolidateComplete: boolean;
  consolidateFailed: boolean;
  agentStarted: boolean;
  agentComplete: boolean;
  agentFailed: boolean;
  exportStarted: boolean;
  exportComplete: boolean;
  exportFailed: boolean;
};

type VariableResponse = {
  globalVariables?: Array<{ name?: string; value?: unknown }>;
  elements?: Array<{ elementId?: string; outputs?: Record<string, unknown>; inputs?: Record<string, unknown> }>;
};

type MaestroInternalClient = {
  get<T>(path: string, options?: { headers?: Record<string, string> }): Promise<{ data: T }>;
};

type ElementExecutionsResponse = {
  elementExecutions?: ElementExecutionShape[];
};

type ElementExecutionShape = {
  name?: string;
  elementName?: string;
  displayName?: string;
  source?: string;
  elementId?: string;
  status?: string;
  state?: string;
  startedTime?: string | null;
  endTime?: string | null;
  completedTime?: string | null;
  elementRuns?: ElementRunShape[];
  runs?: ElementRunShape[];
};

type ElementRunShape = {
  status?: string;
  state?: string;
  startedTime?: string | null;
  endTime?: string | null;
  completedTime?: string | null;
};

export async function startReviewProcess(sdk: UiPath, inputArguments: Record<string, unknown>) {
  const processes = new Processes(sdk);
  const jobs = await processes.start(
    {
      processKey: testSolutionConfig.processKey,
      strategy: StartStrategy.ModernJobsCount,
      jobsCount: 1,
      inputArguments: JSON.stringify(inputArguments),
    },
    testSolutionConfig.folderId,
  );

  return jobs[0];
}

export async function loadRecentSolutionRuns(sdk: UiPath): Promise<ReviewRun[]> {
  const processInstances = new ProcessInstances(sdk);
  const jobs = new Jobs(sdk);
  const allJobs = await loadRecentJobs(jobs);
  const instances = (await processInstances.getAll({
    processKey: testSolutionConfig.processKey,
    pageSize: 50,
  } as never)) as { items?: ProcessInstanceShape[] };

  const recent = (instances.items ?? [])
    .sort((a, b) => String(b.startedTime ?? '').localeCompare(String(a.startedTime ?? '')))
    .slice(0, 25);

  const agenticJobs = await loadRecentAgenticJobs(jobs, allJobs);
  const instanceRuns = await Promise.all(recent.map((instance) => mapInstanceToReviewRun(instance, processInstances)));
  const jobRuns = await Promise.all(agenticJobs.map((job) => mapJobToReviewRun(job, processInstances, jobs, getChildProgress(job, allJobs))));

  return dedupeAndSortRuns([...jobRuns, ...instanceRuns]).slice(0, 25);
}

export async function loadSolutionRunById(sdk: UiPath, jobKey: string): Promise<ReviewRun | null> {
  const processInstances = new ProcessInstances(sdk);
  const jobs = new Jobs(sdk);
  const allJobs = await loadRecentJobs(jobs);

  try {
    const job = (await jobs.getById(jobKey, testSolutionConfig.folderId)) as unknown as JobShape;
    return mapJobToReviewRun(job, processInstances, jobs, getChildProgress(job, allJobs));
  } catch {
    try {
      const instance = (await processInstances.getById(jobKey, testSolutionConfig.folderKey)) as unknown as ProcessInstanceShape;
      return mapInstanceToReviewRun(instance, processInstances);
    } catch {
      return null;
    }
  }
}

async function mapInstanceToReviewRun(instance: ProcessInstanceShape, processInstances: ProcessInstances): Promise<ReviewRun> {
  const variables = await safeGetVariables(instance, processInstances);
  const valueByName = buildVariableMap(variables);
  const state = instance.latestRunStatus ?? 'Unknown';
  const finished = isCompleted(state);
  const faulted = isFaulted(state);
  const findings = rowsToFindings(parseExcelRows(getReportRowsJson(valueByName)));
  const overallStatus: RuleResult = faulted ? 'Flag' : finished ? 'Pass' : 'Not Applicable';
  const sharePointStatus = readString(valueByName, 'sharePointUploadStatus');
  const report = buildSharePointReport(sharePointStatus, readString(valueByName, 'localWorkbookPath'), readString(valueByName, 'bucketFilePath'));
  const inputDocuments = getInputDocumentNames(valueByName);
  const populatedFindings = findings.length ? findings : pendingFinding(finished, sharePointStatus);
  const progress = await safeGetExecutionProgress(instance.instanceId, processInstances, isRunning(state));

  return {
    id: instance.instanceId,
    displayName: buildRunDisplayName(inputDocuments),
    status: finished ? 'Complete' : isRunning(state) ? 'Running' : 'Draft',
    createdAt: formatDate(instance.startedTime),
    packageName: `${testSolutionConfig.processName} ${instance.packageVersion ?? testSolutionConfig.processVersion}`,
    analyst: 'UiPath user',
    overallStatus: findings.length ? summarizeFindings(findings) : overallStatus,
    documentsProcessed: inputDocuments.length,
    reportPath: report.label,
    sharePointItem: report.itemName || sharePointStatus || 'Pending',
    reportUrl: report.url,
    instanceId: instance.instanceId,
    latestRunId: instance.latestRunId,
    maestroUrl: buildMaestroInstanceUrl(instance.instanceId),
    steps: buildStepStatuses(valueByName, state, progress),
    findings: populatedFindings,
  };
}

async function mapJobToReviewRun(job: JobShape, processInstances: ProcessInstances, jobs: Jobs, childProgress?: StepProgress): Promise<ReviewRun> {
  const variables = await safeGetVariablesFromJob(job.key, processInstances, jobs);
  const inputVariables = inputArgumentsToVariables(job.inputArguments);
  const valueByName = buildVariableMap(variables);
  const inputValueByName = buildVariableMap(inputVariables);
  for (const [key, value] of inputValueByName) {
    if (!valueByName.has(key)) valueByName.set(key, value);
  }
  const state = job.state ?? 'Unknown';
  const finished = isCompleted(state);
  const faulted = isFaulted(state) || Boolean(childProgress?.duFailed || childProgress?.agentFailed || childProgress?.exportFailed);
  const findings = rowsToFindings(parseExcelRows(getReportRowsJson(valueByName)));
  const overallStatus: RuleResult = faulted ? 'Flag' : finished ? 'Pass' : 'Not Applicable';
  const sharePointStatus = readString(valueByName, 'sharePointUploadStatus');
  const report = buildSharePointReport(sharePointStatus, readString(valueByName, 'localWorkbookPath'), readString(valueByName, 'bucketFilePath'));
  const inputDocuments = getInputDocumentNames(valueByName);
  const populatedFindings = findings.length ? findings : pendingFinding(finished, sharePointStatus);
  const historyProgress = await safeGetExecutionProgress(job.key, processInstances, isRunning(state));
  const progress = mergeProgress(childProgress, historyProgress);

  return {
    id: job.key,
    displayName: buildRunDisplayName(inputDocuments),
    status: faulted ? 'Failed' : finished || progress?.exportComplete ? 'Complete' : isRunning(state) ? 'Running' : 'Draft',
    createdAt: formatDate(getJobTimestamp(job)),
    packageName: `${testSolutionConfig.processName} ${job.processVersion ?? testSolutionConfig.processVersion}`,
    analyst: 'UiPath user',
    overallStatus: findings.length ? summarizeFindings(findings) : overallStatus,
    documentsProcessed: inputDocuments.length || 0,
    reportPath: report.label,
    sharePointItem: report.itemName || sharePointStatus || 'Pending',
    reportUrl: report.url,
    instanceId: job.key,
    latestRunId: job.key,
    maestroUrl: buildMaestroInstanceUrl(job.key),
    steps: buildStepStatuses(valueByName, state, progress),
    findings: populatedFindings,
  };
}

async function safeGetVariables(instance: ProcessInstanceShape, processInstances: ProcessInstances): Promise<VariableResponse> {
  try {
    if (instance.getVariables) return await instance.getVariables();
  } catch {
    // Fall back to the service method below.
  }

  try {
    return await processInstances.getVariables(instance.instanceId, testSolutionConfig.folderKey);
  } catch {
    return {};
  }
}

async function safeGetVariablesFromJob(jobKey: string, processInstances: ProcessInstances, jobs: Jobs): Promise<VariableResponse> {
  try {
    const variables = await processInstances.getVariables(jobKey, testSolutionConfig.folderKey);
    if (hasVariables(variables)) return variables;
  } catch {
    // Fall back to Orchestrator output arguments below.
  }

  try {
    const output = await jobs.getOutput(jobKey, testSolutionConfig.folderId);
    return outputToVariables(output);
  } catch {
    return {};
  }
}

async function safeGetExecutionProgress(instanceId: string, processInstances: ProcessInstances, shouldLoad: boolean): Promise<StepProgress | undefined> {
  if (!shouldLoad) return undefined;

  const elementProgress = await safeGetElementExecutionProgress(instanceId, processInstances);
  if (elementProgress) return elementProgress;

  try {
    const history = await processInstances.getExecutionHistory(instanceId, testSolutionConfig.folderKey);
    return buildProgressFromHistory(history);
  } catch {
    return undefined;
  }
}

async function safeGetElementExecutionProgress(instanceId: string, processInstances: ProcessInstances): Promise<StepProgress | undefined> {
  try {
    const client = processInstances as unknown as MaestroInternalClient;
    const response = await client.get<ElementExecutionsResponse>(`pims_/api/v1/instances/${instanceId}/element-executions`, {
      headers: { 'X-UIPATH-FolderKey': testSolutionConfig.folderKey },
    });
    const elements = response.data.elementExecutions ?? [];
    return elements.length ? buildProgressFromElementExecutions(elements) : undefined;
  } catch {
    return undefined;
  }
}

async function loadRecentJobs(jobs: Jobs): Promise<JobShape[]> {
  const response = (await jobs.getAll({
    folderId: testSolutionConfig.folderId,
    pageSize: 100,
  } as never)) as { items?: JobShape[] };

  return response.items ?? [];
}

async function loadRecentAgenticJobs(jobs: Jobs, allJobs: JobShape[]): Promise<JobShape[]> {
  const summaries = allJobs
    .filter((job) => (job.processName ?? job.releaseName) === testSolutionConfig.processName)
    .sort((a, b) => getJobTimestampMs(b) - getJobTimestampMs(a))
    .slice(0, 25);

  return Promise.all(
    summaries.map(async (job) => {
      try {
        return ((await jobs.getById(job.key, testSolutionConfig.folderId)) as unknown as JobShape) ?? job;
      } catch {
        return job;
      }
    }),
  ).then((items) => items.sort((a, b) => getJobTimestampMs(b) - getJobTimestampMs(a)));
}

function getChildProgress(parentJob: JobShape, allJobs: JobShape[]): StepProgress {
  const childJobs = getChildJobs(parentJob, allJobs);
  const duJobs = childJobs.filter((job) => getProcessName(job).startsWith('IDP_'));
  const agentJobs = childJobs.filter((job) => getProcessName(job) === 'Contract Validation Agent');
  const exportJobs = childJobs.filter((job) => getProcessName(job) === 'Export Contract Findings');
  const agentStarted = agentJobs.length > 0;
  const exportStarted = exportJobs.length > 0;

  return {
    duStarted: duJobs.length > 0,
    duComplete: agentStarted || exportStarted || (duJobs.length > 0 && duJobs.every((job) => isCompleted(job.state))),
    duFailed: duJobs.some((job) => isFaulted(job.state)),
    consolidateStarted: false,
    consolidateComplete: agentStarted || exportStarted,
    consolidateFailed: false,
    agentStarted,
    agentComplete: exportStarted || agentJobs.some((job) => isCompleted(job.state)),
    agentFailed: agentJobs.some((job) => isFaulted(job.state)),
    exportStarted,
    exportComplete: exportJobs.some((job) => isCompleted(job.state)),
    exportFailed: exportJobs.some((job) => isFaulted(job.state)),
  };
}

function buildProgressFromHistory(history: ProcessInstanceExecutionHistoryResponse[]): StepProgress {
  const duSpans = history.filter((span) => isDuSpan(span.name));
  const consolidateSpans = history.filter((span) => isConsolidateSpan(span.name));
  const agentSpans = history.filter((span) => isAgentSpan(span.name));
  const exportSpans = history.filter((span) => isExportSpan(span.name));
  const consolidateStarted = consolidateSpans.length > 0;
  const agentStarted = agentSpans.length > 0;
  const exportStarted = exportSpans.length > 0;

  return {
    duStarted: duSpans.length > 0,
    duComplete: consolidateStarted || agentStarted || exportStarted || (duSpans.length > 0 && duSpans.every(isSpanComplete)),
    duFailed: false,
    consolidateStarted,
    consolidateComplete: agentStarted || exportStarted || consolidateSpans.some(isSpanComplete),
    consolidateFailed: false,
    agentStarted,
    agentComplete: exportStarted || agentSpans.some(isSpanComplete),
    agentFailed: false,
    exportStarted,
    exportComplete: exportSpans.some(isSpanComplete),
    exportFailed: false,
  };
}

function buildProgressFromElementExecutions(elements: ElementExecutionShape[]): StepProgress {
  const duElements = elements.filter((element) => isDuSpan(getElementExecutionName(element)));
  const consolidateElements = elements.filter((element) => isConsolidateSpan(getElementExecutionName(element)));
  const agentElements = elements.filter((element) => isAgentSpan(getElementExecutionName(element)));
  const exportElements = elements.filter((element) => isExportSpan(getElementExecutionName(element)));
  const consolidateStarted = consolidateElements.some(isElementStarted);
  const agentStarted = agentElements.some(isElementStarted);
  const exportStarted = exportElements.some(isElementStarted);

  return {
    duStarted: duElements.some(isElementStarted),
    duComplete: consolidateStarted || agentStarted || exportStarted || (duElements.length > 0 && duElements.every(isElementComplete)),
    duFailed: duElements.some(isElementFailed),
    consolidateStarted,
    consolidateComplete: agentStarted || exportStarted || consolidateElements.some(isElementComplete),
    consolidateFailed: consolidateElements.some(isElementFailed),
    agentStarted,
    agentComplete: exportStarted || agentElements.some(isElementComplete),
    agentFailed: agentElements.some(isElementFailed),
    exportStarted,
    exportComplete: exportElements.some(isElementComplete),
    exportFailed: exportElements.some(isElementFailed),
  };
}

function mergeProgress(...items: Array<StepProgress | undefined>) {
  const progressItems = items.filter(Boolean) as StepProgress[];
  if (!progressItems.length) return undefined;

  return progressItems.reduce<StepProgress>(
    (merged, progress) => ({
      duStarted: merged.duStarted || progress.duStarted,
      duComplete: merged.duComplete || progress.duComplete,
      duFailed: merged.duFailed || progress.duFailed,
      consolidateStarted: merged.consolidateStarted || progress.consolidateStarted,
      consolidateComplete: merged.consolidateComplete || progress.consolidateComplete,
      consolidateFailed: merged.consolidateFailed || progress.consolidateFailed,
      agentStarted: merged.agentStarted || progress.agentStarted,
      agentComplete: merged.agentComplete || progress.agentComplete,
      agentFailed: merged.agentFailed || progress.agentFailed,
      exportStarted: merged.exportStarted || progress.exportStarted,
      exportComplete: merged.exportComplete || progress.exportComplete,
      exportFailed: merged.exportFailed || progress.exportFailed,
    }),
    emptyProgress(),
  );
}

function emptyProgress(): StepProgress {
  return {
    duStarted: false,
    duComplete: false,
    duFailed: false,
    consolidateStarted: false,
    consolidateComplete: false,
    consolidateFailed: false,
    agentStarted: false,
    agentComplete: false,
    agentFailed: false,
    exportStarted: false,
    exportComplete: false,
    exportFailed: false,
  };
}

function isSpanComplete(span: ProcessInstanceExecutionHistoryResponse) {
  return Boolean(span.endTime);
}

function getElementExecutionName(element: ElementExecutionShape) {
  return element.name ?? element.elementName ?? element.displayName ?? element.source ?? element.elementId ?? '';
}

function isElementStarted(element: ElementExecutionShape) {
  const runs = getElementRuns(element);
  return Boolean(element.startedTime || element.status || element.state || runs.length);
}

function isElementComplete(element: ElementExecutionShape) {
  const status = normalizeStatus(element.status ?? element.state);
  if (isCompleted(status)) return true;
  if (element.endTime || element.completedTime) return true;
  const runs = getElementRuns(element);
  return runs.length > 0 && runs.every((run) => isCompleted(normalizeStatus(run.status ?? run.state)) || Boolean(run.endTime || run.completedTime));
}

function isElementFailed(element: ElementExecutionShape) {
  const status = normalizeStatus(element.status ?? element.state);
  if (isFaulted(status)) return true;
  return getElementRuns(element).some((run) => isFaulted(normalizeStatus(run.status ?? run.state)));
}

function getElementRuns(element: ElementExecutionShape) {
  return element.elementRuns ?? element.runs ?? [];
}

function isDuSpan(name: string) {
  const normalized = normalizeStepName(name);
  return normalized.includes('idp') || normalized.includes('documentunderstanding');
}

function isConsolidateSpan(name: string) {
  const normalized = normalizeStepName(name);
  return normalized.includes('consolidate') || normalized.includes('combineduoutput') || normalized.includes('prepareagentinput');
}

function isAgentSpan(name: string) {
  const normalized = normalizeStepName(name);
  return normalized.includes('contractvalidationagent');
}

function isExportSpan(name: string) {
  const normalized = normalizeStepName(name);
  return normalized.includes('exportfindings') || normalized.includes('exportcontractfindings') || normalized.includes('excelexport');
}

function normalizeStepName(name: string) {
  return name.toLowerCase().replace(/[^a-z0-9]/g, '');
}

function getChildJobs(parentJob: JobShape, allJobs: JobShape[]) {
  const parentStart = getJobTimestampMs(parentJob);
  const nextParentStart = getNextParentStart(parentJob, allJobs);
  const parentEnd = getJobEndTimestampMs(parentJob);
  const windowEnd = Math.min(
    ...[nextParentStart, parentEnd ? parentEnd + 2 * 60 * 1000 : 0, Date.now()].filter(Boolean),
  );
  const childProcessNames = new Set(['IDP_DD2579', 'IDP_SF1449', 'IDP_SAAD', 'IDP_SAAD_v2', 'IDP_DF_TR', 'Contract Validation Agent', 'Export Contract Findings']);

  return allJobs.filter((job) => {
    if (job.key === parentJob.key) return false;
    if (!childProcessNames.has(getProcessName(job))) return false;
    if (job.parentJobKey && job.parentJobKey === parentJob.key) return true;

    const started = getJobTimestampMs(job);
    return Boolean(parentStart && started >= parentStart && started <= windowEnd);
  });
}

function getNextParentStart(parentJob: JobShape, allJobs: JobShape[]) {
  const parentStart = getJobTimestampMs(parentJob);
  if (!parentStart) return 0;
  return allJobs
    .filter((job) => getProcessName(job) === testSolutionConfig.processName)
    .map(getJobTimestampMs)
    .filter((started) => started > parentStart)
    .sort((a, b) => a - b)[0] ?? 0;
}

function getProcessName(job: JobShape) {
  return job.processName ?? job.releaseName ?? '';
}

function hasVariables(variables: VariableResponse) {
  return Boolean((variables.globalVariables?.length ?? 0) || (variables.elements?.length ?? 0));
}

function outputToVariables(output: Record<string, unknown> | null): VariableResponse {
  if (!output) return {};
  return {
    globalVariables: Object.entries(output).map(([name, value]) => ({ name, value })),
  };
}

function inputArgumentsToVariables(inputArguments: string | null | undefined): VariableResponse {
  if (!inputArguments?.trim()) return {};
  try {
    const parsed = JSON.parse(inputArguments);
    if (!isRecord(parsed)) return {};
    return {
      globalVariables: Object.entries(parsed).map(([name, value]) => ({ name, value })),
    };
  } catch {
    return {};
  }
}

function getJobTimestamp(job: JobShape) {
  return job.startTime ?? job.startedTime ?? job.createdTime ?? job.creationTime ?? undefined;
}

function getJobTimestampMs(job: JobShape) {
  const value = getJobTimestamp(job);
  if (!value) return 0;
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? 0 : date.getTime();
}

function getJobEndTimestampMs(job: JobShape) {
  if (!job.endTime) return 0;
  const date = new Date(job.endTime);
  return Number.isNaN(date.getTime()) ? 0 : date.getTime();
}

function dedupeAndSortRuns(runItems: ReviewRun[]) {
  const byId = new Map<string, ReviewRun>();
  for (const runItem of runItems) {
    const key = runItem.latestRunId ?? runItem.id;
    const existing = byId.get(key);
    if (!existing || runHasMoreData(runItem, existing)) {
      byId.set(key, runItem);
    }
  }
  return [...byId.values()].sort((a, b) => getRunTimestampMs(b) - getRunTimestampMs(a));
}

function runHasMoreData(candidate: ReviewRun, existing: ReviewRun) {
  if (candidate.status === 'Failed' && existing.status !== 'Failed') return true;
  if (candidate.findings.length !== existing.findings.length) return candidate.findings.length > existing.findings.length;
  if (candidate.reportUrl && !existing.reportUrl) return true;
  if (candidate.status === 'Complete' && existing.status !== 'Complete') return true;
  if (candidate.documentsProcessed !== existing.documentsProcessed) return candidate.documentsProcessed > existing.documentsProcessed;
  return false;
}

function getRunTimestampMs(runItem: ReviewRun) {
  const date = new Date(runItem.createdAt);
  return Number.isNaN(date.getTime()) ? 0 : date.getTime();
}

function buildVariableMap(variables: VariableResponse) {
  const map = new Map<string, unknown>();
  for (const variable of variables.globalVariables ?? []) {
    if (variable.name) map.set(variable.name, variable.value);
  }
  for (const element of variables.elements ?? []) {
    for (const [key, value] of Object.entries(element.outputs ?? {})) map.set(key, value);
    for (const [key, value] of Object.entries(element.inputs ?? {})) {
      if (!map.has(key)) map.set(key, value);
    }
  }
  return map;
}

function readString(values: Map<string, unknown>, key: string) {
  const value = values.get(key);
  if (typeof value === 'string') return value;
  if (value == null) return '';
  return String(value);
}

function getReportRowsJson(values: Map<string, unknown>) {
  return readString(values, 'normalizedExcelReportRowsJson') || readString(values, 'excelReportRowsJson');
}

function getInputDocumentNames(values: Map<string, unknown>) {
  return ['dd2579DocumentPath', 'sf1449DocumentPath1', 'sf1449DocumentPath2', 'saadDocumentPath', 'dftrDocumentPath']
    .map((key) => readString(values, key))
    .filter(Boolean);
}

function buildRunDisplayName(inputDocuments: string[]) {
  const tokenCounts = new Map<string, number>();
  for (const documentName of inputDocuments) {
    const match = documentName.match(/\bTE\s*[-_ ]?(\d+)\b/i) ?? documentName.match(/\bTE(\d+)\b/i);
    if (!match) continue;
    const token = `TE${match[1]}`;
    tokenCounts.set(token, (tokenCounts.get(token) ?? 0) + 1);
  }

  const [bestToken] =
    [...tokenCounts.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))[0] ?? [];
  return bestToken ? `Contract Review Package ${bestToken}` : 'Contract Review Package';
}

function buildStepStatuses(values: Map<string, unknown>, state: string, progress?: StepProgress): RunStep[] {
  const hasDuOutput = Boolean(readString(values, 'dd2579ContractDataJson') || readString(values, 'sf1449ContractDataJson'));
  const hasConsolidatedJson = Boolean(readString(values, 'contractDataJson'));
  const hasAgentOutput = Boolean(readString(values, 'content') || readString(values, 'excelReportRowsJson'));
  const hasExportOutput = Boolean(readString(values, 'sharePointUploadStatus') || readString(values, 'localWorkbookPath'));
  const hasDownstreamOutput = hasConsolidatedJson || hasAgentOutput || hasExportOutput;
  const statuses: StepStatus[] = ['Complete', 'Pending', 'Pending', 'Pending', 'Pending'];

  if (progress) {
    statuses[1] = progress.duFailed ? 'Error' : progress.duComplete || progress.consolidateStarted || hasDuOutput || hasDownstreamOutput ? 'Complete' : progress.duStarted ? 'Running' : 'Pending';
    statuses[2] = progress.consolidateFailed
      ? 'Error'
      : progress.consolidateComplete || progress.agentStarted || progress.exportStarted || hasDownstreamOutput
        ? 'Complete'
        : progress.consolidateStarted
          ? 'Running'
          : statuses[1] === 'Complete'
            ? 'Running'
            : 'Pending';
    statuses[3] = progress.agentFailed
      ? 'Error'
      : progress.agentComplete || progress.exportStarted || hasAgentOutput || hasExportOutput
        ? 'Complete'
        : progress.agentStarted
          ? 'Running'
          : 'Pending';
    statuses[4] = progress.exportFailed ? 'Error' : progress.exportComplete || hasExportOutput ? 'Complete' : progress.exportStarted ? 'Running' : 'Pending';
  } else {
    statuses[1] = hasDuOutput || hasDownstreamOutput ? 'Complete' : 'Pending';
    statuses[2] = hasDownstreamOutput ? 'Complete' : statuses[1] === 'Complete' ? 'Running' : 'Pending';
    statuses[3] = hasAgentOutput || hasExportOutput ? 'Complete' : hasConsolidatedJson ? 'Running' : 'Pending';
    statuses[4] = hasExportOutput ? 'Complete' : hasAgentOutput ? 'Running' : 'Pending';
  }

  if (isFaulted(state)) {
    const failedIndex = Math.max(0, statuses.findIndex((status) => status !== 'Complete'));
    statuses[failedIndex === -1 ? statuses.length - 1 : failedIndex] = 'Error';
  } else if (isRunning(state)) {
    const runningIndex = statuses.findIndex((status) => status === 'Running' || status === 'Pending');
    if (runningIndex >= 0 && statuses[runningIndex] !== 'Error') statuses[runningIndex] = 'Running';
  } else if (isCompleted(state) && hasExportOutput) {
    for (let index = 0; index < statuses.length; index += 1) statuses[index] = 'Complete';
  }

  return defaultSteps.map((step, index) => ({ ...step, status: statuses[index] }));
}

function parseExcelRows(raw: string): Array<Record<string, unknown>> {
  if (!raw.trim()) return [];
  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) return parsed.filter(isRecord);
    if (isRecord(parsed) && Array.isArray(parsed.rows)) return parsed.rows.filter(isRecord);
  } catch {
    return [];
  }
  return [];
}

function rowsToFindings(rows: Array<Record<string, unknown>>): AgentFinding[] {
  return rows.map((row, index) => {
    const finding: AgentFinding = {
      check: readRow(row, ['Check', 'Business Rule', 'Rule', 'check']) || `Finding ${index + 1}`,
      result: normalizeResult(readRow(row, ['Result', 'Status', 'result'])),
      recommendation: readRow(row, ['Recommendation', 'Recommended Action', 'Action', 'recommendation']) || 'Review',
      fieldsReviewed: readRow(row, ['Fields Reviewed', 'Field Reviewed', 'fieldsReviewed']) || 'See workbook row',
      valuesCompared: readRow(row, ['Values Compared', 'Value Compared', 'valuesCompared']) || 'See workbook row',
      summary: readRow(row, ['Result Summary', 'Summary', 'Issue', 'resultSummary']) || 'See workbook row',
      action: readRow(row, ['Recommended Action', 'Action', 'action']) || 'Review the generated workbook.',
      notes: readRow(row, ['Data Completeness Notes', 'Completeness Notes', 'Notes', 'notes']),
      sources: readRow(row, ['Source Documents', 'Sources', 'sourceDocuments']) || 'Maestro output',
    };

    return normalizeSamFinding(finding);
  });
}

function normalizeSamFinding(finding: AgentFinding): AgentFinding {
  if (normalizeKey(finding.check) !== 'samexclusionsearchdate') return finding;

  const evidence = [finding.valuesCompared, finding.summary, finding.notes, finding.action].join(' ');
  const count = extractNumber(evidence, /computed(?:\s+business-day)?(?:\s+count|\s+conclusion)?(?:\s+is|:)?\s*(\d+)/i);
  const requiredWindow = extractNumber(evidence, /(?:required\s+window|requires\s+a|required)\s*(\d+)\s*-?\s*business/i);
  if (count == null || requiredWindow == null) return finding;

  const passes = count <= requiredWindow;
  return {
    ...finding,
    result: passes ? 'Pass' : 'Flag',
    recommendation: passes ? 'No action needed' : 'Review',
    action: passes ? 'No action needed' : 'Review required',
    notes: passes ? '' : `Computed business-day count ${count} is outside the required ${requiredWindow}-business-day window.`,
    summary: passes
      ? `The SAM exclusion search date check passes because the computed business-day count is ${count}, which is within the required ${requiredWindow}-business-day window.`
      : `The SAM exclusion search date check is flagged because the computed business-day count is ${count}, which is outside the required ${requiredWindow}-business-day window.`,
  };
}

function extractNumber(text: string, pattern: RegExp) {
  const match = text.match(pattern);
  return match ? Number.parseInt(match[1], 10) : null;
}

function readRow(row: Record<string, unknown>, keys: string[]) {
  for (const key of keys) {
    const value = row[key];
    if (value != null && String(value).trim()) return String(value);
  }
  const normalizedKeys = new Map(Object.keys(row).map((key) => [normalizeKey(key), key]));
  for (const key of keys) {
    const actual = normalizedKeys.get(normalizeKey(key));
    if (actual && row[actual] != null && String(row[actual]).trim()) return String(row[actual]);
  }
  return '';
}

function normalizeKey(key: string) {
  return key.toLowerCase().replace(/[^a-z0-9]/g, '');
}

function normalizeResult(value: string): RuleResult {
  const normalized = value.toLowerCase();
  if (normalized.includes('flag') || normalized.includes('fail')) return 'Flag';
  if (normalized.includes('pass')) return 'Pass';
  return 'Not Applicable';
}

function summarizeFindings(findings: AgentFinding[]): RuleResult {
  if (findings.some((finding) => finding.result === 'Flag')) return 'Flag';
  if (findings.some((finding) => finding.result === 'Pass')) return 'Pass';
  return 'Not Applicable';
}

function pendingFinding(finished: boolean, sharePointStatus: string): AgentFinding[] {
  if (!finished) return [];
  return [
    {
      check: 'Agent findings unavailable in app',
      result: 'Not Applicable',
      recommendation: 'Open the generated workbook',
      fieldsReviewed: 'Maestro output variables',
      valuesCompared: sharePointStatus || 'The job completed, but Excel row JSON was not available to the app.',
      summary: 'The automation completed, but the app could not parse agent finding rows from the Maestro variables.',
      action: 'Open the SharePoint Excel report and confirm the agent output.',
      notes: '',
      sources: 'Maestro process instance',
    },
  ];
}

function buildSharePointReport(status: string, localPath: string, bucketPath: string) {
  const folderMatch = status.match(/Uploaded to SharePoint folder:\s*(.*?);\s*item:\s*(.+)$/i);
  if (folderMatch) {
    const folderUrl = folderMatch[1].trim();
    const itemName = normalizeSharePointItemName(folderMatch[2].trim(), localPath, bucketPath);
    return {
      label: itemName,
      itemName,
      url: buildSharePointFileUrl(folderUrl, itemName),
    };
  }

  const itemName = localPath.split(/[\\/]/).pop() || bucketPath.split('/').pop() || '';
  return {
    label: itemName || bucketPath || 'Pending',
    itemName,
    url: '',
  };
}

function normalizeSharePointItemName(itemName: string, localPath: string, bucketPath: string) {
  if (/\.[a-z0-9]+$/i.test(itemName)) return itemName;
  const localName = localPath.split(/[\\/]/).pop() || bucketPath.split('/').pop() || '';
  if (localName && localName.startsWith(itemName) && /\.[a-z0-9]+$/i.test(localName)) return localName;
  return `${itemName}.xlsx`;
}

function buildSharePointFileUrl(folderUrl: string, itemName: string) {
  try {
    const url = new URL(folderUrl);
    const folderPath = url.searchParams.get('id');
    if (folderPath) {
      const filePath = `${folderPath.replace(/\/$/, '')}/${itemName}`;
      url.searchParams.set('id', filePath);
      url.searchParams.set('parent', folderPath);
      url.searchParams.set('web', '1');
      return url.toString();
    }
  } catch {
    // Fall through to direct path construction below.
  }

  return `${folderUrl.replace(/\/$/, '')}/${encodeURIComponent(itemName)}`;
}

export function buildMaestroInstanceUrl(instanceId: string) {
  return `https://cloud.uipath.com/${testSolutionConfig.orgName}/${testSolutionConfig.tenantName}/maestro_/processes/${testSolutionConfig.processKey.toLowerCase()}/instances/${instanceId}?folderKey=${testSolutionConfig.folderKey}`;
}

export function buildMaestroProcessUrl() {
  return `https://cloud.uipath.com/${testSolutionConfig.orgName}/${testSolutionConfig.tenantName}/maestro_/processes/${testSolutionConfig.processKey.toLowerCase()}?folderKey=${testSolutionConfig.folderKey}`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value));
}

function isCompleted(state: string) {
  return ['successful', 'completed', 'complete', 'succeeded'].includes(normalizeStatus(state));
}

function isRunning(state: string) {
  return ['running', 'pending', 'resuming', 'retrying'].includes(normalizeStatus(state));
}

function isFaulted(state: string) {
  return ['faulted', 'failed', 'stopped', 'cancelled', 'canceled'].includes(normalizeStatus(state));
}

function normalizeStatus(state: string | undefined) {
  return (state ?? '').toLowerCase().replace(/[^a-z]/g, '');
}

function formatDate(value: string | undefined) {
  if (!value) return 'Unknown time';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString([], { month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' });
}
