export type DocumentSlotId = 'dd2579' | 'sf1449Award' | 'sf1449Solicitation' | 'saad' | 'dftr';

export type DocumentSlot = {
  id: DocumentSlotId;
  label: string;
  helper: string;
  required: boolean;
  fileName: string;
};

export type RunStatus = 'Draft' | 'Ready' | 'Running' | 'Complete' | 'Failed';

export type StepStatus = 'Pending' | 'Running' | 'Complete' | 'Error';

export type RunStep = {
  label: string;
  detail: string;
  status: StepStatus;
};

export type RuleResult = 'Pass' | 'Flag' | 'Not Applicable';

export type AgentFinding = {
  check: string;
  result: RuleResult;
  recommendation: string;
  fieldsReviewed: string;
  valuesCompared: string;
  summary: string;
  action: string;
  notes: string;
  sources: string;
};

export type ReviewRun = {
  id: string;
  displayName: string;
  status: RunStatus;
  createdAt: string;
  packageName: string;
  analyst: string;
  overallStatus: RuleResult;
  documentsProcessed: number;
  reportPath: string;
  sharePointItem: string;
  reportUrl?: string;
  instanceId?: string;
  latestRunId?: string;
  maestroUrl: string;
  steps: RunStep[];
  findings: AgentFinding[];
};
