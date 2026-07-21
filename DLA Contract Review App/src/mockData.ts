import type { AgentFinding, DocumentSlot, ReviewRun, RunStep } from './types';

export const initialDocumentSlots: DocumentSlot[] = [
  {
    id: 'dd2579',
    label: 'DD2579',
    helper: 'Required request document. One file only.',
    required: true,
    fileName: 'DD2579-TE3.pdf',
  },
  {
    id: 'sf1449Award',
    label: 'SF1449 Award',
    helper: 'Required award document.',
    required: true,
    fileName: 'SF1449-TE3-A.pdf',
  },
  {
    id: 'sf1449Solicitation',
    label: 'SF1449 Solicitation',
    helper: 'Required solicitation document.',
    required: true,
    fileName: 'SF1449-TE3-S.pdf',
  },
  {
    id: 'saad',
    label: 'SAAD',
    helper: 'Optional. Include when available.',
    required: false,
    fileName: 'SAAD-TE3.pdf',
  },
  {
    id: 'dftr',
    label: 'D&F / TR',
    helper: 'Optional. Include when applicable.',
    required: false,
    fileName: 'DF-TE3.pdf',
  },
];

export const defaultSteps: RunStep[] = [
  {
    label: 'Document Intake',
    detail: 'Files staged for the Maestro process',
    status: 'Complete',
  },
  {
    label: 'Document Understanding',
    detail: 'DD2579, SF1449, SAAD, and D&F extractors completed',
    status: 'Complete',
  },
  {
    label: 'Consolidate JSON',
    detail: 'Combined DU output prepared for the agent',
    status: 'Complete',
  },
  {
    label: 'Contract Validation Agent',
    detail: 'Business rules evaluated with policy grounding',
    status: 'Complete',
  },
  {
    label: 'Excel Export',
    detail: 'Workbook saved to storage bucket and SharePoint',
    status: 'Complete',
  },
];

export const mockFindings: AgentFinding[] = [
  {
    check: 'NAICS / PSC / Size Standard Match',
    result: 'Pass',
    recommendation: 'No action needed',
    fieldsReviewed: 'NAICS, PSC, size standard',
    valuesCompared:
      'DD2579 DD2579-TE3: NAICS 488190; PSC J011; size standard $34M. SF1449 award SF1449-TE3-A: NAICS 488190; PSC J011; size standard $34M. SF1449 solicitation SF1449-TE3-S: NAICS 488190; PSC J011; size standard $34M.',
    summary: 'The required values align across the DD2579 and both SF1449 records.',
    action: 'No action needed',
    notes: 'SF1449 PSC field was blank in the primary extraction but confirmed from continuation-page evidence.',
    sources: 'DD2579-TE3, SF1449-TE3-A, SF1449-TE3-S',
  },
  {
    check: 'NAICS / SBA Size Standard',
    result: 'Flag',
    recommendation: 'Review',
    fieldsReviewed: 'NAICS, size standard',
    valuesCompared: 'DD2579 DD2579-TE3: NAICS 488190; size standard $34M. SBA reference: NAICS 488190 = $40.0M.',
    summary: 'The DD2579 size standard does not match the SBA reference returned for NAICS 488190.',
    action: 'Confirm the NAICS and size standard against the SBA Table of Size Standards.',
    notes: 'Reference evidence was available and indicates a mismatch rather than missing data.',
    sources: 'DD2579-TE3, SBA Table of Size Standards',
  },
  {
    check: 'Semantic Alignment',
    result: 'Flag',
    recommendation: 'Review',
    fieldsReviewed: 'Item/service description, PSC, NAICS',
    valuesCompared:
      'DD2579 description: Aircraft component inspection and repair services, turboprop engines. PSC J011. NAICS 488190. NAICS reference: Other Support Activities for Air Transportation.',
    summary: 'The service description appears aligned to aircraft support activity, but PSC reference detail was incomplete.',
    action: 'Review PSC J011 and NAICS 488190 against the PSC and NAICS manuals.',
    notes: 'The policy response did not return a clean PSC Manual definition for J011.',
    sources: 'DD2579-TE3, NAICS Manual, PSC Manual',
  },
  {
    check: 'SAM Exclusion Search Date',
    result: 'Pass',
    recommendation: 'No action needed',
    fieldsReviewed: 'SAM exclusion search date, award date, issued-by code',
    valuesCompared:
      'SAAD SAAD-TE3: SAM exclusion search date 04/16/2026. SF1449 award SF1449-TE3-A: award date 04/24/2026. Issued-by code SPRBL1. Required window: within 7 business days prior to award. Computed conclusion: 6 business days before award.',
    summary: 'The SAM exclusion search date falls within the required 7-business-day window.',
    action: 'No action needed',
    notes: '',
    sources: 'SAAD-TE3, SF1449-TE3-A',
  },
  {
    check: 'D&F Requirement',
    result: 'Flag',
    recommendation: 'Review',
    fieldsReviewed: 'CLIN type, award date, D&F signature date',
    valuesCompared:
      'SF1449 award continuation evidence: Time-and-Materials contract noted with labor categories priced by hour. Award date: 04/24/2026. D&F DF-TE3 present with signature date 04/16/2026.',
    summary: 'A D&F is present and signed before award, but T&M CLIN evidence came from continuation text rather than structured extraction fields.',
    action: 'Inspect the SF1449 award schedule or continuation pages to confirm the T&M CLINs.',
    notes: 'The DU payload did not include explicit CLIN detail fields.',
    sources: 'SF1449-TE3-A, DF-TE3',
  },
];

export const completedRun: ReviewRun = {
  id: 'DLA-CQR-20260615162246',
  displayName: 'Contract Review Package TE3',
  status: 'Complete',
  createdAt: 'Jun 15, 2026 12:25 PM',
  packageName: 'DLA Contract Quality Review POC',
  analyst: 'Ryan Sikand',
  overallStatus: 'Flag',
  documentsProcessed: 5,
  reportPath: 'reports/DLA_Contract_Quality_Findings_20260615162524.xlsx',
  sharePointItem: 'DLA_Contract_Quality_Findings_20260615162524',
  maestroUrl: 'https://cloud.uipath.com/uipathlabs/studio_/designer/f5308d15-f2f8-4c25-9f89-5dd0a7cf8ead?solutionId=33e26b6e-c6ea-4b82-3347-08dec0aa4471',
  steps: defaultSteps,
  findings: mockFindings,
};

export const runHistory: ReviewRun[] = [
  completedRun,
  {
    ...completedRun,
    id: 'DLA-CQR-20260615162528',
    createdAt: 'Jun 15, 2026 12:25 PM',
    documentsProcessed: 5,
    reportPath: 'reports/DLA_Contract_Quality_Findings_20260615162528.xlsx',
    sharePointItem: 'DLA_Contract_Quality_Findings_20260615162528',
    overallStatus: 'Flag',
    findings: mockFindings.map((finding) =>
      finding.check === 'SAM Exclusion Search Date'
        ? {
            ...finding,
            result: 'Flag',
            recommendation: 'Review',
            summary: 'The SAM exclusion search date was outside the required business-day window.',
          }
        : finding,
    ),
  },
  {
    ...completedRun,
    id: 'DLA-CQR-20260615162454',
    createdAt: 'Jun 15, 2026 12:24 PM',
    documentsProcessed: 4,
    reportPath: 'reports/DLA_Contract_Quality_Findings_20260615162454.xlsx',
    sharePointItem: 'DLA_Contract_Quality_Findings_20260615162454',
    overallStatus: 'Pass',
    findings: mockFindings.map((finding) => ({
      ...finding,
      result: finding.check === 'D&F Requirement' ? 'Not Applicable' : 'Pass',
      recommendation: finding.check === 'D&F Requirement' ? 'Not applicable' : 'No action needed',
    })),
  },
];
