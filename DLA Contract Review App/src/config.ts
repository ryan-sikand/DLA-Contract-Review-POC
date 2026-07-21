import type { DocumentSlot } from './types';

export const uipathConfig = {
  clientId: import.meta.env.VITE_UIPATH_CLIENT_ID ?? '',
  orgName: import.meta.env.VITE_UIPATH_ORG_NAME ?? 'uipathlabs',
  tenantName: import.meta.env.VITE_UIPATH_TENANT_NAME ?? 'Playground',
  baseUrl: import.meta.env.VITE_UIPATH_BASE_URL ?? 'https://api.uipath.com',
  redirectUri: import.meta.env.VITE_UIPATH_REDIRECT_URI ?? window.location.origin,
  scope:
    import.meta.env.VITE_UIPATH_SCOPE ??
    'OR.Administration OR.Execution.Read OR.Jobs OR.Jobs.Read OR.Jobs.Write OR.Tasks OR.Tasks.Read OR.Tasks.Write PIMS DataFabric.Schema.Read DataFabric.Data.Read DataFabric.Data.Write',
};

export const testSolutionConfig = {
  label: 'Main solution folder',
  orgName: 'uipathlabs',
  tenantName: 'Playground',
  folderPath: 'AMER Presales/Public Sector/DLA Contract Quality Review POC',
  folderKey: '90d1ef22-feca-4f37-aac7-d5b8a32b8d3a',
  folderId: 2164189,
  processName: 'Agentic Process',
  processKey: '24BF5442-7C51-45F0-B362-D27A067510F2',
  packageProcessKey: 'DLA.Contract.Quality.Review.agentic.Agentic.Process',
  processVersion: '1.0.65',
  bucketName: 'DLA_Contract_Docs',
  bucketFolderPath: 'AMER Presales/Public Sector/DLA Contract Quality Review POC',
  studioUrl:
    'https://cloud.uipath.com/uipathlabs/studio_/designer/f5308d15-f2f8-4c25-9f89-5dd0a7cf8ead?solutionId=33e26b6e-c6ea-4b82-3347-08dec0aa4471',
};

export function isUiPathConfigured() {
  return Boolean(uipathConfig.clientId && uipathConfig.orgName && uipathConfig.tenantName && uipathConfig.baseUrl);
}

export function buildMaestroInputs(documents: DocumentSlot[]) {
  const byId = Object.fromEntries(documents.map((doc) => [doc.id, normalizeDocumentName(doc.fileName)]));

  return {
    dd2579DocumentPath: byId.dd2579,
    sf1449DocumentPath1: byId.sf1449Award,
    sf1449DocumentPath2: byId.sf1449Solicitation,
    saadDocumentPath: byId.saad,
    dftrDocumentPath: byId.dftr,
  };
}

function normalizeDocumentName(fileName: string) {
  return fileName.trim().replace(/\.pdf$/i, '');
}
