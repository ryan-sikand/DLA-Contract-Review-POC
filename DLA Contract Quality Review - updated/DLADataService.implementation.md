# DLA Data Service Integration Notes

This solution uses `DLADataService` as the persistent contract quality review record store.

## Required Authenticated Setup

Run these after `uip login --interactive` and tenant selection:

```powershell
uip df entities list --native-only --output json
uip df entities create DLADataService --file DLADataService.schema.json --output json
```

If the entity already exists, compare it with `DLADataService.schema.json` and add only missing fields with `uip df entities update`.

## Workflow Change Points

- `IDP_DD2579/Main.xaml`: write after `Extract Document Data (DD2579)` and after validation if validation is re-enabled.
- `IDP_SF1449/Main.xaml`: write after `Extract Document Data (SF1449)` and after validation if validation is re-enabled.
- `IDP_SAAD/Main.xaml`: write after `Validate SAAD extraction`.
- `IDP_DF_TR/Main.xaml`: write after `Extract Document Data (DF TR)` and after validation if validation is re-enabled.

Each write should use `DocumentRecordKey = ContractBatchId + ":" + DocumentType + ":" + SourceFilePath`.

## Agent Input Contract

The Contract Validation Agent now accepts:

```json
{
  "contractDataJson": "<JSON assembled from DLADataService records>"
}
```

The process orchestration should query `DLADataService` by `ContractBatchId`, assemble one JSON payload containing all matching document records, call the agent, and write the agent response to `AgentOutputJson`.

## Binding Notes

Do not hand-edit `resources/solution_folder`. After the Data Service entity exists and the workflow activities/context are added in Studio Web or by CLI enrichment, run:

```powershell
uip solution resource refresh "DLA Contract Quality Review.uipx" --output json
```

Then pack the solution with the older CLI surface:

```powershell
uip solution pack . ".\DLA Contract Quality Review.updated.zip" --output json
```
