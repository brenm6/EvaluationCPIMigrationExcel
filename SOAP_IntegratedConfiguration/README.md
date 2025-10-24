# SAP PI/PO IntegratedConfigurationInService (WSDL)

This folder contains ready-to-use SOAP requests for the administrative PI/PO web service
`IntegratedConfigurationInService / IntegratedConfigurationInImplBean` (Design-time API).
Use it to read/query/create/change/delete Integrated Configurations (ICOs) in the Integration Directory.

- WSDL file you provided: `IntegratedConfigurationInImplBean_WSDL.xml`
- SOAP version: 1.1 (document/literal)
- Namespace: `http://sap.com/xi/BASIS`
- SOAPAction: `http://sap.com/xi/WebService/soap1.1`

## Endpoint
Replace the host/port with your system when calling:

- Example: `http://pi.server:50000/IntegratedConfigurationInService/IntegratedConfigurationInImplBean`
- The WSDL may show an internal host (e.g., `sapbpxapp01:52100`) – always use your actual host/port.

## Authentication
- HTTP Basic Auth against AS Java (user with Integration Directory rights, e.g. `SAP_XI_DIRECTORY_*`).
- Headers:
  - `Content-Type: text/xml; charset=utf-8`
  - `SOAPAction: "http://sap.com/xi/WebService/soap1.1"`

## Files
- `request-read.xml` – Read one or more ICOs by header (sender/interface[/receiver]).
- `request-query.xml` – Query ICOs using description/admin filters.
- `request-open-for-edit.xml` – Open ICO(s) for edit and obtain a ChangeListID.
- `request-change-skeleton.xml` – Skeleton for Change/Create (fill the structure as needed).

## Quick test (PowerShell)

```powershell
$uri = "http://pi.server:50000/IntegratedConfigurationInService/IntegratedConfigurationInImplBean"
$body = Get-Content -Raw -Path .\request-read.xml

Invoke-WebRequest -Uri $uri -Method POST -Headers @{ "SOAPAction" = "http://sap.com/xi/WebService/soap1.1" } `
  -ContentType "text/xml; charset=utf-8" `
  -Body $body `
  -Authentication Basic -Credential (Get-Credential)
```

> Tip: You can import the WSDL into SoapUI to generate a project automatically.
