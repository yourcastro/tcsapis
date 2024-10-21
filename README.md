As discussed on Friday, We are calling the APIs from CreditRiskPortalAPI to CreditRiskPortalGenericAPI (for SharePoint & Import Scorecard Data), since both APIs are hosted on the same server, we should use the server’s IP address. However, I have tried but it’s not working.

CreditRiskPortalAPI: https://dev-intranet-std.ca.sunlife/CreditRiskPortalAPI/
CreditRiskPortalGenericAPI: https://dev-intranet-std.ca.sunlife/CreditRiskPortalGenericAPI/ (This URL is working in our local but when we deployed to server and try to access it’s not working)

Server IP URL: https://159.208.208.142/CreditRiskPortalGenericAPI/api/sp/CreateNewPDScorecardByTemplate
Exception:  AuthenticationException: The remote certificate is invalid according to the validation procedure: RemoteCertificateNameMismatch, RemoteCertificateChainErrors

Could you please advise on how to resolve this issue?

Thank you!
