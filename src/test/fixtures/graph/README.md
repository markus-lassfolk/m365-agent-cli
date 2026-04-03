# Microsoft Graph fixtures (tests & docs)

Official request/response schemas are defined by Microsoft (OpenAPI / metadata). This folder stores **minimal examples** we assert on in unit tests so we catch regressions beyond happy paths.

| Resource | Canonical docs |
| --- | --- |
| Error JSON (`error.code`, `error.message`, `innerError`) | [Microsoft Graph error responses](https://learn.microsoft.com/en-us/graph/errors) |
| OAuth2 token endpoint errors (`error`, `error_description`, optional `error_codes`) | [Azure AD OAuth errors](https://learn.microsoft.com/en-us/azure/active-directory/develop/reference-aadsts-error-codes) — e.g. `invalid_grant` when refresh token does not match client |
| `sendMail` | [sendMail](https://learn.microsoft.com/en-us/graph/api/user-sendmail) — body is `{ message, saveToSentItems }` |
| OData | [OData protocol](https://learn.microsoft.com/en-us/graph/use-the-api) |

**Do not** treat these JSON files as exhaustive schemas; extend them when adding tests for new failure modes (throttling, `innerError.date`, etc.).
