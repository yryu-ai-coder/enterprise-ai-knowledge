# Security Notes

## Principles

- Keep AI secrets out of the SPFx web part.
- Use Function App as the trust boundary.
- Prefer Managed Identity and Key Vault for secrets.
- Validate request origin and tenant/site context.
- Do not log sensitive prompts/responses by default.
- Add permission trimming before production grounding.

## POC vs Production

For POC, CORS and auth may be simplified. For production, require Entra ID-authenticated calls and validate user/site authorization server-side.
