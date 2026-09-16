# SaaS email connections

Email accounts are workspace data, not environment variables. One connection is
identified by `user_id + workspace_id + provider`. Gmail refresh tokens and App
Passwords are encrypted with AES-256-GCM before they are written to PostgreSQL.

## Platform setup

Configure these values in the deployment environment:

- `DATABASE_URL`: PostgreSQL connection string.
- `AUTH_SESSION_SECRET`: verifies the signed `hrdoc_session` HttpOnly cookie.
- `EMAIL_CREDENTIAL_ENCRYPTION_KEY`: platform encryption root key. Rotate it with
  a planned credential migration; changing it immediately makes existing
  credentials unreadable.
- `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`, `GOOGLE_OAUTH_REDIRECT_URI`: Google
  OAuth application used by all tenants.

Never add a user's Gmail address, refresh token, or App Password to Vercel
environment variables. Those values belong to the encrypted `email_connections`
record.

## Trusted identity boundary

The Vercel functions require a signed `hrdoc_session` cookie containing:

```json
{
  "sub": "stable-user-id",
  "workspaceId": "stable-workspace-id",
  "email": "user@example.com",
  "role": "Admin",
  "exp": 1789420800
}
```

The JSON payload is base64url encoded and signed with HMAC-SHA256 using
`AUTH_SESSION_SECRET`. Replace the current browser-only demo login with the chosen
identity provider before enabling email connections in production. The API does
not accept `user_id` or `workspace_id` from request bodies.

## Deployment sequence

1. Provision PostgreSQL and apply `docs/email-connections.sql`.
2. Configure trusted server authentication and the HttpOnly session cookie.
3. Create a Google OAuth application and add the production callback URL.
4. Add the platform variables in Vercel and redeploy.
5. Open Settings > Email connections and connect a test Gmail account.
6. Send a test email, revoke the connection, then reconnect to verify the full
   credential lifecycle.

Until steps 1-4 are complete, the UI deliberately shows **Cloud setup required**
and disables connection actions. It never falls back to localStorage or a local
`.env` file.
