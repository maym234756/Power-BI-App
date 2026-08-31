# Power BI App Project

This is a React + Vite starter app that signs in with Microsoft Entra ID, reads your Power BI workspaces and reports through the Power BI REST API, and embeds the selected report with `powerbi-client`.

## Run the app

```powershell
npm install
npm run dev
```

Open the local URL that Vite prints. The redirect URI in Azure must exactly match that origin, including whether it uses `localhost` or `127.0.0.1`.

## Connect it to Power BI

1. Go to the Microsoft Entra admin center and open **App registrations**.
2. Create a new registration for this local app.
3. In **Authentication**, add a **Single-page application** platform with the redirect URI shown in the app. For this workspace, use `http://localhost:5173`.
4. In **API permissions**, add delegated permissions for **Power BI Service**:
   - `Workspace.Read.All`
   - `Dashboard.Read.All`
   - `Report.Read.All`
   - `Dataset.Read.All`
5. Grant admin consent if your organization requires it.
6. Copy the **Application (client) ID** and paste it into the app, or create `.env.local` from `.env.example`:

```powershell
Copy-Item .env.example .env.local
```

Then edit `.env.local`:

```env
VITE_POWER_BI_CLIENT_ID=your-client-id-here
VITE_POWER_BI_TENANT_ID=common
```

Use your tenant ID instead of `common` if your app registration is single-tenant.

When you click **Sign in to Power BI**, the app redirects the current browser tab to Microsoft sign-in and then returns to `http://localhost:5173`. This avoids popup windows inside VS Code.

## Notes

- This starter uses the Power BI "user owns data" pattern. The person signing in must already have access to the workspace and report.
- The app is intentionally locked to the `Miles May` workspace for the `PBC & OMG Parts and Service` app work. It does not offer other workspaces in the UI.
- The visual shell is intentionally scoped to Parts and Service only. The Parts and Service dropdowns are wired to the `Parts Inventory` and `PY&OMG Service` dashboards when those are available in `Miles May`.
- No client secret belongs in this app. Browser apps use public client sign-in.
- Your organization may require a Power BI Pro or Premium Per User license, or backed capacity, depending on the content you embed.