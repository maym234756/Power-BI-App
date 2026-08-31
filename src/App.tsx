import { useEffect, useRef, useState } from 'react'
import { PublicClientApplication } from '@azure/msal-browser'
import type { AccountInfo } from '@azure/msal-browser'
import { factories, models, service } from 'powerbi-client'
import './App.css'

const powerBiScopes = [
  'https://analysis.windows.net/powerbi/api/Workspace.Read.All',
  'https://analysis.windows.net/powerbi/api/Dashboard.Read.All',
  'https://analysis.windows.net/powerbi/api/Report.Read.All',
  'https://analysis.windows.net/powerbi/api/Dataset.Read.All',
]

const clientIdStorageKey = 'powerbi-app-client-id'
const tenantIdStorageKey = 'powerbi-app-tenant-id'
const targetWorkspaceName = 'Miles May'
const targetAppName = 'PBC & OMG Parts and Service'

type Workspace = {
  id: string
  name: string
  type: 'personal' | 'group'
}

type PowerBiReport = {
  id: string
  name: string
  embedUrl: string
  webUrl: string
  datasetId?: string
}

type PowerBiDashboard = {
  id: string
  displayName: string
  embedUrl: string
  isReadOnly?: boolean
}

type PowerBiContent = {
  key: string
  kind: 'dashboard' | 'report'
  id: string
  name: string
  embedUrl: string
  webUrl?: string
}

type PowerBiWorkspacesResponse = {
  value: Array<{ id: string; name: string }>
}

type PowerBiReportsResponse = {
  value: PowerBiReport[]
}

type PowerBiDashboardsResponse = {
  value: PowerBiDashboard[]
}

type ConnectionConfig = {
  clientId: string
  tenantId: string
}

type DepartmentName = 'Parts' | 'Service' | 'Technicians'

type DepartmentItem = {
  label: string
  targetName?: string
  kind?: PowerBiContent['kind']
}

type DepartmentTab = {
  name: DepartmentName
  summary: string
  items: DepartmentItem[]
}

const departmentTabs: DepartmentTab[] = [
  {
    name: 'Parts',
    summary: 'Inventory, backorders, sales, turns, and margin checkpoints.',
    items: [
      { label: 'Parts Inventory', targetName: 'Parts Inventory', kind: 'dashboard' },
      { label: 'Parts sales' },
      { label: 'Backorders' },
      { label: 'Gross margin' },
    ],
  },
  {
    name: 'Service',
    summary: 'Repair order flow, revenue, labor recovery, and service aging.',
    items: [
      { label: 'PY&OMG Service', targetName: 'PY&OMG Service', kind: 'dashboard' },
      { label: 'Repair orders' },
      { label: 'Revenue trends' },
      { label: 'RO aging' },
    ],
  },
  {
    name: 'Technicians',
    summary: 'Productivity, efficiency, billed hours, and technician scorecards.',
    items: [
      { label: 'Tech scorecards' },
      { label: 'Efficiency' },
      { label: 'Billed hours' },
      { label: 'Clocked time' },
    ],
  },
]

function getInitialConfig(): ConnectionConfig {
  return {
    clientId:
      import.meta.env.VITE_POWER_BI_CLIENT_ID ??
      localStorage.getItem(clientIdStorageKey) ??
      '',
    tenantId:
      import.meta.env.VITE_POWER_BI_TENANT_ID ??
      localStorage.getItem(tenantIdStorageKey) ??
      'common',
  }
}

function createMsalApp(config: ConnectionConfig) {
  return new PublicClientApplication({
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId || 'common'}`,
      redirectUri: window.location.origin,
    },
    cache: {
      cacheLocation: 'localStorage',
    },
  })
}

async function powerBiFetch<T>(path: string, accessToken: string): Promise<T> {
  const response = await fetch(`https://api.powerbi.com/v1.0/myorg${path}`, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  })

  if (!response.ok) {
    const message = await response.text()
    throw new Error(message || `Power BI request failed: ${response.status}`)
  }

  return response.json() as Promise<T>
}

function App() {
  const [config, setConfig] = useState<ConnectionConfig>(getInitialConfig)
  const [draftClientId, setDraftClientId] = useState(config.clientId)
  const [draftTenantId, setDraftTenantId] = useState(config.tenantId)
  const [msalApp, setMsalApp] = useState<PublicClientApplication | null>(null)
  const [account, setAccount] = useState<AccountInfo | null>(null)
  const [accessToken, setAccessToken] = useState('')
  const [workspaces, setWorkspaces] = useState<Workspace[]>([])
  const [selectedWorkspaceId, setSelectedWorkspaceId] = useState('')
  const [reports, setReports] = useState<PowerBiReport[]>([])
  const [dashboards, setDashboards] = useState<PowerBiDashboard[]>([])
  const [selectedContentKey, setSelectedContentKey] = useState('')
  const [status, setStatus] = useState('Add your Microsoft Entra app client ID to begin.')
  const [isBusy, setIsBusy] = useState(false)
  const [activeDepartment, setActiveDepartment] = useState<DepartmentName>('Parts')
  const [openDepartment, setOpenDepartment] = useState<DepartmentName | null>(null)
  const embedContainerRef = useRef<HTMLDivElement | null>(null)
  const powerBiServiceRef = useRef<service.Service | null>(null)

  const contentItems: PowerBiContent[] = [
    ...dashboards.map((dashboard) => ({
      key: `dashboard:${dashboard.id}`,
      kind: 'dashboard' as const,
      id: dashboard.id,
      name: dashboard.displayName,
      embedUrl: dashboard.embedUrl,
    })),
    ...reports.map((report) => ({
      key: `report:${report.id}`,
      kind: 'report' as const,
      id: report.id,
      name: report.name,
      embedUrl: report.embedUrl,
      webUrl: report.webUrl,
    })),
  ]
  const selectedContent = contentItems.find((content) => content.key === selectedContentKey)
  const activeDepartmentTab = departmentTabs.find(
    (department) => department.name === activeDepartment,
  )
  const canConnect = Boolean(config.clientId)

  function findContentForItem(item?: DepartmentItem) {
    if (!item?.targetName) {
      return undefined
    }

    const targetName = item.targetName.toLocaleLowerCase()
    return contentItems.find(
      (content) =>
        content.kind === item.kind &&
        content.name.toLocaleLowerCase() === targetName,
    )
  }

  useEffect(() => {
    if (!config.clientId) {
      setMsalApp(null)
      return
    }

    let isCurrent = true
    const app = createMsalApp(config)

    app
      .initialize()
      .then(async () => {
        if (!isCurrent) {
          return
        }

        setMsalApp(app)
        const redirectResult = await app.handleRedirectPromise()
        if (redirectResult?.account) {
          setAccount(redirectResult.account)
          setAccessToken(redirectResult.accessToken)
          setStatus(`Signed in as ${redirectResult.account.username}.`)
          return
        }

        const existingAccount = app.getAllAccounts()[0]
        if (!existingAccount) {
          setStatus('Ready to sign in with your Power BI account.')
          return
        }

        const tokenResult = await app.acquireTokenSilent({
          account: existingAccount,
          scopes: powerBiScopes,
        })

        setAccount(existingAccount)
        setAccessToken(tokenResult.accessToken)
        setStatus(`Signed in as ${existingAccount.username}.`)
      })
      .catch((error: unknown) => {
        setStatus(error instanceof Error ? error.message : 'Microsoft sign-in setup failed.')
      })

    return () => {
      isCurrent = false
    }
  }, [config])

  useEffect(() => {
    if (!accessToken) {
      setWorkspaces([])
      return
    }

    let isCurrent = true
    setIsBusy(true)
    setStatus('Loading your Power BI workspaces...')

    powerBiFetch<PowerBiWorkspacesResponse>('/groups', accessToken)
      .then((response) => {
        if (!isCurrent) {
          return
        }

        const targetWorkspace = response.value.find(
          (workspace) =>
            workspace.name.toLocaleLowerCase() === targetWorkspaceName.toLocaleLowerCase(),
        )

        if (!targetWorkspace) {
          setWorkspaces([])
          setSelectedWorkspaceId('')
          setStatus(`${targetWorkspaceName} workspace was not found for this account.`)
          return
        }

        const lockedWorkspace = { ...targetWorkspace, type: 'group' as const }
        setWorkspaces([lockedWorkspace])
        setSelectedWorkspaceId(lockedWorkspace.id)
        setStatus(`Locked to ${targetWorkspaceName} for ${targetAppName}.`)
      })
      .catch((error: unknown) => {
        setStatus(error instanceof Error ? error.message : 'Could not load workspaces.')
      })
      .finally(() => {
        if (isCurrent) {
          setIsBusy(false)
        }
      })

    return () => {
      isCurrent = false
    }
  }, [accessToken])

  useEffect(() => {
    if (!accessToken || !selectedWorkspaceId) {
      setReports([])
      setDashboards([])
      return
    }

    let isCurrent = true
    const reportsPath = `/groups/${selectedWorkspaceId}/reports`
    const dashboardsPath = `/groups/${selectedWorkspaceId}/dashboards`

    setIsBusy(true)
    setStatus('Loading Miles May app content...')

    Promise.all([
      powerBiFetch<PowerBiReportsResponse>(reportsPath, accessToken),
      powerBiFetch<PowerBiDashboardsResponse>(dashboardsPath, accessToken),
    ])
      .then(([reportsResponse, dashboardsResponse]) => {
        if (!isCurrent) {
          return
        }

        const loadedDashboards = dashboardsResponse.value
        const loadedReports = reportsResponse.value
        const partsDashboard = loadedDashboards.find(
          (dashboard) => dashboard.displayName.toLocaleLowerCase() === 'parts inventory',
        )
        const firstDashboard = partsDashboard ?? loadedDashboards[0]
        const firstReport = loadedReports[0]

        setDashboards(loadedDashboards)
        setReports(loadedReports)
        setSelectedContentKey(
          firstDashboard
            ? `dashboard:${firstDashboard.id}`
            : firstReport
              ? `report:${firstReport.id}`
              : '',
        )
        setStatus(
          loadedDashboards.length || loadedReports.length
            ? `Loaded ${targetAppName} content from ${targetWorkspaceName}.`
            : `No dashboards or reports were found in ${targetWorkspaceName}.`,
        )
      })
      .catch((error: unknown) => {
        setReports([])
        setDashboards([])
        setSelectedContentKey('')
        setStatus(error instanceof Error ? error.message : 'Could not load Miles May content.')
      })
      .finally(() => {
        if (isCurrent) {
          setIsBusy(false)
        }
      })

    return () => {
      isCurrent = false
    }
  }, [accessToken, selectedWorkspaceId])

  useEffect(() => {
    const container = embedContainerRef.current
    if (!container || !selectedContent || !accessToken) {
      return
    }

    if (!powerBiServiceRef.current) {
      powerBiServiceRef.current = new service.Service(
        factories.hpmFactory,
        factories.wpmpFactory,
        factories.routerFactory,
      )
    }

    powerBiServiceRef.current.reset(container)
    powerBiServiceRef.current.embed(container, {
      type: selectedContent.kind,
      id: selectedContent.id,
      embedUrl: selectedContent.embedUrl,
      accessToken,
      tokenType: models.TokenType.Aad,
      permissions: models.Permissions.Read,
      settings: {
        panes: {
          filters: { expanded: false, visible: true },
          pageNavigation: { visible: true },
        },
        background: models.BackgroundType.Transparent,
      },
    })
  }, [accessToken, selectedContent])

  function saveConfig() {
    const nextConfig = {
      clientId: draftClientId.trim(),
      tenantId: draftTenantId.trim() || 'common',
    }

    localStorage.setItem(clientIdStorageKey, nextConfig.clientId)
    localStorage.setItem(tenantIdStorageKey, nextConfig.tenantId)
    setConfig(nextConfig)
    setStatus('Connection settings saved. You can sign in now.')
  }

  async function signIn() {
    if (!msalApp) {
      setStatus('Save a client ID before signing in.')
      return
    }

    setIsBusy(true)
    try {
      setStatus('Redirecting to Microsoft sign-in...')
      await msalApp.loginRedirect({
        scopes: powerBiScopes,
        prompt: 'select_account',
      })
    } catch (error) {
      setStatus(error instanceof Error ? error.message : 'Sign-in failed.')
      setIsBusy(false)
    }
  }

  async function signOut() {
    if (msalApp && account) {
      await msalApp.logoutRedirect({ account })
      return
    }
    setAccount(null)
    setAccessToken('')
    setReports([])
    setDashboards([])
    setSelectedContentKey('')
    setSelectedWorkspaceId('')
    setStatus('Signed out.')
  }

  function chooseDepartment(departmentName: DepartmentName, item?: DepartmentItem) {
    setActiveDepartment(departmentName)
    setOpenDepartment(null)
    if (!item?.targetName) {
      setStatus(`${item?.label ?? departmentName} is ready for a custom view next.`)
      return
    }

    const matchingContent = findContentForItem(item)

    if (!matchingContent) {
      setStatus(`${item.targetName} was not found in ${targetWorkspaceName}.`)
      return
    }

    setSelectedContentKey(matchingContent.key)
    setStatus(`${item.targetName} selected from ${targetWorkspaceName}.`)
  }

  return (
    <main className="app-shell">
      <section className="masthead">
        <div>
          <p className="eyebrow">Power BI app project</p>
          <h1>Your reports, inside your own React app.</h1>
          <p className="lede">
            Connect with Microsoft sign-in, browse the workspaces you can access,
            and embed a report from Power BI Service.
          </p>
        </div>
        <div className="status-panel" aria-live="polite">
          <span className={accessToken ? 'status-dot connected' : 'status-dot'} />
          <p>{status}</p>
        </div>
      </section>

      <nav className="department-tabs" aria-label="Parts and service app sections">
        {departmentTabs.map((department) => {
          const isOpen = openDepartment === department.name
          const isActive = activeDepartment === department.name

          return (
            <div className="department-tab" key={department.name}>
              <button
                type="button"
                className={isActive ? 'department-button active' : 'department-button'}
                aria-expanded={isOpen}
                aria-controls={`${department.name.toLowerCase()}-menu`}
                onClick={() => {
                  setActiveDepartment(department.name)
                  setOpenDepartment(isOpen ? null : department.name)
                }}
              >
                <span>{department.name}</span>
                <span aria-hidden="true">v</span>
              </button>
              {isOpen && (
                <div className="department-menu" id={`${department.name.toLowerCase()}-menu`}>
                  <p>{department.summary}</p>
                  <div className="department-menu-items">
                    {department.items.map((item) => (
                      <button
                        type="button"
                        className={
                          findContentForItem(item) ? 'menu-item connected' : 'menu-item pending'
                        }
                        key={item.label}
                        onClick={() => chooseDepartment(department.name, item)}
                      >
                        <span>{item.label}</span>
                        <span>{item.targetName ? 'Connected' : 'Design next'}</span>
                      </button>
                    ))}
                  </div>
                </div>
              )}
            </div>
          )
        })}
      </nav>

      <section className="department-summary" aria-live="polite">
        <p className="section-label">Current focus</p>
        <h2>{activeDepartment}</h2>
        <p>{activeDepartmentTab?.summary}</p>
      </section>

      <section className="workspace-grid">
        <aside className="control-panel">
          <div className="panel-section">
            <p className="section-label">1. App registration</p>
            <label htmlFor="client-id">Application client ID</label>
            <input
              id="client-id"
              placeholder="00000000-0000-0000-0000-000000000000"
              value={draftClientId}
              onChange={(event) => setDraftClientId(event.target.value)}
            />

            <label htmlFor="tenant-id">Tenant ID or common</label>
            <input
              id="tenant-id"
              value={draftTenantId}
              onChange={(event) => setDraftTenantId(event.target.value)}
            />

            <button type="button" onClick={saveConfig}>
              Save connection settings
            </button>
          </div>

          <div className="panel-section">
            <p className="section-label">2. Sign in</p>
            <button type="button" onClick={signIn} disabled={!canConnect || isBusy}>
              {account ? 'Refresh Power BI token' : 'Sign in to Power BI'}
            </button>
            {account && (
              <button type="button" className="secondary" onClick={signOut}>
                Sign out
              </button>
            )}
          </div>

          <div className="panel-section">
            <p className="section-label">3. App content</p>
            <label htmlFor="workspace">Workspace</label>
            <select
              id="workspace"
              value={selectedWorkspaceId}
              onChange={(event) => setSelectedWorkspaceId(event.target.value)}
              disabled
            >
              {!workspaces.length && <option value="">Locked to {targetWorkspaceName}</option>}
              {workspaces.map((workspace) => (
                <option key={workspace.id} value={workspace.id}>
                  {workspace.name}
                </option>
              ))}
            </select>

            <label htmlFor="content">Dashboard or report</label>
            <select
              id="content"
              value={selectedContentKey}
              onChange={(event) => setSelectedContentKey(event.target.value)}
              disabled={!contentItems.length || isBusy}
            >
              {contentItems.map((content) => (
                <option key={content.key} value={content.key}>
                  {content.name} ({content.kind})
                </option>
              ))}
            </select>
          </div>
        </aside>

        <section className="report-stage" aria-label="Power BI report viewer">
          {selectedContent ? (
            <>
              <div className="report-toolbar">
                <div>
                  <p className="section-label">Embedded {selectedContent.kind}</p>
                  <h2>{selectedContent.name}</h2>
                </div>
                {selectedContent.webUrl && (
                  <a href={selectedContent.webUrl} target="_blank" rel="noreferrer">
                    Open in Power BI
                  </a>
                )}
              </div>
              <div className="embed-frame" ref={embedContainerRef} />
            </>
          ) : (
            <div className="empty-state">
              <p className="section-label">Waiting for connection</p>
              <h2>Set up Microsoft sign-in to display your Power BI data.</h2>
              <p>
                Create an app registration in Azure, add this local URL as a
                single-page app redirect URI, then paste the client ID here.
              </p>
            </div>
          )}
        </section>
      </section>

      <section className="setup-strip">
        <div>
          <strong>Azure redirect URI</strong>
          <span>{window.location.origin}</span>
        </div>
        <div>
          <strong>Power BI API permissions</strong>
          <span>Workspace.Read.All, Dashboard.Read.All, Report.Read.All, Dataset.Read.All</span>
        </div>
      </section>
    </main>
  )
}

export default App