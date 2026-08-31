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

type DepartmentName = 'Parts' | 'Service'

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

type AppRoute = {
  department: DepartmentName
  label: string
}

type TemplateCard = {
  title: string
  detail: string
}

type ServiceKpi = {
  label: string
  value: string
  active?: boolean
}

type ServiceMetricOption = {
  label: string
  value: string
  active?: boolean
}

type ServiceMonthlyRow = {
  year: string
  months: string[]
  total: string
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
      { label: 'Service Metrics', targetName: 'PY&OMG Service', kind: 'dashboard' },
      { label: 'Efficiency Calendar' },
      { label: 'Service Efficiency' },
      { label: 'Cashiered Detail' },
      { label: 'Service Logged' },
      { label: 'RO Aging' },
      { label: 'Timeclock' },
    ],
  },
]

const dealerTiles = ['PBC', 'OMG', 'LCN', 'AL', 'ETX', 'MS', 'STX']

const metricTiles = [
  { label: 'ROs', value: '2K' },
  { label: 'Hours', value: '12K' },
  { label: 'Labor', value: '1.6M' },
  { label: 'Parts', value: '1.4M' },
  { label: 'Profit', value: '2.4M' },
]

const serviceMetricOptions: ServiceMetricOption[] = [
  { label: 'RO#', value: '3K' },
  { label: 'Hours', value: '13K' },
  { label: 'Labor', value: '1.6M', active: true },
  { label: 'Parts', value: '1.5M' },
  { label: 'OTC', value: '929K' },
  { label: 'Revenue', value: '4.0M' },
  { label: 'Profit', value: '2.5M' },
  { label: 'Parts %', value: '37.5%' },
  { label: 'Labor / RO', value: '642' },
  { label: 'Parts / RO', value: '575' },
  { label: 'Total / RO', value: '2K' },
  { label: 'Hours / RO', value: '5' },
  { label: 'Billable Labor', value: '3.0%' },
  { label: 'OTC Profit', value: '347K' },
  { label: 'OTC %', value: '37.3%' },
  { label: 'Days In Service', value: '181' },
  { label: 'Boat Sales', value: '54.1M' },
]

const serviceTopKpis: ServiceKpi[] = [
  { label: 'E', value: '1.6M', active: true },
  { label: 'I', value: '1.5M' },
  { label: 'AL', value: '128K' },
  { label: 'ETX', value: '455K' },
  { label: 'MS', value: '375K' },
  { label: 'STX', value: '668K' },
  { label: 'BMT', value: '146K' },
  { label: 'HOU', value: '214K' },
  { label: 'JSP', value: '95K' },
  { label: 'AP', value: '302K' },
  { label: 'CC', value: '240K' },
  { label: 'SA', value: '125K' },
  { label: 'OS', value: '246K' },
  { label: 'GP', value: '129K' },
  { label: 'DA', value: '81K' },
  { label: 'FOL', value: '48K' },
]

const serviceMonths = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug']

const serviceAnnualTotals = [
  { year: '21', value: '1.9M', change: '19%', height: '56%' },
  { year: '22', value: '2.2M', change: '16%', height: '71%' },
  { year: '23', value: '2.3M', change: '8%', height: '77%' },
  { year: '24', value: '2.3M', change: '-1%', height: '73%' },
  { year: '25', value: '2.4M', change: '2%', height: '80%' },
  { year: '26', value: '1.6M', change: '5%', height: '49%' },
]

const serviceMonthlyRows: ServiceMonthlyRow[] = [
  { year: '2026', months: ['118K', '182K', '217K', '212K', '195K', '219K', '262K', '221K'], total: '1.6M' },
  { year: '2025', months: ['129K', '154K', '205K', '216K', '209K', '184K', '202K', '244K'], total: '2.4M' },
  { year: '2024', months: ['92K', '174K', '227K', '190K', '197K', '210K', '191K', '254K'], total: '2.3M' },
  { year: '2023', months: ['145K', '197K', '234K', '209K', '231K', '224K', '162K', '244K'], total: '2.3M' },
  { year: '2022', months: ['124K', '160K', '217K', '201K', '177K', '208K', '189K', '218K'], total: '2.2M' },
  { year: '2021', months: ['122K', '109K', '179K', '171K', '137K', '153K', '181K', '162K'], total: '1.9M' },
]

const serviceStoreBars = [
  { label: 'AP', value: '302K', change: '3%', height: '92%', tone: 'bad' },
  { label: 'OS', value: '246K', change: '4%', height: '75%', tone: 'good' },
  { label: 'CC', value: '240K', change: '4%', height: '73%', tone: 'good' },
  { label: 'HOU', value: '214K', change: '16%', height: '65%', tone: 'good' },
  { label: 'BMT', value: '146K', change: '66%', height: '45%', tone: 'good' },
  { label: 'GP', value: '129K', change: '53%', height: '39%', tone: 'good' },
  { label: 'SA', value: '125K', change: '-3%', height: '38%', tone: 'bad' },
  { label: 'JSP', value: '95K', change: '18%', height: '29%', tone: 'good' },
  { label: 'DA', value: '81K', change: '-21%', height: '25%', tone: 'bad' },
  { label: 'FOL', value: '48K', change: '87%', height: '15%', tone: 'good' },
]

function ServiceMetricsDashboard() {
  return (
    <section className="service-metrics-dashboard" aria-label="Service Metrics dashboard">
      <header className="service-metrics-header">
        <div className="service-metrics-brand">
          <strong>Premier</strong>
          <span>Yamaha</span>
          <small>Boat Centers</small>
        </div>
        <h2>Service</h2>
        <div className="service-top-kpis" aria-label="Service store totals">
          <div className="service-top-kpi-group service-top-kpi-primary">
            {serviceTopKpis.slice(0, 2).map((kpi) => (
              <button
                type="button"
                className={kpi.active ? 'service-top-kpi active' : 'service-top-kpi'}
                key={kpi.label}
              >
                <strong>{kpi.label}</strong>
                <span>{kpi.value}</span>
              </button>
            ))}
          </div>
          <div className="service-top-kpi-group service-top-kpi-stores">
            {serviceTopKpis.slice(2).map((kpi) => (
              <button
                type="button"
                className={kpi.active ? 'service-top-kpi active' : 'service-top-kpi'}
                key={kpi.label}
              >
                <strong>{kpi.label}</strong>
                <span>{kpi.value}</span>
              </button>
            ))}
          </div>
        </div>
        <div className="service-month-tabs" aria-label="Service month filters">
          {serviceMonths.map((month) => (
            <button type="button" key={month}>
              {month}
            </button>
          ))}
        </div>
      </header>

      <div className="service-dashboard-grid">
        <aside className="service-metric-selector" aria-label="Select metric">
          <div>Select Metric</div>
          {serviceMetricOptions.map((metric) => (
            <button
              type="button"
              className={metric.active ? 'service-metric-option active' : 'service-metric-option'}
              key={metric.label}
            >
              <span>{metric.label}</span>
              <strong>{metric.value}</strong>
            </button>
          ))}
        </aside>

        <div className="service-dashboard-main">
          <div className="service-year-tabs" aria-label="Service year filters">
            {['2023', '2024', '2025', '2026'].map((year) => (
              <button type="button" className={year === '2026' ? 'active' : undefined} key={year}>
                {year}
              </button>
            ))}
          </div>

          <section className="service-headline-panel">
            <h3>Labor</h3>
            <strong>1.6M</strong>
            <span>PY: 1.5M (+.1M +5.5%)</span>
          </section>

          <section className="service-chart-row">
            <article className="service-chart-panel service-bar-panel">
              <div className="service-region-filter">
                <button type="button">Region</button>
                <button type="button" className="active">LCN</button>
              </div>
              <div className="service-bars" aria-label="Labor by store">
                {serviceStoreBars.map((bar) => (
                  <div className="service-bar-column" key={bar.label}>
                    <span>{bar.value}</span>
                    <small>{bar.change}</small>
                    <div
                      className={bar.tone === 'bad' ? 'service-bar bad' : 'service-bar'}
                      style={{ height: bar.height }}
                    />
                    <strong>{bar.label}</strong>
                  </div>
                ))}
              </div>
            </article>
          </section>

          <section className="service-bottom-row">
            <article className="service-chart-panel service-annual-panel">
              <h4>Annual</h4>
              <div className="service-annual-bars">
                {serviceAnnualTotals.map((year) => (
                  <div className="service-annual-column" key={year.year}>
                    <span>{year.value}</span>
                    <small>{year.change}</small>
                    <div style={{ height: year.height }} />
                    <strong>{year.year}</strong>
                  </div>
                ))}
              </div>
            </article>

            <article className="service-chart-panel service-monthly-panel">
              <h4>Monthly</h4>
              <table>
                <thead>
                  <tr>
                    <th>Year</th>
                    {serviceMonths.map((month) => <th key={month}>{month}</th>)}
                    <th>Total</th>
                  </tr>
                </thead>
                <tbody>
                  {serviceMonthlyRows.map((row) => (
                    <tr key={row.year}>
                      <th>{row.year}</th>
                      {row.months.map((month, index) => <td key={`${row.year}-${serviceMonths[index]}`}>{month}</td>)}
                      <td>{row.total}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </article>

            <article className="service-chart-panel service-allocation-panel">
              <h4>Allocation</h4>
              <div className="service-donut" aria-label="Labor allocation by store" />
              <div className="service-donut-labels">
                <span>AP 19%</span>
                <span>OS 15%</span>
                <span>CC 15%</span>
                <span>HOU 13%</span>
                <span>BMT 9%</span>
                <span>GP 8%</span>
                <span>SA 8%</span>
                <span>JSP 6%</span>
              </div>
            </article>
          </section>
        </div>
      </div>
    </section>
  )
}

function getTemplateCards(route: AppRoute): TemplateCard[] {
  if (route.department === 'Service') {
    return [
      { title: 'Repair orders', detail: 'Open ROs, aging, promised dates, and service advisor ownership.' },
      { title: 'Revenue', detail: 'Labor, parts, warranty, and internal sales trend sections.' },
      { title: 'Throughput', detail: 'Hours per RO, cycle time, and stalled work indicators.' },
    ]
  }

  return [
    { title: 'Inventory health', detail: 'On-hand value, aging, turns, and stocked versus special-order parts.' },
    { title: 'Sales and margin', detail: 'Parts sales, cost, gross profit, and margin by category.' },
    { title: 'Exceptions', detail: 'Backorders, lost sales, negative quantity, and variance alerts.' },
  ]
}

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
  const [selectedRoute, setSelectedRoute] = useState<AppRoute>({
    department: 'Parts',
    label: 'Parts Inventory',
  })
  const [openDepartment, setOpenDepartment] = useState<DepartmentName | null>(null)
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false)
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
  const templateCards = getTemplateCards(selectedRoute)
  const canConnect = Boolean(config.clientId)
  const isServiceMetricsRoute =
    selectedRoute.department === 'Service' && selectedRoute.label === 'Service Metrics'

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
    setSelectedRoute({ department: departmentName, label: item?.label ?? departmentName })
    setOpenDepartment(departmentName)
    if (!item?.targetName) {
      setSelectedContentKey('')
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
    <main className={isSidebarCollapsed ? 'app-shell sidebar-collapsed' : 'app-shell'}>
      <aside className="side-nav" aria-label="Parts and Service navigation">
        <button
          type="button"
          className="collapse-button"
          aria-label={isSidebarCollapsed ? 'Expand navigation' : 'Collapse navigation'}
          onClick={() => setIsSidebarCollapsed((isCollapsed) => !isCollapsed)}
        >
          {isSidebarCollapsed ? '>' : '<<'}
        </button>

        <div className="side-brand">
          <span className="side-logo">PBC</span>
          <strong>Premier Yamaha</strong>
          <span>{targetAppName}</span>
        </div>

        <nav className="side-tabs" aria-label="App tabs">
          {departmentTabs.map((department) => {
            const isOpen = openDepartment === department.name
            const isActive = activeDepartment === department.name

            return (
              <div className="side-tab-group" key={department.name}>
                <button
                  type="button"
                  className={isActive ? 'side-tab active' : 'side-tab'}
                  data-short-label={department.name.slice(0, 1)}
                  aria-expanded={isOpen}
                  aria-controls={`${department.name.toLowerCase()}-side-menu`}
                  onClick={() => {
                    setActiveDepartment(department.name)
                    setSelectedRoute({ department: department.name, label: department.name })
                    setOpenDepartment(isOpen ? null : department.name)
                  }}
                >
                  <span>{department.name}</span>
                  <span aria-hidden="true">v</span>
                </button>

                {isOpen && !isSidebarCollapsed && (
                  <div className="side-subnav" id={`${department.name.toLowerCase()}-side-menu`}>
                    {department.items.map((item) => (
                      <button
                        type="button"
                        className={
                          selectedRoute.label === item.label ? 'side-subtab active' : 'side-subtab'
                        }
                        key={item.label}
                        onClick={() => chooseDepartment(department.name, item)}
                      >
                        <span>{item.label}</span>
                        <span>{item.targetName ? 'Live' : 'Draft'}</span>
                      </button>
                    ))}
                  </div>
                )}
              </div>
            )
          })}
        </nav>

        <div className="side-footer">Miles May workspace</div>
      </aside>

      <div className="app-content">
        {isServiceMetricsRoute ? (
          <ServiceMetricsDashboard />
        ) : (
          <>
      <section className="masthead">
        <div>
          <p className="eyebrow">Premier Boat Center</p>
          <h1>{targetAppName}</h1>
          <p className="lede">
            Custom Parts and Service command center for the Miles May workspace.
          </p>
        </div>
        <div className="status-panel" aria-live="polite">
          <span className={accessToken ? 'status-dot connected' : 'status-dot'} />
          <p>{status}</p>
        </div>
      </section>

      <section className="scoreboard" aria-label="Parts and service metric shortcuts">
        {metricTiles.map((tile) => (
          <div className="score-tile" key={tile.label}>
            <span>{tile.label}</span>
            <strong>{tile.value}</strong>
          </div>
        ))}
      </section>

      <section className="dealer-strip" aria-label="Dealer filters">
        {dealerTiles.map((dealer) => (
          <span className={dealer === 'LCN' ? 'dealer-pill active' : 'dealer-pill'} key={dealer}>
            {dealer}
          </span>
        ))}
      </section>

      <section className="department-summary" aria-live="polite">
        <p className="section-label">Current focus</p>
        <h2>{activeDepartment}</h2>
        <p>{activeDepartmentTab?.summary}</p>
        <div className="route-chip">Route: {selectedRoute.department} / {selectedRoute.label}</div>
      </section>

      <section className="workspace-grid">
        <aside className="control-panel">
          <div className="panel-brand">
            <span className="brand-mark">PBC</span>
            <strong>{activeDepartment}</strong>
            <span>{targetWorkspaceName}</span>
          </div>
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
            <div className="metric-list" aria-label="Metric shortcuts">
              {metricTiles.map((tile) => (
                <button
                  type="button"
                  className={tile.label === 'Profit' ? 'metric-button active' : 'metric-button'}
                  key={tile.label}
                >
                  <span>{tile.label}</span>
                  <strong>{tile.value}</strong>
                </button>
              ))}
            </div>
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
            <div className="template-board">
              <p className="section-label">{selectedRoute.department} template</p>
              <h2>{selectedRoute.label}</h2>
              <p>
                This route is ready for layout design. Once a matching Power BI
                dashboard, report, or Lightspeed export is assigned, it can be connected here.
              </p>
              {!accessToken && (
                <div className="template-callout">
                  Sign in to load Miles May Power BI content. The template can be designed now.
                </div>
              )}
              <div className="template-card-grid">
                {templateCards.map((card) => (
                  <article className="template-card" key={card.title}>
                    <strong>{card.title}</strong>
                    <span>{card.detail}</span>
                  </article>
                ))}
              </div>
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
          </>
        )}
      </div>
    </main>
  )
}

export default App