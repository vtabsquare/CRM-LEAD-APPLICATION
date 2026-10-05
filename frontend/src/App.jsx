import React, { useState, useMemo, useEffect } from 'react';
import axios from 'axios';
import { Routes, Route, Navigate, useLocation } from "react-router-dom";
import { User, XCircle, Loader2, ShieldCheck, Lock, Activity, CheckCircle2 } from 'lucide-react';

// Components
import Sidebar from './Sidebar';
import EnquiryPage from './EnquiryPage';
import AdmissionRegistration from './AdmissionRegistration';
import BillingSummary from "./BillingSummary";
import AnalyticsReport from "./AnalysisReport";
import SchemaEditor from './SchemaEditor';
import ChargeSummary from './ChargeSummary';
import Documents from './Documents';
import NotificationSettings from './NotificationSettings';
import FileManager from './FileManager';
import BedManagement from './BedManagement';
import SearchData from './SearchData';
import AIChat from './AIChat';
import Home from './Home';

// Invoice Module
import InvoiceList from './Invoice/InvoiceList';
import InvoiceCreate from './Invoice/InvoiceCreate';
import InvoiceCreateNew from './Invoice/InvoiceCreateNew';
import InvoiceView from './Invoice/InvoiceView';
import InvoicePageNew from './Invoice/InvoicePageNew';
import InvoiceUpload from './Invoice/InvoiceUpload';
import InvoiceMonitor from './Invoice/InvoiceMonitor';
import ServiceCatalog from './components/ServiceCatalog';

// Home Care Module
import HomeCareList from './HomeCare/HomeCareList';
import HomeCareBillingHistory from './HomeCare/HomeCareBillingHistory';
import HomeCareBillingPreview from './HomeCare/HomeCareBillingPreview';
import HomeCareCreate from './HomeCare/HomeCareCreate';
import HomeCareEdit from './HomeCare/HomeCareEdit';

// Patient Admission Module
import PatientAdmissionList from './PatientAdmission/PatientAdmissionList';
import PatientAdmissionBillingHistory from './PatientAdmission/PatientAdmissionBillingHistory';
import PatientAdmissionBillingPreview from './PatientAdmission/PatientAdmissionBillingPreview';
import PatientAdmissionCreate from './PatientAdmission/PatientAdmissionCreate';
import PatientAdmissionEdit from './PatientAdmission/PatientAdmissionEdit';

// Import API configuration
import API_BASE_URL from './config';
import { warmUpBackend } from './api';

const LOGIN_STATS = [
  { label: 'Monthly Admissions Processed', value: '1.2K+' },
  { label: 'Decision Center Uptime', value: '99.9%' },
  { label: 'Avg. Response Velocity', value: '2.3s' }
];

const LOGIN_FEATURES = [
  'Unified admissions + home care workspace',
  'Role-aware governance & change tracking',
  'Instant sync with Google Sheets data lake'
];

function App() {
  // Initialize state from localStorage if available
  const [isAuthenticated, setIsAuthenticated] = useState(() => {
    return localStorage.getItem('isAuthenticated') === 'true';
  });
  const [loginUser, setLoginUser] = useState(() => {
    return localStorage.getItem('loginUser') || '';
  });
  const [loginPass, setLoginPass] = useState('');
  const [loginLoading, setLoginLoading] = useState(false);
  const [loginError, setLoginError] = useState('');
  const [isCollapsed, setIsCollapsed] = useState(false);
  const location = useLocation();

  // Warm up the backend immediately on app load (handles Render free-tier cold starts)
  useEffect(() => {
    warmUpBackend();
  }, []);

  // Save authentication state to localStorage whenever it changes
  useEffect(() => {
    localStorage.setItem('isAuthenticated', isAuthenticated.toString());
  }, [isAuthenticated]);

  // Save username to localStorage whenever it changes
  useEffect(() => {
    if (loginUser) {
      localStorage.setItem('loginUser', loginUser);
    }
  }, [loginUser]);

  const todayDisplay = useMemo(() => {
    return new Date().toLocaleDateString('en-IN', {
      day: '2-digit',
      month: 'short',
      year: 'numeric'
    });
  }, []);

  const handleLogin = async (e) => {
    e.preventDefault();
    try {
      setLoginLoading(true);
      setLoginError('');

      if (!loginUser.trim() || !loginPass.trim()) {
        setLoginError('Please enter both username and password');
        setLoginLoading(false);
        return;
      }

      const response = await axios.post(`${API_BASE_URL}/login`, {
        User_name: loginUser.trim(),
        Password: loginPass.trim()
      });

      if (response.data?.status === 'ok') {
        // Store JWT token for API authentication
        if (response.data.token) {
          localStorage.setItem('authToken', response.data.token);
        }
        setIsAuthenticated(true);
        setLoginError('');
      } else {
        setLoginError('Invalid username or password');
      }
    } catch (err) {
      let errorMessage = 'Login failed. Please try again.';
      if (err.response?.status === 401) {
        errorMessage = 'Invalid username or password';
      } else if (err.response?.data?.detail) {
        errorMessage = err.response.data.detail;
      }
      setLoginError(errorMessage);
    } finally {
      setLoginLoading(false);
    }
  };

  const handleLogout = () => {
    setIsAuthenticated(false);
    setLoginUser('');
    setLoginPass('');
    setLoginError('');
    // Clear localStorage on logout
    localStorage.removeItem('isAuthenticated');
    localStorage.removeItem('loginUser');
    localStorage.removeItem('authToken');
  };

  // Render Login Screen if not authenticated
  if (!isAuthenticated) {
    return (
      <div className="relative min-h-screen overflow-hidden bg-slate-950 text-white">
        <div className="login-aurora" />
        <div className="login-aurora delay-150" />
        <div className="login-grid" />
        <div className="relative z-10 flex min-h-screen flex-col items-center justify-center gap-10 px-6 py-12 lg:flex-row lg:items-stretch">
          <div className="w-full max-w-3xl space-y-8">
            <div className="inline-flex items-center gap-2 rounded-full border border-white/20 bg-white/10 px-4 py-1 text-xs font-semibold tracking-widest uppercase text-emerald-200">
              <ShieldCheck className="h-4 w-4" />
              Trusted Healthcare CRM Suite
            </div>
            <div className="space-y-6">
              <h1 className="text-4xl font-semibold leading-tight text-white md:text-5xl">
                Intelligent lead orchestration for <span className="text-emerald-300">home care&nbsp;&amp; admissions.</span>
              </h1>
              <p className="text-lg text-white/80 md:text-xl">
                Elevate every conversation with real-time visibility, AI-assisted forecasts, and dependable governance—all wrapped in a secure corporate workspace.
              </p>
            </div>
            <div className="grid gap-4 md:grid-cols-3">
              {LOGIN_STATS.map((stat) => (
                <div
                  key={stat.label}
                  className="rounded-2xl border border-white/15 bg-white/10 px-4 py-5 shadow-lg shadow-black/20 backdrop-blur"
                >
                  <p className="text-3xl font-semibold text-white">{stat.value}</p>
                  <p className="mt-2 text-xs font-semibold uppercase tracking-wide text-white/70">
                    {stat.label}
                  </p>
                </div>
              ))}
            </div>
            <div className="rounded-3xl border border-white/15 bg-white/5 p-6 shadow-2xl shadow-black/30 backdrop-blur">
              <div className="mb-4 flex items-center gap-3 text-sm font-semibold uppercase tracking-[0.3em] text-white/60">
                <Activity className="h-4 w-4 text-emerald-300" />
                Platform Intelligence Layer
              </div>
              <ul className="space-y-3 text-sm text-white/85">
                {LOGIN_FEATURES.map((feature) => (
                  <li key={feature} className="flex items-start gap-3">
                    <span className="rounded-full bg-emerald-500/20 p-1.5 text-emerald-200">
                      <CheckCircle2 className="h-4 w-4" />
                    </span>
                    <span>{feature}</span>
                  </li>
                ))}
              </ul>
            </div>
          </div>

          <div className="w-full max-w-md">
            <div className="rounded-[32px] border border-white/60 bg-white/95 shadow-2xl shadow-emerald-900/40 backdrop-blur-xl">
              <div className="border-b border-slate-100 px-8 py-6">
                <p className="text-xs font-semibold uppercase tracking-[0.4em] text-slate-500">
                  Login
                </p>
                <h3 className="mt-2 text-2xl font-semibold text-slate-900">CRM Control Center</h3>
                <p className="mt-3 flex items-center gap-2 text-sm font-medium text-emerald-600">
                  <Lock className="h-4 w-4" />
                  Encrypted sessions with adaptive security
                </p>
              </div>

              <div className="px-8 py-8">
                {loginError && (
                  <div className="mb-4 flex items-center gap-2 rounded-2xl border border-red-200 bg-red-50/80 px-4 py-3 text-sm font-medium text-red-800">
                    <XCircle className="h-5 w-5" />
                    {loginError}
                  </div>
                )}

                <form onSubmit={handleLogin} className="space-y-5">
                  <div className="space-y-2">
                    <label className="text-sm font-semibold text-slate-700">Username</label>
                    <div className="relative">
                      <input
                        type="text"
                        value={loginUser}
                        onChange={(e) => setLoginUser(e.target.value)}
                        className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 pr-12 text-slate-900 shadow-inner placeholder-slate-400 transition focus:border-emerald-500 focus:ring-2 focus:ring-emerald-100"
                        placeholder="Enter your corporate ID"
                        required
                        autoComplete="username"
                      />
                      <User className="absolute right-4 top-1/2 h-5 w-5 -translate-y-1/2 text-slate-400" />
                    </div>
                  </div>

                  <div className="space-y-2">
                    <label className="text-sm font-semibold text-slate-700">Password</label>
                    <div className="relative">
                      <input
                        type="password"
                        value={loginPass}
                        onChange={(e) => setLoginPass(e.target.value)}
                        className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 pr-12 text-slate-900 shadow-inner placeholder-slate-400 transition focus:border-emerald-500 focus:ring-2 focus:ring-emerald-100"
                        placeholder="••••••••"
                        required
                        autoComplete="current-password"
                      />
                      <Lock className="absolute right-4 top-1/2 h-5 w-5 -translate-y-1/2 text-slate-400" />
                    </div>
                  </div>

                  <div className="flex items-center justify-between text-xs font-medium text-slate-500">
                    <div className="flex items-center gap-2">
                      <span className="inline-block h-2 w-2 rounded-full bg-emerald-500 animate-pulse" />
                      ISO 27001 compliant workspace
                    </div>
                    <a href="mailto:support@grandworld.com" className="text-emerald-600 hover:text-emerald-500">
                      Contact Support
                    </a>
                  </div>

                  <button
                    type="submit"
                    disabled={loginLoading}
                    className="relative w-full overflow-hidden rounded-2xl bg-gradient-to-r from-blue-600 via-emerald-500 to-green-500 py-3 text-base font-semibold text-white shadow-lg shadow-emerald-900/30 transition hover:shadow-2xl disabled:opacity-70"
                  >
                    <span className="relative z-10">
                      {loginLoading ? (
                        <span className="flex items-center justify-center gap-2">
                          <Loader2 className="h-5 w-5 animate-spin" />
                          Securing Session...
                        </span>
                      ) : (
                        'Launch Workspace'
                      )}
                    </span>
                    <span className="absolute inset-0 bg-white/10 opacity-0 transition group-hover:opacity-100" />
                  </button>
                </form>

                <div className="mt-6 rounded-2xl border border-slate-100 bg-slate-50/80 p-4 text-xs text-slate-500">
                  By signing in you agree to the Grand World Security Policy and CRM Usage Guidelines.
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col">
      {/* Subtle gradient background */}
      <div className="fixed inset-0 bg-gradient-to-br from-blue-50/50 via-emerald-50/50 to-blue-50/50 pointer-events-none" />
      {/* Grid overlay */}
      <div className="fixed inset-0 bg-[linear-gradient(to_right,#80808012_1px,transparent_1px),linear-gradient(to_bottom,#80808012_1px,transparent_1px)] bg-[size:32px_32px] pointer-events-none" />
      
      {/* Main content */}
      <div className="relative flex-1 flex flex-col z-10">
        <header className="sticky top-0 z-40 bg-gradient-to-r from-blue-500 to-green-500 border-b-4 border-green-600 shadow-lg">
          <div className="w-full px-4 py-1.5 flex items-center justify-between">
            <div className="flex items-center gap-2">
              <img
                src="/grand-world-logo.svg"
                alt="Grand World Elder Care"
                className="w-12 h-10 shadow-lg bg-white/90 object-contain p-2 rounded"
              />
              <div>
                <h1 className="text-2xl font-semibold text-white drop-shadow-lg">CRM Lead Application</h1>
              </div>
            </div>
            <div className="flex items-center gap-2">
              <div className="hidden sm:flex items-center gap-2 text-xs text-white bg-white/20 px-3 py-1.5 backdrop-blur rounded">
                <div className="h-1.5 w-3 rounded-full bg-green-300 animate-pulse" />
                System Ready
              </div>

              <div className="flex items-center gap-3 bg-white/20 px-4 py-2 backdrop-blur rounded-md">
                <div className="text-white text-sm opacity-90">
                  User: <span className="font-semibold">{loginUser}</span>
                </div>
                <div className="text-white text-sm opacity-90">
                  Date: <span className="font-semibold">{todayDisplay}</span>
                </div>
                <div className="flex gap-2">
                  <button
                    onClick={handleLogout}
                    className="px-4 py-1.5 text-sm font-medium rounded-md transition-all bg-gray-400 hover:bg-red-500 text-white shadow-md"
                  >
                    <XCircle className="w-4 h-4 inline mr-1" />
                    Logout
                  </button>
                </div>
              </div>
            </div>
          </div>
        </header>

        <div className="flex-1 flex">
          {/* Sidebar */}
          <Sidebar />

          {/* Main Content Area */}
          <main className="flex-1 h-screen overflow-y-auto animate-fade-in">
            <Routes>
              {/* Home */}
              <Route path="/" element={<Home />} />
              <Route path="/home" element={<Home />} />
              <Route path="/search" element={<SearchData />} />

              {/* Enquiries */}
              <Route path="/enquiries" element={<EnquiryPage />} />

              {/* Patient Admission */}
              <Route path="/admission" element={<AdmissionRegistration />} />
              <Route path="/patientadmission/clients" element={<PatientAdmissionList />} />
              <Route path="/patientadmission/create" element={<PatientAdmissionCreate />} />
              <Route path="/patientadmission/edit/:patientName" element={<PatientAdmissionEdit />} />
              <Route path="/patientadmission/billing-history/:patientName" element={<PatientAdmissionBillingHistory />} />
              <Route path="/patientadmission/billing-preview" element={<PatientAdmissionBillingPreview />} />
              <Route path="/bed-availability" element={<BedManagement />} />
              <Route path="/billing-summary" element={<BillingSummary />} />

              {/* Home Care */}
              <Route path="/homecare/clients" element={<HomeCareList />} />
              <Route path="/homecare/create" element={<HomeCareCreate />} />
              <Route path="/homecare/edit/:patientName" element={<HomeCareEdit />} />
              <Route path="/homecare/billing-history/:patientName" element={<HomeCareBillingHistory />} />
              <Route path="/homecare/billing-preview" element={<HomeCareBillingPreview />} />

              {/* Finance */}
              <Route path="/invoice" element={<InvoiceList />} />
              <Route path="/invoice/create" element={<InvoiceCreate />} />
              <Route path="/invoice/create-new" element={<InvoiceCreateNew />} />
              <Route path="/invoice/view/:invoiceId" element={<InvoiceView />} />
              <Route path="/invoice/new" element={<InvoicePageNew />} />
              <Route path="/invoice/upload" element={<InvoiceUpload />} />
              <Route path="/invoice/monitor" element={<InvoiceMonitor />} />
              <Route path="/service-catalog" element={<ServiceCatalog />} />

              {/* Management */}
              <Route path="/documents" element={<Documents />} />
              <Route path="/schema" element={<SchemaEditor />} />
              <Route path="/charge-summary" element={<ChargeSummary />} />
              <Route path="/notifications" element={<NotificationSettings />} />
              <Route path="/file-manager" element={<FileManager />} />

              {/* Analytics */}
              <Route path="/analysis" element={<AnalyticsReport />} />

              <Route path="*" element={<Navigate to="/" replace />} />
            </Routes>
          </main>
        </div>
      </div>

      {/* AI Chat Overlay */}
      <AIChat />
    </div>
  );
}

export default App;
