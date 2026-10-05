/**
 * Centralized API Client
 * 
 * All API calls should use this module instead of raw axios.
 * Automatically attaches JWT auth token to every request and
 * handles 401 (unauthorized) responses by redirecting to login.
 * Also handles Render free-tier cold-starts with auto-retry on timeout.
 */
import axios from 'axios';
import API_BASE_URL from './config';

// Create a dedicated axios instance
const api = axios.create({
  baseURL: API_BASE_URL,
  timeout: 90000, // 90 second timeout — handles Render free tier cold starts (can take 60s+)
  headers: {
    'Content-Type': 'application/json',
  },
});

// --- Request Interceptor ---
// Attach JWT token from localStorage to every outgoing request
api.interceptors.request.use(
  (config) => {
    const token = localStorage.getItem('authToken');
    if (token) {
      config.headers.Authorization = `Bearer ${token}`;
    }
    // Track retry count per request
    config._retryCount = config._retryCount || 0;
    return config;
  },
  (error) => {
    return Promise.reject(error);
  }
);

// --- Response Interceptor ---
// Handle 401 responses globally by clearing auth state and redirecting to login
// Handle timeout/network errors with auto-retry (for Render cold starts)
api.interceptors.response.use(
  (response) => {
    return response;
  },
  async (error) => {
    const config = error.config;

    // Auto-retry on timeout or network error (Render cold start), up to 2 retries
    const isTimeout = error.code === 'ECONNABORTED' || error.code === 'ERR_NETWORK';
    const canRetry = config && config._retryCount < 2;

    if (isTimeout && canRetry) {
      config._retryCount += 1;
      const delay = config._retryCount * 5000; // wait 5s, then 10s between retries
      await new Promise(resolve => setTimeout(resolve, delay));
      return api(config);
    }

    if (error.response && error.response.status === 401) {
      // Token expired or invalid — clear auth state
      localStorage.removeItem('authToken');
      localStorage.removeItem('isAuthenticated');
      localStorage.removeItem('loginUser');

      // Redirect to login (reload forces React to re-render from unauthenticated state)
      if (window.location.pathname !== '/') {
        window.location.href = '/';
      }
    }
    return Promise.reject(error);
  }
);

/**
 * Ping the backend to wake it up from Render free-tier sleep.
 * Call this early in the app lifecycle (e.g. on login page mount).
 */
export const warmUpBackend = async () => {
  try {
    await axios.get(`${API_BASE_URL}/health`, { timeout: 90000 });
  } catch {
    // Silently ignore — this is just a wake-up call
  }
};

export default api;
