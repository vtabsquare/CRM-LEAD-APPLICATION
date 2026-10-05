// API Configuration
// Uses environment variable in production, falls back to localhost in development
// Supports both VITE_API_BASE_URL and VITE_API_URL (Render uses VITE_API_URL)
const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || import.meta.env.VITE_API_URL || 'http://localhost:8000';

export default API_BASE_URL;
