"""
Security Module for CRM Lead Application
Provides JWT authentication, password hashing, login attempt tracking,
and security middleware for FastAPI.
"""

import os
import time
import hashlib
import hmac
import secrets
from datetime import datetime, timedelta, timezone
from typing import Optional, Dict, Any
from functools import wraps

import jwt
from fastapi import Request, HTTPException, status
from fastapi.responses import JSONResponse
from starlette.middleware.base import BaseHTTPMiddleware

from dotenv import load_dotenv
load_dotenv()
load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))

# --- Configuration ---
JWT_SECRET_KEY = os.getenv("JWT_SECRET_KEY", secrets.token_hex(32))
JWT_ALGORITHM = "HS256"
JWT_EXPIRY_MINUTES = int(os.getenv("JWT_EXPIRY_MINUTES", "60"))

ALLOWED_ORIGINS = [
    origin.strip()
    for origin in os.getenv(
        "ALLOWED_ORIGINS",
        "http://localhost:3000,http://localhost:5173,http://localhost:5174,https://crm-lead-application-5apw.onrender.com"
    ).split(",")
    if origin.strip()
]
# Ensure production frontend URL is always allowed
if "https://crm-lead-application-5apw.onrender.com" not in ALLOWED_ORIGINS:
    ALLOWED_ORIGINS.append("https://crm-lead-application-5apw.onrender.com")


def _get_cors_headers(request: Request) -> dict:
    """Return CORS headers if the request Origin is in the allowed list."""
    origin = request.headers.get("origin", "")
    if origin in ALLOWED_ORIGINS:
        return {
            "Access-Control-Allow-Origin": origin,
            "Access-Control-Allow-Credentials": "true",
        }
    return {}

# Paths that don't require authentication
PUBLIC_PATHS = {
    "/",
    "/login",
    "/health",
    "/docs",
    "/redoc",
    "/openapi.json",
}

# Prefixes that are public (for FastAPI internal routes)
PUBLIC_PREFIXES = (
    "/docs",
    "/redoc",
    "/openapi",
)


# --- JWT Token Functions ---

def create_access_token(username: str, role: str = "user", extra: dict = None) -> str:
    """Create a JWT access token."""
    now = datetime.now(timezone.utc)
    payload = {
        "sub": username,
        "role": role,
        "iat": now,
        "exp": now + timedelta(minutes=JWT_EXPIRY_MINUTES),
        "jti": secrets.token_hex(16),  # Unique token ID
    }
    if extra:
        payload.update(extra)
    return jwt.encode(payload, JWT_SECRET_KEY, algorithm=JWT_ALGORITHM)


def verify_access_token(token: str) -> Dict[str, Any]:
    """Verify and decode a JWT access token. Raises HTTPException on failure."""
    try:
        payload = jwt.decode(token, JWT_SECRET_KEY, algorithms=[JWT_ALGORITHM])
        return payload
    except jwt.ExpiredSignatureError:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Token has expired",
            headers={"WWW-Authenticate": "Bearer"},
        )
    except jwt.InvalidTokenError:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid token",
            headers={"WWW-Authenticate": "Bearer"},
        )


# --- Password Hashing (bcrypt-compatible via hashlib fallback) ---
# Using PBKDF2-SHA256 so we don't require a C compiler for bcrypt on all platforms.

HASH_ITERATIONS = 260000  # OWASP recommended for PBKDF2-SHA256

def hash_password(password: str) -> str:
    """Hash a password using PBKDF2-SHA256 with a random salt."""
    salt = secrets.token_hex(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), HASH_ITERATIONS)
    return f"pbkdf2:sha256:{HASH_ITERATIONS}${salt}${dk.hex()}"


def verify_password(password: str, hashed: str) -> bool:
    """Verify a password against a PBKDF2-SHA256 hash."""
    try:
        if hashed.startswith("pbkdf2:sha256:"):
            # Our format: pbkdf2:sha256:<iterations>$<salt>$<hash>
            parts = hashed.split("$")
            if len(parts) != 3:
                return False
            iterations = int(parts[0].split(":")[-1])
            salt = parts[1]
            expected_hash = parts[2]
            dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), iterations)
            return hmac.compare_digest(dk.hex(), expected_hash)
        else:
            # Fallback: plaintext comparison for legacy passwords (transition period)
            return hmac.compare_digest(password, hashed)
    except Exception:
        return False


# --- Login Attempt Tracking ---

# In-memory store: {username: {"count": int, "locked_until": float}}
_login_attempts: Dict[str, Dict[str, Any]] = {}
MAX_LOGIN_ATTEMPTS = 5
LOCKOUT_DURATION_SECONDS = 900  # 15 minutes


def check_login_attempts(username: str) -> None:
    """Check if the user is locked out. Raises HTTPException if locked."""
    record = _login_attempts.get(username)
    if record and record.get("locked_until", 0) > time.time():
        remaining = int(record["locked_until"] - time.time())
        raise HTTPException(
            status_code=status.HTTP_429_TOO_MANY_REQUESTS,
            detail=f"Account locked due to too many failed attempts. Try again in {remaining} seconds.",
        )


def record_failed_attempt(username: str) -> None:
    """Record a failed login attempt. Lock account after MAX_LOGIN_ATTEMPTS."""
    if username not in _login_attempts:
        _login_attempts[username] = {"count": 0, "locked_until": 0}

    _login_attempts[username]["count"] += 1

    if _login_attempts[username]["count"] >= MAX_LOGIN_ATTEMPTS:
        _login_attempts[username]["locked_until"] = time.time() + LOCKOUT_DURATION_SECONDS


def record_successful_login(username: str) -> None:
    """Reset login attempt counter on successful login."""
    _login_attempts.pop(username, None)


# --- Password Validation ---

# Top common passwords to reject
COMMON_PASSWORDS = {
    "password", "123456", "12345678", "qwerty", "abc123", "monkey", "1234567",
    "letmein", "trustno1", "dragon", "baseball", "iloveyou", "master", "sunshine",
    "ashley", "bailey", "shadow", "123123", "654321", "superman", "qazwsx",
    "michael", "football", "password1", "password123", "admin", "admin123",
    "root", "toor", "pass", "test", "guest", "changeme", "welcome",
}


def validate_password(password: str) -> tuple[bool, str]:
    """
    Validate password against security requirements.
    Returns (is_valid, error_message).
    """
    if len(password) < 12:
        return False, "Password must be at least 12 characters long"
    if len(password) > 128:
        return False, "Password must not exceed 128 characters"
    if not any(c.islower() for c in password):
        return False, "Password must contain at least one lowercase letter"
    if not any(c.isupper() for c in password):
        return False, "Password must contain at least one uppercase letter"
    if not any(c.isdigit() for c in password):
        return False, "Password must contain at least one digit"
    if not any(c in "!@#$%^&*()-_=+[]{}|;:'\",.<>?/`~" for c in password):
        return False, "Password must contain at least one special character"
    if password.lower() in COMMON_PASSWORDS:
        return False, "This password is too common. Please choose a stronger password."
    return True, ""


# --- Authentication Middleware ---

class AuthenticationMiddleware(BaseHTTPMiddleware):
    """
    Middleware that checks for a valid JWT token on all requests
    except those to public paths. This avoids modifying every route handler.
    """

    async def dispatch(self, request: Request, call_next):
        path = request.url.path.rstrip("/") or "/"

        # Allow public paths
        if path in PUBLIC_PATHS or any(path.startswith(p) for p in PUBLIC_PREFIXES):
            return await call_next(request)

        # Allow OPTIONS (CORS preflight)
        if request.method == "OPTIONS":
            return await call_next(request)

        # Extract token from Authorization header
        auth_header = request.headers.get("Authorization", "")
        if not auth_header.startswith("Bearer "):
            return JSONResponse(
                status_code=status.HTTP_401_UNAUTHORIZED,
                content={"detail": "Authentication required"},
                headers={"WWW-Authenticate": "Bearer", **_get_cors_headers(request)},
            )

        token = auth_header[7:]  # Strip "Bearer "
        try:
            payload = verify_access_token(token)
            # Store user info on request state for route handlers to access if needed
            request.state.user = payload
        except HTTPException as e:
            return JSONResponse(
                status_code=e.status_code,
                content={"detail": e.detail},
                headers={"WWW-Authenticate": "Bearer", **_get_cors_headers(request)},
            )

        return await call_next(request)


# --- Security Headers Middleware ---

class SecurityHeadersMiddleware(BaseHTTPMiddleware):
    """Add security headers to all responses."""

    async def dispatch(self, request: Request, call_next):
        response = await call_next(request)

        # Do not override CORS preflight responses — let CORSMiddleware handle them cleanly
        if request.method == "OPTIONS":
            return response

        # Prevent MIME type sniffing
        response.headers["X-Content-Type-Options"] = "nosniff"

        # Prevent clickjacking
        response.headers["X-Frame-Options"] = "DENY"

        # Control referrer information
        response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"

        # HSTS — enforce HTTPS (only effective when served over HTTPS)
        response.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"

        # Permissions Policy — restrict browser features
        response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"

        # Remove server identification headers
        if "X-Powered-By" in response.headers:
            del response.headers["X-Powered-By"]
        if "Server" in response.headers:
            del response.headers["Server"]

        # Content Security Policy
        response.headers["Content-Security-Policy"] = (
            "default-src 'self'; "
            "script-src 'self' 'unsafe-inline' 'unsafe-eval'; "
            "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
            "font-src 'self' https://fonts.gstatic.com; "
            "img-src 'self' data: blob:; "
            "connect-src 'self' " + " ".join(ALLOWED_ORIGINS) + " https://*.googleapis.com https://*.onrender.com; "
            "frame-src 'self' https://app.powerbi.com; "
            "frame-ancestors 'none';"
        )

        return response

