"""
Simple authentication for admin interface
"""
from fastapi import HTTPException, Depends, status
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi import Request
from starlette.middleware.sessions import SessionMiddleware
import secrets
import os
import hashlib

# Simple session storage (in production, use Redis or database)
active_sessions = set()

# Admin credentials (in production, use environment variables and hashed passwords)
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "admin123")  # Change this!

# Survey credentials
SURVEY_USERNAME = os.getenv("SURVEY_USERNAME", "Survey")
SURVEY_PASSWORD = os.getenv("SURVEY_PASSWORD", "Filip")

def verify_password(username: str, password: str) -> bool:
    """Verify admin credentials"""
    return username == ADMIN_USERNAME and password == ADMIN_PASSWORD

def verify_survey_password(username: str, password: str) -> bool:
    """Verify survey credentials"""
    return username == SURVEY_USERNAME and password == SURVEY_PASSWORD

def create_session() -> str:
    """Create a new session token"""
    session_token = secrets.token_urlsafe(32)
    active_sessions.add(session_token)
    return session_token

def verify_session(session_token: str) -> bool:
    """Verify if session token is valid"""
    return session_token in active_sessions

def remove_session(session_token: str):
    """Remove a session token"""
    active_sessions.discard(session_token)

async def get_current_admin(request: Request):
    """Dependency to check if user is authenticated as admin"""
    session_token = request.session.get("admin_session")
    
    if not session_token or not verify_session(session_token):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Not authenticated. Please log in.",
            headers={"WWW-Authenticate": "Bearer"},
        )
    
    return {"username": ADMIN_USERNAME, "session": session_token}

async def get_current_survey_user(request: Request):
    """Dependency to check if user is authenticated for survey"""
    session_token = request.session.get("survey_session")
    
    if not session_token or not verify_session(session_token):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Not authenticated. Please log in.",
            headers={"WWW-Authenticate": "Bearer"},
        )
    
    return {"username": SURVEY_USERNAME, "session": session_token}

