"""
Main FastAPI application for Research Data Collection and Admin Interface
"""
from fastapi import FastAPI, HTTPException, Depends, Request
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from starlette.templating import Jinja2Templates
from typing import List, Optional
from datetime import datetime
import os

from app.database import get_db, init_db
from app.routers import survey
from starlette.middleware.sessions import SessionMiddleware
import secrets

app = FastAPI(
    title="Quantitative Assessment - Survey Data Collection",
    description="Survey system for collecting consumer attitudes towards fuel prices, global warming, and alternative fuels",
    version="1.0.0"
)

# Add session middleware for authentication
app.add_middleware(SessionMiddleware, secret_key=secrets.token_urlsafe(32))

# Initialize database
init_db()

# Include routers
app.include_router(survey.router, prefix="/api/survey", tags=["Survey"])

# Templates
templates = Jinja2Templates(directory="app/templates")

@app.get("/", response_class=HTMLResponse)
async def root(request: Request):
    """Survey form interface (requires authentication)"""
    from app.auth import get_current_survey_user
    try:
        # Check authentication
        await get_current_survey_user(request)
        return templates.TemplateResponse("survey_form.html", {"request": request})
    except HTTPException:
        # Redirect to login if not authenticated
        return RedirectResponse(url="/login", status_code=302)

@app.get("/login", response_class=HTMLResponse)
async def survey_login_page(request: Request):
    """
    Survey login page
    """
    return templates.TemplateResponse("survey_login.html", {"request": request})

@app.post("/login")
async def survey_login(request: Request):
    """
    Survey login endpoint
    """
    from app.auth import verify_survey_password, create_session
    
    # Get form data
    form_data = await request.form()
    username = form_data.get("username", "")
    password = form_data.get("password", "")
    
    if verify_survey_password(username, password):
        # Create session
        session_token = create_session()
        request.session["survey_session"] = session_token
        return RedirectResponse(url="/", status_code=302)
    else:
        return HTMLResponse(
            content="""
            <html>
                <body>
                    <h2>Login Failed</h2>
                    <p>Invalid username or password.</p>
                    <a href="/login">Try again</a>
                </body>
            </html>
            """,
            status_code=401
        )

@app.get("/logout")
async def survey_logout(request: Request):
    """
    Survey logout endpoint
    """
    from app.auth import remove_session
    
    session_token = request.session.get("survey_session")
    if session_token:
        remove_session(session_token)
        request.session.clear()
    
    return RedirectResponse(url="/login", status_code=302)

@app.get("/admin", response_class=HTMLResponse)
async def admin_interface(request: Request):
    """
    Admin interface for viewing survey responses (requires authentication)
    """
    from app.auth import get_current_admin
    try:
        # Check authentication
        await get_current_admin(request)
        return templates.TemplateResponse("admin_survey.html", {"request": request})
    except HTTPException:
        # Redirect to login if not authenticated
        return RedirectResponse(url="/admin/login", status_code=302)

@app.get("/admin/login", response_class=HTMLResponse)
async def admin_login_page(request: Request):
    """
    Admin login page
    """
    return templates.TemplateResponse("admin_login.html", {"request": request})

@app.post("/admin/login")
async def admin_login(request: Request):
    """
    Admin login endpoint
    """
    from app.auth import verify_password, create_session
    
    # Get form data
    form_data = await request.form()
    username = form_data.get("username", "")
    password = form_data.get("password", "")
    
    if verify_password(username, password):
        # Create session
        session_token = create_session()
        request.session["admin_session"] = session_token
        return RedirectResponse(url="/admin", status_code=302)
    else:
        return HTMLResponse(
            content="""
            <html>
                <body>
                    <h2>Login Failed</h2>
                    <p>Invalid username or password.</p>
                    <a href="/admin/login">Try again</a>
                </body>
            </html>
            """,
            status_code=401
        )

@app.get("/admin/logout")
async def admin_logout(request: Request):
    """
    Admin logout endpoint
    """
    from app.auth import remove_session
    
    session_token = request.session.get("admin_session")
    if session_token:
        remove_session(session_token)
        request.session.clear()
    
    return RedirectResponse(url="/admin/login", status_code=302)

@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", "8000"))
    uvicorn.run(app, host="0.0.0.0", port=port)

