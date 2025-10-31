"""
Database configuration and session management
"""
from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, Session
import os

# Database URL
# For Docker: use /app/data directory for persistence
DB_DIR = os.getenv("DB_DIR", "./data")
os.makedirs(DB_DIR, exist_ok=True)
DATABASE_URL = os.getenv("DATABASE_URL", f"sqlite:///{os.path.join(DB_DIR, 'survey_data.db')}")

engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

def init_db():
    """Initialize database tables"""
    # Import survey models to ensure they're registered
    from app.survey_models import SurveyResponse
    Base.metadata.create_all(bind=engine)

def get_db():
    """Dependency for getting database session"""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

