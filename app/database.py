"""
Database configuration and session management
"""
from sqlalchemy import create_engine, Column, Integer, String, DateTime, Text, Float
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, Session
from datetime import datetime
import os

# Database URL
# For Docker: use /app/data directory for persistence
DB_DIR = os.getenv("DB_DIR", "./data")
os.makedirs(DB_DIR, exist_ok=True)
DATABASE_URL = os.getenv("DATABASE_URL", f"sqlite:///{os.path.join(DB_DIR, 'research_data.db')}")

engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

class ResearchDataModel(Base):
    """Database model for research data"""
    __tablename__ = "research_data"
    
    id = Column(Integer, primary_key=True, index=True)
    url = Column(String, index=True)
    title = Column(String, nullable=True)
    content = Column(Text, nullable=True)
    collected_at = Column(DateTime, default=datetime.utcnow)
    extra_data = Column(Text, nullable=True)  # JSON string for additional data (renamed from metadata to avoid SQLAlchemy conflict)
    category = Column(String, nullable=True)
    status = Column(String, default="collected")

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

