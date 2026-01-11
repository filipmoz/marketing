"""
Pydantic models for request/response validation
"""
from pydantic import BaseModel, HttpUrl
from typing import Optional, Dict, Any
from datetime import datetime

class ResearchDataBase(BaseModel):
    url: str
    title: Optional[str] = None
    content: Optional[str] = None
    category: Optional[str] = None
    metadata: Optional[Dict[str, Any]] = None

class ResearchDataCreate(ResearchDataBase):
    """Model for creating research data"""
    pass

class ResearchData(ResearchDataBase):
    """Model for research data response"""
    id: int
    collected_at: datetime
    status: str
    
    class Config:
        from_attributes = True

class ResearchRequest(BaseModel):
    """Request model for research collection"""
    url: str
    category: Optional[str] = None
    extract_text: bool = True
    extract_links: bool = False
    extract_images: bool = False

class ExportRequest(BaseModel):
    """Request model for data export"""
    category: Optional[str] = None
    start_date: Optional[datetime] = None
    end_date: Optional[datetime] = None
    format: str = "xls"  # xls or xlsx

