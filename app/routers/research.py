"""
Research router for data collection endpoints
"""
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from typing import List
from datetime import datetime
import json

from app.database import get_db, ResearchDataModel
from app.models import ResearchRequest, ResearchData, ResearchDataCreate
from app.research_service import ResearchService

router = APIRouter()
research_service = ResearchService()

@router.post("/collect", response_model=ResearchData)
async def collect_data(
    request: ResearchRequest,
    db: Session = Depends(get_db)
):
    """
    Collect data from a website URL
    """
    try:
        # Collect data using research service
        collected_data = research_service.collect_data(
            url=request.url,
            category=request.category,
            extract_text=request.extract_text,
            extract_links=request.extract_links,
            extract_images=request.extract_images
        )
        
        # Save to database
        db_data = ResearchDataModel(
            url=collected_data['url'],
            title=collected_data.get('title'),
            content=collected_data.get('content'),
            category=collected_data.get('category'),
            extra_data=json.dumps(collected_data.get('metadata', {})),
            status=collected_data.get('status', 'collected')
        )
        
        db.add(db_data)
        db.commit()
        db.refresh(db_data)
        
        # Convert to response model
        result = ResearchData(
            id=db_data.id,
            url=db_data.url,
            title=db_data.title,
            content=db_data.content,
            category=db_data.category,
            metadata=json.loads(db_data.extra_data) if db_data.extra_data else None,
            collected_at=db_data.collected_at,
            status=db_data.status
        )
        
        return result
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error collecting data: {str(e)}")

@router.post("/collect-batch")
async def collect_batch(
    urls: List[str],
    category: str = None,
    db: Session = Depends(get_db)
):
    """
    Collect data from multiple URLs
    """
    results = []
    for url in urls:
        try:
            collected_data = research_service.collect_data(
                url=url,
                category=category
            )
            
            # Save to database
            db_data = ResearchDataModel(
                url=collected_data['url'],
                title=collected_data.get('title'),
                content=collected_data.get('content'),
                category=collected_data.get('category'),
                extra_data=json.dumps(collected_data.get('metadata', {})),
                status=collected_data.get('status', 'collected')
            )
            
            db.add(db_data)
            db.commit()
            db.refresh(db_data)
            
            results.append({
                "id": db_data.id,
                "url": db_data.url,
                "status": db_data.status
            })
            
        except Exception as e:
            results.append({
                "url": url,
                "status": "error",
                "error": str(e)
            })
    
    return {"collected": len(results), "results": results}

