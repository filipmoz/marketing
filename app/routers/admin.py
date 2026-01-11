"""
Admin router for viewing and managing research data
"""
from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session
from sqlalchemy import desc
from typing import List, Optional
from datetime import datetime
import json

from app.database import get_db, ResearchDataModel
from app.models import ResearchData

router = APIRouter()

@router.get("/data", response_model=List[ResearchData])
async def get_all_data(
    skip: int = Query(0, ge=0),
    limit: int = Query(100, ge=1, le=1000),
    category: Optional[str] = None,
    status: Optional[str] = None,
    db: Session = Depends(get_db)
):
    """
    Get all research data with optional filtering
    """
    query = db.query(ResearchDataModel)
    
    if category:
        query = query.filter(ResearchDataModel.category == category)
    
    if status:
        query = query.filter(ResearchDataModel.status == status)
    
    query = query.order_by(desc(ResearchDataModel.collected_at))
    data = query.offset(skip).limit(limit).all()
    
    results = []
    for item in data:
        results.append(ResearchData(
            id=item.id,
            url=item.url,
            title=item.title,
            content=item.content,
            category=item.category,
            metadata=json.loads(item.extra_data) if item.extra_data else None,
            collected_at=item.collected_at,
            status=item.status
        ))
    
    return results

@router.get("/data/{data_id}", response_model=ResearchData)
async def get_data_by_id(
    data_id: int,
    db: Session = Depends(get_db)
):
    """
    Get specific research data by ID
    """
    data = db.query(ResearchDataModel).filter(ResearchDataModel.id == data_id).first()
    
    if not data:
        raise HTTPException(status_code=404, detail="Data not found")
    
    return ResearchData(
        id=data.id,
        url=data.url,
        title=data.title,
        content=data.content,
        category=data.category,
        metadata=json.loads(data.extra_data) if data.extra_data else None,
        collected_at=data.collected_at,
        status=data.status
    )

@router.get("/stats")
async def get_stats(db: Session = Depends(get_db)):
    """
    Get statistics about collected data
    """
    total = db.query(ResearchDataModel).count()
    collected = db.query(ResearchDataModel).filter(ResearchDataModel.status == "collected").count()
    errors = db.query(ResearchDataModel).filter(ResearchDataModel.status == "error").count()
    
    # Get categories
    categories = db.query(ResearchDataModel.category).distinct().all()
    category_list = [cat[0] for cat in categories if cat[0]]
    
    return {
        "total": total,
        "collected": collected,
        "errors": errors,
        "categories": category_list,
        "categories_count": len(category_list)
    }

@router.delete("/data/{data_id}")
async def delete_data(
    data_id: int,
    db: Session = Depends(get_db)
):
    """
    Delete research data by ID
    """
    data = db.query(ResearchDataModel).filter(ResearchDataModel.id == data_id).first()
    
    if not data:
        raise HTTPException(status_code=404, detail="Data not found")
    
    db.delete(data)
    db.commit()
    
    return {"message": "Data deleted successfully", "id": data_id}

