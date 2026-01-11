"""
Export router for generating Excel files
"""
from fastapi import APIRouter, Depends, HTTPException, Query
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session
from sqlalchemy import desc
from typing import Optional
from datetime import datetime
import json
import io

from app.database import get_db, ResearchDataModel
from app.excel_export import ExcelExporter

router = APIRouter()
exporter = ExcelExporter()

@router.get("/excel")
async def export_to_excel(
    category: Optional[str] = None,
    status: Optional[str] = None,
    db: Session = Depends(get_db)
):
    """
    Export research data to Excel file (.xlsx)
    """
    try:
        # Query data
        query = db.query(ResearchDataModel)
        
        if category:
            query = query.filter(ResearchDataModel.category == category)
        
        if status:
            query = query.filter(ResearchDataModel.status == status)
        
        query = query.order_by(desc(ResearchDataModel.collected_at))
        data_items = query.all()
        
        # Convert to dictionary format
        data = []
        for item in data_items:
            data.append({
                'id': item.id,
                'url': item.url,
                'title': item.title,
                'content': item.content,
                'category': item.category,
                'collected_at': item.collected_at,
                'status': item.status
            })
        
        if not data:
            raise HTTPException(status_code=404, detail="No data found to export")
        
        # Generate Excel file
        excel_bytes = exporter.export_to_bytes(data)
        
        # Generate filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"research_data_{timestamp}.xlsx"
        if category:
            filename = f"research_data_{category}_{timestamp}.xlsx"
        
        return StreamingResponse(
            io.BytesIO(excel_bytes.read()),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error exporting data: {str(e)}")

@router.get("/excel/{data_id}")
async def export_single_to_excel(
    data_id: int,
    db: Session = Depends(get_db)
):
    """
    Export single research data entry to Excel
    """
    data_item = db.query(ResearchDataModel).filter(ResearchDataModel.id == data_id).first()
    
    if not data_item:
        raise HTTPException(status_code=404, detail="Data not found")
    
    data = [{
        'id': data_item.id,
        'url': data_item.url,
        'title': data_item.title,
        'content': data_item.content,
        'category': data_item.category,
        'collected_at': data_item.collected_at,
        'status': data_item.status
    }]
    
    excel_bytes = exporter.export_to_bytes(data)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"research_data_{data_id}_{timestamp}.xlsx"
    
    return StreamingResponse(
        io.BytesIO(excel_bytes.read()),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

