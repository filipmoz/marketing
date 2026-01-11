"""
Export router for survey data to Excel
"""
from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session
from typing import Optional
from datetime import datetime
import io

from app.database import get_db
from app.survey_models import SurveyResponse
from app.survey_excel_export import SurveyExcelExporter
from app.auth import get_current_admin

router = APIRouter()
exporter = SurveyExcelExporter()

@router.get("/excel")
async def export_survey_to_excel(
    db: Session = Depends(get_db),
    admin: dict = Depends(get_current_admin)
):
    """
    Export all survey responses to Excel file (.xlsx) in code book format
    Ready for statistical analysis (crosstabs, pivot tables, etc.)
    """
    try:
        # Get all survey responses
        responses = db.query(SurveyResponse).order_by(SurveyResponse.submitted_at).all()
        
        if not responses:
            raise HTTPException(status_code=404, detail="No survey responses found to export")
        
        # Generate Excel file
        excel_bytes = exporter.export_survey_data(responses)
        
        # Generate filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"survey_data_codebook_{timestamp}.xlsx"
        
        return StreamingResponse(
            io.BytesIO(excel_bytes.read()),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error exporting survey data: {str(e)}")

