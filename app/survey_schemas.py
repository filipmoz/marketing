"""
Pydantic schemas for survey data validation
"""
from pydantic import BaseModel, Field
from typing import Optional
from datetime import datetime

class SurveyResponseCreate(BaseModel):
    """Schema for creating a survey response"""
    # Attitude Questions (1-7 Likert scale)
    q1_worried_global_warming: int = Field(..., ge=1, le=7, description="I am worried about global warming")
    q2_global_warming_threat: int = Field(..., ge=1, le=7, description="Global warming is a real threat")
    q3_british_use_too_much_petrol: int = Field(..., ge=1, le=7, description="British use too much Petrol")
    q4_look_petrol_substitutes: int = Field(..., ge=1, le=7, description="We should be looking for Petrol substitutes")
    q5_petrol_prices_too_high: int = Field(..., ge=1, le=7, description="Petrol prices are too high now")
    q6_high_prices_impact_cars: int = Field(..., ge=1, le=7, description="High gasoline prices will impact what type of cars are purchased")
    
    # Personality Types (1-7 scale)
    personality_novelist: int = Field(..., ge=1, le=7)
    personality_innovator: int = Field(..., ge=1, le=7)
    personality_trendsetter: int = Field(..., ge=1, le=7)
    personality_forerunner: int = Field(..., ge=1, le=7)
    personality_mainstreamer: int = Field(..., ge=1, le=7)
    personality_classic: int = Field(..., ge=1, le=7)
    
    # Demographics
    gender: str = Field(..., pattern="^(Male|Female)$")
    marital_status: str = Field(..., pattern="^(Unmarried|Married)$")
    age_category: str = Field(..., pattern="^(18 to 34|35 to 65|65 and older)$")

class SurveyResponse(BaseModel):
    """Schema for survey response"""
    id: int
    submitted_at: datetime
    q1_worried_global_warming: int
    q2_global_warming_threat: int
    q3_british_use_too_much_petrol: int
    q4_look_petrol_substitutes: int
    q5_petrol_prices_too_high: int
    q6_high_prices_impact_cars: int
    personality_novelist: int
    personality_innovator: int
    personality_trendsetter: int
    personality_forerunner: int
    personality_mainstreamer: int
    personality_classic: int
    gender: str
    marital_status: str
    age_category: str
    
    class Config:
        from_attributes = True

