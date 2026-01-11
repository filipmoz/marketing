"""
Survey data models for quantitative assessment
Based on the car manufacturer research survey
"""
from sqlalchemy import Column, Integer, String, DateTime, Float
from sqlalchemy.ext.declarative import declarative_base
from datetime import datetime
from app.database import Base

class SurveyResponse(Base):
    """Survey response model matching the assessment requirements"""
    __tablename__ = "survey_responses"
    
    # Primary key
    id = Column(Integer, primary_key=True, index=True)
    submitted_at = Column(DateTime, default=datetime.utcnow)
    
    # Attitude Questions (Likert Scale: 1=Very strongly disagree, 7=Very strongly agree)
    q1_worried_global_warming = Column(Integer)  # I am worried about global warming
    q2_global_warming_threat = Column(Integer)   # Global warming is a real threat
    q3_british_use_too_much_petrol = Column(Integer)  # British use too much Petrol
    q4_look_petrol_substitutes = Column(Integer)  # We should be looking for Petrol substitutes
    q5_petrol_prices_too_high = Column(Integer)   # Petrol prices are too high now
    q6_high_prices_impact_cars = Column(Integer)  # High gasoline prices will impact what type of cars are purchased
    
    # Personality Types (Scale: 1=does not describe me at all, 7=describes me perfectly)
    personality_novelist = Column(Integer)        # Novelist - very early adopter, risk taker
    personality_innovator = Column(Integer)       # Innovator - early adopter, less risk taker
    personality_trendsetter = Column(Integer)      # Trendsetter - opinion leaders
    personality_forerunner = Column(Integer)      # Forerunner - early majority
    personality_mainstreamer = Column(Integer)    # Mainstreamer - late majority
    personality_classic = Column(Integer)         # Classic - laggards
    
    # Demographics
    gender = Column(String)                       # Male, Female
    marital_status = Column(String)               # Unmarried, Married
    age_category = Column(String)                 # 18 to 34, 35 to 65, 65 and older
    
    def to_dict(self):
        """Convert to dictionary for Excel export"""
        return {
            'ID': self.id,
            'Submitted At': self.submitted_at.isoformat() if self.submitted_at else '',
            # Attitude questions
            'Q1_Worried_Global_Warming': self.q1_worried_global_warming,
            'Q2_Global_Warming_Threat': self.q2_global_warming_threat,
            'Q3_British_Use_Too_Much_Petrol': self.q3_british_use_too_much_petrol,
            'Q4_Look_Petrol_Substitutes': self.q4_look_petrol_substitutes,
            'Q5_Petrol_Prices_Too_High': self.q5_petrol_prices_too_high,
            'Q6_High_Prices_Impact_Cars': self.q6_high_prices_impact_cars,
            # Personality types
            'Personality_Novelist': self.personality_novelist,
            'Personality_Innovator': self.personality_innovator,
            'Personality_Trendsetter': self.personality_trendsetter,
            'Personality_Forerunner': self.personality_forerunner,
            'Personality_Mainstreamer': self.personality_mainstreamer,
            'Personality_Classic': self.personality_classic,
            # Demographics
            'Gender': self.gender,
            'Marital_Status': self.marital_status,
            'Age_Category': self.age_category,
        }

