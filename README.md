# Quantitative Assessment - Survey Data Collection System

A comprehensive web application for collecting survey data on consumer attitudes towards fuel prices, global warming, and alternative fuels. Designed for quantitative research assessment with Excel export functionality for statistical analysis.

## Research Topic

A prominent car manufacturer is seeking to understand consumer attitudes towards fuel prices, global warming, and alternative fuels. With the increasing concern over environmental issues and rising petrol costs, the company aims to gather insights that can guide their future vehicle development and marketing strategies.

**Research Question:** How do consumer perceptions of global warming, petrol usage, and fuel prices influence their preferences for alternative fuel vehicles?

## Features

- 📝 **Survey Form**: Complete survey interface for manual data entry
- 📊 **Admin Interface**: Tutor/admin interface to view and edit all survey responses
- ✏️ **Editable Demographics**: Inline editing of Gender, Marital Status, and Age Category
- 📥 **Excel Export**: Generate advanced Excel files with charts, statistics, and analysis templates
- 🗄️ **Database Storage**: SQLite database for persistent data storage
- 📈 **Statistics Dashboard**: View statistics about collected responses
- 🔍 **Code Book**: Automatic code book generation in Excel
- 📊 **Charts & Visualizations**: 5 charts included in Excel export
- 📋 **Analysis Templates**: Ready-to-use templates for crosstabs, t-tests, ANOVA, chi-square tests

## Survey Structure

### Attitude Questions (1-7 Likert Scale)
1. I am worried about global warming
2. Global warming is a real threat
3. British use too much Petrol
4. We should be looking for Petrol substitutes
5. Petrol prices are too high now
6. High gasoline prices will impact what type of cars are purchased

### Personality Types (1-7 Scale)
1. Novelist - very early adopter, risk taker
2. Innovator - early adopter, less risk taker
3. Trendsetter - opinion leaders
4. Forerunner - early majority
5. Mainstreamer - late majority
6. Classic - laggards

### Demographics
- Gender (Male, Female)
- Marital Status (Unmarried, Married)
- Age Category (18 to 34, 35 to 65, 65 and older)

## Installation

### Quick Setup (Recommended)

1. **Run the setup script:**
   ```bash
   chmod +x setup.sh
   ./setup.sh
   ```

2. **Run the application:**
   ```bash
   chmod +x run.sh
   ./run.sh
   ```

### Manual Setup

1. **Create a virtual environment:**
   ```bash
   python -m venv venv
   ```

2. **Activate the virtual environment:**
   ```bash
   # On Linux/Mac:
   source venv/bin/activate
   
   # On Windows:
   venv\Scripts\activate
   ```

3. **Install Python dependencies:**
   ```bash
   pip install --upgrade pip
   pip install -r requirements.txt
   ```

4. **Run the application:**
   ```bash
   python run.py
   ```

5. **Access the application:**
   - Survey Form: http://localhost:8000/
   - Admin Interface: http://localhost:8000/admin
   - API Documentation: http://localhost:8000/docs

## Usage

### For Survey Respondents

1. Navigate to http://localhost:8000/
2. Complete all survey questions
3. Submit the form

### For Tutors/Admins

1. Navigate to http://localhost:8000/admin
2. View all survey responses in the data table
3. Check statistics dashboard
4. Click "Export to Excel (Code Book)" to download data for statistical analysis

The Excel file includes:
- **Survey Data Sheet**: All responses in code book format
- **Code Book Sheet**: Variable descriptions and coding schemes

## Excel Export Format

The exported Excel file is formatted for statistical analysis and includes:

- All survey responses with proper variable names
- Code book sheet explaining variable codes
- Ready for:
  - Crosstabs (pivot tables)
  - Chi-square tests
  - T-tests
  - ANOVA
  - Other statistical analyses

## API Endpoints

### Survey Endpoints
- `POST /api/survey/submit` - Submit a survey response
- `GET /api/survey/responses` - Get all survey responses (admin)
- `GET /api/survey/stats` - Get statistics about responses
- `GET /api/survey/export/excel` - Export all data to Excel

## Project Structure

```
.
├── app/
│   ├── __init__.py
│   ├── main.py                  # FastAPI application
│   ├── database.py              # Database configuration
│   ├── survey_models.py         # Survey data models
│   ├── survey_schemas.py        # Pydantic validation schemas
│   ├── survey_excel_export.py   # Excel export for survey data
│   ├── routers/
│   │   ├── __init__.py
│   │   ├── survey.py            # Survey endpoints
│   │   └── survey_export.py     # Export endpoints
│   └── templates/
│       ├── survey_form.html     # Survey form interface
│       └── admin_survey.html    # Admin interface
├── requirements.txt
├── setup.sh
├── run.sh
└── README.md
```

## Notes

- Uses Python's built-in `html.parser` (no lxml required) for better compatibility with Python 3.13+
- Excel files are generated in .xlsx format (compatible with Excel)
- Database is SQLite by default (can be changed via DATABASE_URL environment variable)
- All survey fields are validated according to the assessment requirements

## License

This project is for educational and research purposes.
