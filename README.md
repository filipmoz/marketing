# Quantitative Assessment - Survey Data Collection

This web aplication is designed to gather survey data on consumer attitudes towards fuel prices, global warming, and alternative fuels. It allows for exporting data to Excel for statistical analysis and is developed for a quantative research assignment.

## Research Topic

A car manufacturer aims to explore consumer attitudes towards fuel prices, global warming, and alternative energy sources. Our data collection will inform their vehicle and marketing strategies. The research question is: How do perceptions of global warming, petrol usage, and fuel prices influence preferences for alternative fuel vehicles?

## Features

- Survey form for manual data entry
- Admin Interface: Enables viewing and editing of all responses (for tutors/admins)
- Inline editable demographics: Gender, Marital Status, Age
- Excel export with charts, code book, and analysis templates (crosstabs, t-tests, ANOVA, chi-square)
- Uses SQLite for data storage (no additional database setup required)
- Statistics dashboard to display response counts, etc.

## Survey Structure

### Attitude Questions (1–7 Likert Scale)

1. Concern about global warming
2. Global warming is perceived as a genuine threat
3. Excessive petrol consumption in Britain
4. The need to seek petrol alternatives
5. Petrol prices are perceived as too high
6. The impact of high gasoline prices on car purchases

### Personality Types (1–7 Scale)

Categories include Novelist, Innovator, Trendsetter, Forerunner, Mainstreamer, and Classic (ranging from early adopters to laggards).

### Demographics

Gender (Male/Female), Marital Status (Unmarried/Married), Age Groups (18–34, 35–65, 65+).

## Instalation

For a quick setup, run `./setup.sh` followed by `./run.sh` (use chmod +x if necessary).

Manual installation steps:

```bash
python -m venv venv
source venv/bin/activate   # or venv\Scripts\activate on Windows
pip install -r requirements.txt
python run.py
```

Access the survey at http://localhost:8000/, the admin interface at http://localhost:8000/admin, and API documentation at /docs.

## Usage

Respondents visit the survey URL and submit their responses. Administrators can access the /admin section to view response tables, check statistics, and export data to Excel in a code book format. The exported file includes a survey data sheet with variable descriptions and is set up for crosstabs, chi-square, t-tests, and ANOVA.

## API Endpoints

- POST /api/survey/submit – Submit a response
- GET /api/survey/responses – Retrieve all responses (admin)
- GET /api/survey/stats – Access survey statistics
- GET /api/survey/export/excel – Download Excel file

## Project Layout

```
app/
  main.py, database.py, survey_models.py, survey_schemas.py
  survey_excel_export.py
  routers/survey.py, survey_export.py
  templates/ (survey_form.html, admin_survey.html, etc.)
requirements.txt, setup.sh, run.sh, run.py
```

The aplication uses SQLite by default. If you wish to change the database, modify the DATABASE_URL enviornment variable. The Excel files are in .xlsx format, and validation adheres to the assessment specifications.

## License

This aplication is intended for educational or research purposes.

**Student ID**: 22002216

*Note: GitLab Copilot was used solely for code alignment and autocomplete.*
