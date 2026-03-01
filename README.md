

I have scraped University Courses Aggrigator website called MastersPortal.com
Tool: Playwright (headless Chromium)                           

Steps to Run :
create a venv
then install the requirements.txt



Run:
# Create virtual environment
python -m venv venv

# Activate virtual environment (Windows)
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the scraper
py scraper.py

Output:
    university_course_data.xlsx  (2 sheets: Universities + Courses)

