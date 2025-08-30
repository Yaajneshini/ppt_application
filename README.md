# Powerpoint Generator

This repository contains a Flask-based web application developed by Yaajneshini.  
The application converts raw text into well-structured PowerPoint presentations with the help of a Large Language Model (LLM).  

## Features

- Accepts raw text input (with optional guidance) from the user.  
- Processes the text using an LLM API to extract the most relevant points.  
- Produces a structured JSON outline of slides:
  - The first slide contains only a title.  
  - Subsequent slides include a title and three to five concise bullet points.  
- Displays a preview of the generated slides before creating the presentation.  
- Allows users to upload a PowerPoint template (`.pptx`) to maintain a consistent design.  
- Generates a PowerPoint file with:
  - A title slide featuring a bold heading.  
  - Content slides with subtitles and bullet points.  
- Provides the generated presentation as a downloadable `generated.pptx` file.  
- Includes error handling for invalid JSON or API issues.  
- Configured with CORS support to allow integration with external front-end applications.  
- Runs on port 8080 and can be hosted publicly.  

## Tech Stack

- **Backend Framework:** Flask (Python)  
- **Presentation Generation:** python-pptx  
- **API Integration:** Requests library for LLM API calls  
- **Frontend Rendering:** Jinja2 templates (HTML)  
- **Cross-Origin Access:** Flask-CORS  
- **Utilities:** datetime, io, os, json  

## How to Run

1. Install the dependencies:
   ```bash
   pip install -r requirements.txt


## Owner

Owner of this application is Yaajneshini