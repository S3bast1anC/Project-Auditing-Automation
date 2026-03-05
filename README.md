# Project-Auditing-Automation
<b> Requirements </b>
- 2-Factor Authentication enabled on Google account
- App Password for IMAP access
- Ollama (Local server must be running if using `Ollama_Context_Log.py`)

**Pyhon-Only Context Log.py** <br>
Extracts all emails matching specific keywords within a date range. Best for quick and easy data gathering.
- **Output**: Generates a consolidated pdf (`TRANSCRIPT_{CompanyName}.pdf`) and a folder (`{CompanyName} Context Log`) of dated attachments. 

**Ollama Contenxt Log.py** <br>
Integrates a local LLM (Llama 3) to act as a filter. Still 80% Python, 20% AI. Analyzes email subjects and metadata to filter out administrative "noise." It checks a 3-year window (Target Year $\pm$ 1) to ensure no context is lost from the previous or following years. Better if keywords alone produce too many irrelevant results.
- **Output**: Produces a sanitized PDF transcript and attachment folder containing only project-relevant research data.
