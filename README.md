AI Executive Assistant (Google Apps Script + Gemini API)
This project automates the creation of daily executive briefings by integrating the Google Workspace ecosystem (Calendar, Gmail, and Drive) with generative AI via the Gemini API (Google AI Studio).

Features
Calendar Briefing: Lists today's and tomorrow's appointments, including schedules and attendees.
Email Summarization: Parses unread threads in the Inbox, grouping them by project or topic and highlighting action items.
Obsidian Integration (Google Drive): Recursively scans directories and subdirectories in Google Drive to read .md files, consolidating project statuses and management notes.
VIP Priority: Identifies critical senders via a whitelist and prioritizes them at the top of the briefing.
Google Chat Alerts: Asynchronous webhook notifications triggered when VIP emails are detected.
Network Resilience (Exponential Backoff): Intelligent retry mechanism to handle temporary API instability or rate limiting (HTTP 429/503) from the Gemini API.
Smart Execution & Routing: Detects weekends to save compute resources and adjusts the data extraction payload based on the time of day (comprehensive briefings in the morning, concise summaries in the afternoon/evening).
Secure Configuration

The codebase leverages Google Apps Script's PropertiesService to prevent hardcoding API keys, emails, and sensitive data in the source code. You must configure the following environment variables in your Script Properties (Project Settings).
Required Environment Variables:
API_KEY: Your Google AI Studio (Gemini) API Key.
EMAIL_DESTINO: Target email address to receive the HTML briefings.
NOME_USUARIO: Your full name, injected into the LLM prompt for context.
WEBHOOK_CHAT: (Optional) Google Chat webhook URL for VIP alerts.
LISTA_VIP: Comma-separated list of priority domains or emails (e.g., @company.com, ceo@gmail.com).
DIAS_HISTORICO: Lookback period for email processing (e.g., 3d).
PASTAS_OBSIDIAN: Comma-separated Google Drive root folder IDs containing your .md files.

Installation and Usage
Create a new project in Google Apps Script.
Copy the contents of Código.gs from this repository into the editor.
Set up all the environment variables listed above in the Script Properties.
Authorization: Run the gerarResumoDiario function manually for the first time. Google will prompt you to authorize the script's access to your Gmail, Calendar, and Drive.
Set up time-driven Triggers for your preferred briefing schedules (e.g., 07:30, 12:45, 18:30).

License
Distributed under the License. See the LICENSE file for more information.
