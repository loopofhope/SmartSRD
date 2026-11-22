## Tech stack

| Layer | Technology | Used for |
|-------|-----------|----------|
| Language | Python 3.10-3.13 | Entire application |
| Web UI | Streamlit (1.50 or newer) | Forms, pages, buttons, role-based screens, file upload and download |
| Word generation | python-docx, plus direct OOXML editing | Builds the .docx: tables, headers and footers, and the clickable Table of Contents and List of Figures (bookmarks, internal hyperlinks, page-number fields) |
| Document to PDF | Microsoft Word via docx2pdf and pywin32 (Windows), or LibreOffice (headless) | Converts the .docx so it can be shown and downloaded as PDF |
| PDF viewer | PyMuPDF | Renders PDF pages as images, reads the PDF links so Table of Contents entries are clickable in the browser, and works out real page numbers |
| Local AI (optional) | Ollama with `llama3.2`, called over HTTP | Improves project descriptions and suggests requirements, entirely on the local machine |
| Storage | JSON files and a folder, no database | `data/projects.json` (documents, versions, decisions), `data/users.json` (people), `data/figures/` (uploaded images) |
| Packaging | Windows batch scripts and pip | One-click setup and run; the offline bundle ships pre-downloaded Windows wheels |

Python standard library modules also used: `json`, `urllib`, `pathlib`, `subprocess`, `tempfile`, `base64`, `uuid`, `datetime` and `re`.

### How the code is organised

| File | Responsibility |
|------|----------------|
| `app.py` | Streamlit screens: Create SRD, My documents, Review queue, Document history, About |
| `workflow.py` | Workflow logic: users, versions, review and approval decisions, status changes, and the document-control data written into the Word file |
| `srd_generator.py` | Builds the Word document, including the Document Control & Approval page, Table of Contents, List of Figures, Figures and Glossary |
| `pdf_preview.py` | Converts Word to PDF and shows the clickable page viewer in the browser |
| `document_viewer.py` | Simple on-screen fallback preview when no converter is available |
| `ai_assistant.py` | Talks to the local Ollama model |

### Workflow design

Each document follows this status path:

```
Draft -> In Review -> In Approval -> Approved
              \            /
               -> Changes Requested -> author revises -> Draft -> resubmit as next version
```

Each submission creates a numbered version (1.0, 1.1, ...) that keeps its own reviewer, approver, decisions and comments.

### Design choices and limits

- **Local-first:** everything runs on the user's computer, with no cloud service or API key required.
- **Prototype storage:** JSON files are simple and portable, but they don't suit many people editing at once.
- **No real authentication:** users are chosen from a list. For company use, add login, access control, audit logging, backups and a proper database such as PostgreSQL or SQLite.
- **Windows focus:** the setup scripts and the Word-based conversion target Windows.

## Running the app

### Requirements

- Windows 10/11
- 64-bit Python 3.10, 3.11, 3.12 or 3.13 (during installation, tick **Add Python to PATH**)
- A modern browser (Chrome or Edge)
- Optional: Microsoft Word or LibreOffice, to show the document inside the app
- Optional: Ollama with a local model, for AI suggestions

### Option 1: One-click start (needs internet the first time)

1. Download or unzip the project to a folder, for example `C:\SmartSRDBuilder`.
2. Double-click `run_windows.bat`.
3. The first run creates a virtual environment and installs the packages, which takes a few minutes. Later runs start in seconds.
4. Your browser opens at http://localhost:8501. If it doesn't, copy that address into the browser.

Keep the black command window open while you use the app.

### Option 2: Manual start

Open Command Prompt in the project folder and run:

```bat
py -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
streamlit run app.py
```

Next time, you only need:

```bat
.venv\Scripts\activate
streamlit run app.py
```

### Option 3: Offline computer (no internet)

Use the offline bundle (`smart_srd_builder_offline.zip`), which includes all the packages.

1. Copy the zip to the offline computer and unzip it.
2. Install Python 3.10-3.13 from an installer you brought with you.
3. Double-click `setup_offline_windows.bat` (once).
4. Double-click `run_offline_windows.bat` to start the app.

See `README_OFFLINE.md` in that bundle for details.

### Stopping the app

Click the command window and press **Ctrl + C**, or just close the window.

### Letting others on the same network open it

```bat
streamlit run app.py --server.address 0.0.0.0
```

Other people can then open `http://<your-computer-IP>:8501`. Windows may ask you to allow access through the firewall.

### Optional: local AI with Ollama

1. Install Ollama from https://ollama.com/download.
2. Run `ollama pull llama3.2` in Command Prompt.
3. Make sure Ollama is running, then refresh the app. The sidebar will show "Local AI is available".

The app works fully without this.

### Troubleshooting

| Problem | Fix |
|---------|-----|
| `py` is not recognized | Reinstall Python with **Add Python to PATH** ticked, or use `python` instead of `py` |
| Port 8501 is busy | `streamlit run app.py --server.port 8502` |
| Browser doesn't open | Go to http://localhost:8501 manually |
| Document shows a simplified view | Install LibreOffice (or use a PC with Word), then restart the app |
| Offline setup says packages don't match | Your Python version differs; use Python 3.10-3.13 (64-bit) |
| `ModuleNotFoundError` | Activate the virtual environment (`.venv\Scripts\activate`) and run `pip install -r requirements.txt` |
| Word asks to update fields when opening a download | Click **Yes**. It refreshes the page numbers |
| Want to reset all data | Close the app, then delete the `data` folder |



## How to use

The app has three roles. Pick who you are in the sidebar under **Signed in as**.

| Role | Person (default) | What they do |
|------|------------------|--------------|
| Author | Aishwarya (Project Engineer) | Creates the SRD, assigns reviewer and approver, submits |
| Reviewer | Ramesh (Manager) | Checks the document and approves or requests changes |
| Approver | Uttam Singh (DGM) | Gives final approval |

You can add more people from **Add a person** in the sidebar (or by editing `data/users.json`).

### 1. Author: create and submit an SRD

1. Sign in as the author and open **Create SRD**.
2. Fill in the project information. A project name is required; everything else can be completed later.
3. (Optional) Use the AI buttons to enhance the description or suggest requirements. If Ollama isn't installed, skip this; the rest works normally. Always review AI suggestions before submitting.
4. (Optional) Open **Additional sections** to add scope, assumptions, dependencies, risks and acceptance criteria.
5. (Optional) Open **Figures and glossary**:
   - Upload PNG or JPG images and edit each caption. They appear in section 10 and in the List of Figures.
   - Edit the glossary, one entry per line as `Term: definition`.
6. Click **Save draft**. The document gets an ID such as SRD-001. Use **Preview saved draft as a document** to see the real Word layout.
7. Choose a **Reviewer** and an **Approver**, then click **Submit for review**.
   - They must be two different people, and neither can be the author.
   - Submitting locks the draft and creates version 1.0.

### 2. Reviewer: review the document

1. Sign in as the reviewer. The sidebar shows how many documents are waiting for you.
2. Open **Review queue** and select the document.
3. Read it in the document viewer (the real Word layout, shown as pages). Click any Table of Contents or List of Figures entry to jump to that page. Use **Download Word** or **Download PDF** if you prefer to read it offline.
4. Choose **Approved** or **Changes Requested**, add comments (required when requesting changes), and click **Record decision**.

### 3. Approver: give final approval

1. Sign in as the approver and open **Review queue**. The document appears after the reviewer approves it.
2. Read the document, including the reviewer's decision and comments.
3. Choose **Approved** or **Changes Requested**, add comments, and click **Record decision**.
4. When approved, the status becomes **Approved**.

### 4. If changes are requested

1. Sign in as the author. A red alert in the sidebar shows that a document needs changes.
2. Open **My documents**. The reviewer's or approver's comments are shown.
3. Click **Open for editing & resubmit**, make the changes, and click **Save draft**.
4. Check or change the reviewer and approver, then click **Submit for review**.
5. The next version is created (1.1, 1.2, ...). Earlier versions and their decisions are kept in **Document history**.

A document sent back by the approver goes through review again from the start.

### 5. Download the final document

Once approved, open **My documents** (or **Document history**) and click **Download final approved SRD**. The Word file contains the names, dates, decisions, comments, version history and a green **FINAL STATUS: APPROVED** banner. All of this is filled in automatically from the workflow, so nobody types names into the document.

### What the Word document contains

- Cover page
- Document Control & Approval (document details, ownership, review and approval record, version history, final status)
- Table of Contents (clickable)
- List of Figures (clickable)
- Sections 1–9: introduction, stakeholders, scope, functional and non-functional requirements, assumptions, dependencies, risks, acceptance criteria
- Section 10: Figures and Diagrams
- Section 11: Glossary

When you open the downloaded Word file, click **Yes** if Word asks to update fields. This refreshes the page numbers. Ctrl+click a Table of Contents entry to jump to it.

### Pages in the app

| Page | Who uses it | Purpose |
|------|-------------|---------|
| Create SRD | Author | Write, save and submit a document |
| My documents | Author | See status, reopen for edits, download final file |
| Review queue | Reviewer / Approver | Read, comment and decide |
| Document history | Everyone | All documents, versions, decisions and downloads |

### Good to know

- **Document view needs a converter.** The in-browser Word view needs Microsoft Word (Windows) or LibreOffice installed. Without either, you get a simplified preview. Downloads still work.
- **Sign-in is a user picker, not real authentication.** Add proper login before using this with real company data.
- **Everything is stored locally** in the `data` folder (documents, versions, users and uploaded figures). Back up or copy this folder to keep your work.
- **To start fresh,** delete `data/projects.json` (and `data/users.json` to reset the people list).
