import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def create_document():
    doc = docx.Document()
    
    # ── Page setup ─────────────────────────────────────────────────────────
    for section in doc.sections:
        section.top_margin = Inches(1.2)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.header_distance = Inches(0.5)
        
        # Header setup (Logo layout)
        header = section.header
        header_table = header.add_table(1, 2, Inches(6.5))
        header_table.columns[0].width = Inches(3.25)
        header_table.columns[1].width = Inches(3.25)
        
        # Cell 1: SmartBridge Logo (left)
        cell_left = header_table.cell(0, 0)
        p_left = cell_left.paragraphs[0]
        p_left.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_left.paragraph_format.space_before = Pt(0)
        p_left.paragraph_format.space_after = Pt(0)
        r_left = p_left.add_run()
        if os.path.exists('Images/Google/1.jpeg'):
            r_left.add_picture('Images/Google/1.jpeg', width=Inches(1.5))
        else:
            r_left.text = "SMARTBRIDGE"
            
        # Cell 2: SkillWallet Logo (right)
        cell_right = header_table.cell(0, 1)
        p_right = cell_right.paragraphs[0]
        p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_right.paragraph_format.space_before = Pt(0)
        p_right.paragraph_format.space_after = Pt(0)
        r_right = p_right.add_run()
        if os.path.exists('Images/Google/2.jpeg'):
            r_right.add_picture('Images/Google/2.jpeg', width=Inches(1.2))
        else:
            r_right.text = "SKILL WALLET"

    # ── ALL BLACK color definitions ────────────────────────────────────────
    c_black = RGBColor(0, 0, 0)
    c_caption = RGBColor(80, 80, 80)  # Only for figure captions (dark grey)

    # ── Helper: Page Heading ───────────────────────────────────────────────
    def add_page_heading(doc, text, level=1):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.color.rgb = c_black
        if level == 1:
            run.font.name = 'Arial'
            run.font.size = Pt(18)
        elif level == 2:
            run.font.name = 'Arial'
            run.font.size = Pt(14)
        return p

    # ── Helper: Body Paragraph ─────────────────────────────────────────────
    def add_body_paragraph(doc, text, bold_prefix=None, space_after=6):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        
        if bold_prefix:
            run_pre = p.add_run(bold_prefix)
            run_pre.bold = True
            run_pre.font.name = 'Calibri'
            run_pre.font.size = Pt(11)
            run_pre.font.color.rgb = c_black
            
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(11)
        run.font.color.rgb = c_black
        return p

    # ── Helper: Bullet Item ────────────────────────────────────────────────
    def add_bullet_item(doc, bold_title, text):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        
        run_title = p.add_run(bold_title + ": ")
        run_title.bold = True
        run_title.font.name = 'Calibri'
        run_title.font.size = Pt(11)
        run_title.font.color.rgb = c_black
        
        run_text = p.add_run(text)
        run_text.font.name = 'Calibri'
        run_text.font.size = Pt(11)
        run_text.font.color.rgb = c_black
        return p

    # ── Helper: Code Block ─────────────────────────────────────────────────
    def add_code_block(doc, code_text):
        table = doc.add_table(rows=1, cols=1)
        table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
        cell = table.cell(0, 0)
        
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), 'F5F5F5')
        tcPr.append(shd)
        
        borders = OxmlElement('w:tcBorders')
        for border_name in ['top', 'left', 'bottom', 'right']:
            b = OxmlElement(f'w:{border_name}')
            b.set(qn('w:val'), 'single')
            b.set(qn('w:sz'), '4')
            b.set(qn('w:space'), '0')
            b.set(qn('w:color'), 'CCCCCC')
            borders.append(b)
        tcPr.append(borders)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.left_indent = Pt(6)
        p.paragraph_format.right_indent = Pt(6)
        
        run = p.add_run(code_text)
        run.font.name = 'Consolas'
        run.font.size = Pt(9.5)
        run.font.color.rgb = c_black
        doc.add_paragraph().paragraph_format.space_after = Pt(4)
        return table

    # ── Helper: Screenshot ─────────────────────────────────────────────────
    def add_screenshot(doc, img_path, caption, width=5.5):
        if os.path.exists(img_path):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run()
            run.add_picture(img_path, width=Inches(width))
            
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_after = Pt(12)
            run_cap = p_cap.add_run(f"Figure: {caption}")
            run_cap.italic = True
            run_cap.font.name = 'Calibri'
            run_cap.font.size = Pt(9.5)
            run_cap.font.color.rgb = c_caption
        else:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(f"[Screenshot Placeholder: {caption} ({img_path})]")
            run.bold = True
            run.italic = True
            run.font.color.rgb = c_black
            p.paragraph_format.space_after = Pt(12)

    # ══════════════════════════════════════════════════════════════════════════
    # PAGE 1 — PROJECT DESCRIPTION & 4 SCENARIOS
    # ══════════════════════════════════════════════════════════════════════════
    add_page_heading(doc, "InterviewAI: AI-Powered Interview Question Generator", level=1)
    
    add_page_heading(doc, "Project Description:", level=2)
    add_body_paragraph(doc, 
        "InterviewAI is a state-of-the-art recruiter empowerment platform that harnesses Generative AI "
        "to deliver highly customized, role-specific, and context-aware interview questions instantly. "
        "Built with a sleek, premium single-page web interface and backed by a robust Flask microservice, "
        "InterviewAI automates the cognitive burden of generating balanced technical, behavioral, and "
        "situational questions tailored precisely to experience levels and unique organizational contexts. "
        "It integrates seamlessly with the Google Gemini API to guarantee fast, reliable response generation, "
        "coupled with a thorough evaluation framework — complete with green flags, red flags, and scoring scales — to "
        "help recruiters conduct structured, objective, and unbiased candidate assessments.")
        
    add_body_paragraph(doc, 
        "By leveraging environment-based key management for secure enterprise operation, persistent local JSON "
        "history tracking for easy session retrieval, and automated document compilation using python-docx, "
        "InterviewAI establishes a secure, standard-compliant tool. It eliminates repetitive administrative prep work, "
        "allowing HR professionals and technical managers to focus entirely on candidate engagement.")

    add_page_heading(doc, "Scenarios:", level=2)
    
    add_body_paragraph(doc, 
        "A technical lead is preparing to interview a candidate for a 'Senior React Developer' role. "
        "By selecting an Advanced experience level and mixed interview type, the lead generates ten customized questions, "
        "complete with follow-up prompts and evaluation tips. The system also supplies green flags to watch for and "
        "red flags to avoid, ensuring the evaluation is highly objective.",
        "Scenario 1: ")
    
    add_body_paragraph(doc, 
        "An HR manager needs to conduct a soft-skills screening for a 'Product Manager' role. "
        "By choosing the 'Behavioral' interview type, the manager generates situational and behavioral questions "
        "that target agile methodologies, roadmapping, and conflict resolution, without needing deep technical expertise.",
        "Scenario 2: ")
    
    add_body_paragraph(doc, 
        "A hiring team wants to conduct a quick screening. They generate 5 beginner-level questions. "
        "After the session, they export the generated questions to a clean Microsoft Word (.docx) document, "
        "formatted professionally and ready to be printed or shared across the department.",
        "Scenario 3: ")
    
    add_body_paragraph(doc, 
        "A recruiter wants to review an interview set from last week. They navigate to the sidebar history "
        "or the landing page dashboard, click on the saved 'Senior React Developer' session, and instantly "
        "reload the entire questions dashboard without making a new API request, saving time and API resource quota.",
        "Scenario 4: ")

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    # PAGE 2 — TECHNICAL ARCHITECTURE (WITH IMAGE) & PREREQUISITES
    # ══════════════════════════════════════════════════════════════════════════
    add_page_heading(doc, "Technical Architecture", level=1)
    add_body_paragraph(doc, 
        "The technical architecture of InterviewAI follows a decoupled, asynchronous Single-Page Application (SPA) "
        "design, separating client-side presentation from backend business logic. This separation ensures extreme "
        "responsiveness and easy deployment. The diagram below illustrates the complete component flow across "
        "all three layers of the system:")

    # Architecture diagram image
    add_screenshot(doc, "Images/architecture_diagram.png", 
                   "InterviewAI Technical Architecture — Three-Layer Component Flow Diagram")
    
    add_body_paragraph(doc, 
        "The Frontend Layer runs entirely in the user's browser using standard HTML5, CSS3, and ES6 JavaScript. "
        "It communicates with the Backend Layer via asynchronous HTTP POST and GET requests carrying JSON payloads. "
        "The Backend Layer, built on Flask micro-framework, hosts RESTful API endpoints for question generation, "
        "history management, and document export. It connects securely to the Google Gemini API over HTTPS "
        "for AI-powered question generation, and manages local persistent storage through a structured JSON file.")

    add_page_heading(doc, "Pre-requisites:", level=2)
    add_bullet_item(doc, "Flask Framework Knowledge", 
        "Understanding micro-routing, template rendering, and serving static files. Reference: Flask Documentation (https://flask.palletsprojects.com/)")
    add_bullet_item(doc, "Google Gemini API Familiarity", 
        "Understanding prompt structure, system instructions, temperature, and JSON mode generation. Reference: Google Gemini API (https://ai.google.dev/)")
    add_bullet_item(doc, "HTML, CSS, and JavaScript Skills", 
        "Proficiency in document structuring, dynamic DOM manipulation via Fetch, and styling with modern flex/grid. Reference: W3Schools Web Tutorials")
    add_bullet_item(doc, "Python Programming Proficiency", 
        "Solid understanding of modules, exception handling, dictionary parsing, and packages. Reference: Python Official Docs")
    add_bullet_item(doc, "Version Control with Git", 
        "Commit structures, branching, repository management. Reference: Git Documentation")
    add_bullet_item(doc, "Development Environment Setup", 
        "Local virtual environments, environment variable loaders. Reference: Flask Installation Guide")

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    # PAGE 3 — PROJECT WORKFLOW WITH ACTIVITIES (DETAILED)
    # ══════════════════════════════════════════════════════════════════════════
    add_page_heading(doc, "Project Workflow", level=1)
    add_body_paragraph(doc, 
        "The development of the InterviewAI platform was executed systematically through a structured, multi-phase "
        "workflow. Each phase (Activity) was designed to build upon the accomplishments of the previous phase, "
        "ensuring strict standard compliance and code quality throughout the lifecycle. Below is a detailed "
        "breakdown of all activities performed during the project development:")

    add_page_heading(doc, "Activity 1: Model Selection and Architecture", level=2)
    add_body_paragraph(doc, 
        "This activity focused on establishing the foundational infrastructure for the entire project. "
        "It involved generating and securely configuring a Google Gemini API key within the project environment, "
        "researching and evaluating multiple Generative AI models to identify the optimal choice for low-latency "
        "web generation, defining the complete logical architecture detailing data flow between the frontend "
        "and backend layers, and establishing the development workspace by installing Flask, dotenv, "
        "and all necessary generative AI packages inside an isolated Python virtual environment.")
    add_bullet_item(doc, "Activity 1.1", "Generate a secure Google Gemini API key and register in the project environment.")
    add_bullet_item(doc, "Activity 1.2", "Research and select the appropriate Generative AI model for fast response latency.")
    add_bullet_item(doc, "Activity 1.3", "Define the logical architecture, detailing data flow between the frontend and backend.")
    add_bullet_item(doc, "Activity 1.4", "Establish the development workspace, installing Flask, dotenv, and generative packages.")

    add_page_heading(doc, "Activity 2: Core Functionalities Development", level=2)
    add_body_paragraph(doc, 
        "This activity centered on building the core intelligence and data management layer of the application. "
        "It involved developing sophisticated AI generation prompts that instruct Google Gemini to produce "
        "technical, behavioral, and situational interview questions with proper difficulty distribution. "
        "Additionally, a persistent local history storage system was implemented to log, read, and delete "
        "interview sessions using structured JSON, and a python-docx file compilation service was built "
        "to export generated questions as professionally formatted Word documents.")
    add_bullet_item(doc, "Activity 2.1", "Develop core AI generation prompts for technical, behavioral, and situational segments.")
    add_bullet_item(doc, "Activity 2.2", "Implement persistent local history storage to log, read, and delete interview sessions.")
    add_bullet_item(doc, "Activity 2.3", "Implement python-docx file compilation to export generated questions professionally.")

    add_page_heading(doc, "Activity 3: App.py Backend Development", level=2)
    add_body_paragraph(doc, 
        "This activity focused on implementing the Flask micro-framework backend that serves as the central "
        "nervous system of the application. It involved writing the main Flask initialization with CORS setup "
        "and static file serving, implementing the /api/generate endpoint with comprehensive data validation "
        "and error handling, building RESTful history CRUD handlers for listing, reading, and deleting sessions, "
        "and creating the /api/export endpoint that compiles question data into downloadable Word documents.")
    add_bullet_item(doc, "Activity 3.1", "Write main Flask initialization, setup CORS, and route standard index path.")
    add_bullet_item(doc, "Activity 3.2", "Implement /api/generate handler, including data validation and try-except safety.")
    add_bullet_item(doc, "Activity 3.3", "Implement history CRUD handlers for listing, reading individual, and deleting sessions.")
    add_bullet_item(doc, "Activity 3.4", "Implement /api/export handler for Word document compilation and download streaming.")

    add_page_heading(doc, "Activity 4: Frontend Development", level=2)
    add_body_paragraph(doc, 
        "This activity addressed the entire user interface layer. It involved designing a modern responsive "
        "layout using CSS with a customizable dark-theme glassmorphism aesthetic, implementing a Single-Page "
        "Application (SPA) router in vanilla JavaScript with async form submission using the Fetch API, "
        "and refactoring the homepage flow by centering hero content, adding an 'Explore More' reveal "
        "interaction, and creating a premium landing experience with animated glow orbs.")
    add_bullet_item(doc, "Activity 4.1", "Design modern responsive layout using CSS containing customizable dark-theme glassmorphism.")
    add_bullet_item(doc, "Activity 4.2", "Implement SPA single-page view router in vanilla JS and async form submission using Fetch.")
    add_bullet_item(doc, "Activity 4.3", "Refactor homepage flow by centering hero content, removing visual cards, and adding Explore More.")

    add_page_heading(doc, "Activity 5: Deployment & Verification", level=2)
    add_body_paragraph(doc, 
        "This final activity focused on verifying, validating, and deploying the completed system. "
        "It involved preparing the server environment and verifying all local network calls via Chrome DevTools, "
        "confirming CORS headers are active and API responses are correct, and exposing the local port "
        "publicly using Ngrok tunnels to support cross-device mobile testing under real network conditions.")
    add_bullet_item(doc, "Activity 5.1", "Prepare server environment and verify local network calls via Chrome DevTools.")
    add_bullet_item(doc, "Activity 5.2", "Expose local port publicly using Ngrok tunnels to support mobile testing.")

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    # MILESTONE 1 — MODEL SELECTION & ARCHITECTURE
    # ══════════════════════════════════════════════════════════════════════════
    add_page_heading(doc, "Milestone 1: Model Selection and Architecture", level=1)
    add_body_paragraph(doc, 
        "The primary focus of Milestone 1 is to establish a secure project environment, evaluate and select the best "
        "artificial intelligence model, and set up the structural foundation of the workspace. This ensures the "
        "application is built on a compliant, secure, and highly scalable basis.")
        
    add_page_heading(doc, "Activity 1.1: Generate a Gemini API Key", level=2)
    add_body_paragraph(doc, 
        "Before the application can communicate with the Google Generative AI servers, we must obtain a valid API key "
        "and configure it securely within the local project workspace. Below are the precise steps followed:")
        
    add_bullet_item(doc, "Step 1 — Create a Google AI Studio Account", 
        "Visit the Google AI Studio page (https://aistudio.google.com/) and register using a developer or enterprise account.")
    add_bullet_item(doc, "Step 2 — Generate the API Key", 
        "Navigate to the 'Get API Key' section, select 'Create API Key in New Project', and copy the generated secure token string immediately.")
    add_bullet_item(doc, "Step 3 — Local Security Setup", 
        "In the project root folder, verify that a file named .env is present. Save the key using the exact variable name expected by the service wrapper.")
    
    # API Key screenshots from Google folder
    add_screenshot(doc, "Images/Google/1.jpeg", "Google AI Studio — Searching for Google Studio to access the API Key portal.", width=5.0)
    add_screenshot(doc, "Images/Google/2.jpeg", "Google AI Studio — Creating a new API Key in the 'Get API Key' dashboard.", width=5.0)
    
    add_body_paragraph(doc, "Here is the structure used inside the project's configuration file:")
    add_code_block(doc, "GEMINI_API_KEY=your_api_key_here\nPORT=5000")
    
    add_body_paragraph(doc, 
        "To guarantee that this API key is never leaked or committed to source control repositories, the .env file "
        "is explicitly added to the project's .gitignore file, enforcing responsible developer best practices.")

    add_page_heading(doc, "Activity 1.2: Research and Select the Generative AI Model", level=2)
    add_body_paragraph(doc, 
        "Selecting the optimal large language model requires a thorough balance between response latency, context "
        "window capacity, API request quota limits, and output format reliability. We evaluated several options:")
        
    add_bullet_item(doc, "Gemini 1.5 Pro", 
        "Offers extremely deep reasoning and a massive 2-million token context window, but has a higher latency and lower request-per-minute limits on the free tier, making it suboptimal for fast web generation.")
    add_bullet_item(doc, "Gemini 1.5 Flash", 
        "Specifically engineered for high speed, low latency, and highly efficient processing. It features a robust JSON formatting mode and boasts a generous free-tier rate limit, making it the perfect candidate for instant generation.")
    
    add_body_paragraph(doc, 
        "Consequently, we chose the 'gemini-flash-latest' model. In addition, we configured the generator to "
        "operate at a temperature of 0.7, encouraging high creative variety in behavioral/situational "
        "questions while maintaining strict professional standards.")

    add_page_heading(doc, "Activity 1.4: Project Structure", level=2)
    add_body_paragraph(doc, 
        "The workspace structure is organized logically with a clear separation between backend services "
        "and frontend presentation layers as shown below:")
    add_screenshot(doc, "Images/req.txt.jpeg", "Project directory structure showing backend and frontend file organization.")

    add_body_paragraph(doc, 
        "The backend requirements file (requirements.txt) explicitly defines the library constraints to guarantee "
        "environment stability during runtime, including python-docx for document exporting and dotenv for variable loading.")

    add_body_paragraph(doc, "To prepare the virtual environment, the following command line sequence was executed:")
    add_code_block(doc, 
        "# Navigate to backend directory\n"
        "cd backend\n\n"
        "# Create Python Virtual Environment\n"
        "python -m venv venv\n\n"
        "# Activate the environment (Windows PowerShell)\n"
        ".\\venv\\Scripts\\Activate.ps1\n\n"
        "# Install the mandatory dependencies\n"
        "pip install flask flask-cors google-generativeai python-docx python-dotenv requests")

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    # MILESTONE 2 — CORE FUNCTIONALITIES DEVELOPMENT
    # ══════════════════════════════════════════════════════════════════════════
    add_page_heading(doc, "Milestone 2: Core Functionalities Development", level=1)
    add_body_paragraph(doc, 
        "Milestone 2 focuses on building the core intelligence and data management modules that power "
        "the application. This includes the AI prompt engineering, persistent history storage, and "
        "professional document export capabilities.")
    
    add_page_heading(doc, "Activity 2.1: Develop Core AI Question Generation Features", level=2)
    add_body_paragraph(doc, 
        "The core functionality of the AI system lies in constructing a robust, context-sensitive prompt that instructs "
        "Google Gemini to act as a world-class HR consultant and senior technical lead. We implemented this in "
        "ai_service.py using dynamic variables mapped from the user's frontend input. The prompt enforces a strict "
        "JSON schema to ensure reliable frontend parsing, and distributes questions based on the selected interview type:")
        
    add_bullet_item(doc, "Technical Mode", 
        "Distributes questions heavily in favor of domain knowledge, system design, and coding practices (approx. 70% technical, 20% behavioral, 10% situational).")
    add_bullet_item(doc, "Behavioral Mode", 
        "Focuses on soft skills, teamwork, and cultural alignment (approx. 60% behavioral, 30% situational, 10% technical).")
    add_bullet_item(doc, "Mixed Mode", 
        "Provides a comprehensive, well-rounded assessment suitable for standard interviews (approx. 40% technical, 35% behavioral, 25% situational).")

    add_page_heading(doc, "Activity 2.2: Develop Session History Management", level=2)
    add_body_paragraph(doc, 
        "InterviewAI includes a persistent history logging service (history_service.py) that writes session data "
        "locally to a structured JSON file (history.json). This ensures history is preserved across server restarts. "
        "Each history entry is logged with a unique UUID, a timestamp, and metadata containing the role, experience level, "
        "and requested question count, along with the raw AI question payload. This layout facilitates fast loading times "
        "without needing additional API requests.")

    add_page_heading(doc, "Activity 2.3: Develop Word Document Export Service", level=2)
    add_body_paragraph(doc, 
        "Using python-docx, we developed a high-quality, professional export service that compiles the generated "
        "questions, follow-ups, evaluation tips, and full scoring guides directly into a Microsoft Word document. "
        "The compiled file is built dynamically in the system's temporary directory, ensuring server disk space is "
        "managed efficiently, before being streamed securely to the user's web browser.")

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    # MILESTONE 3 — APP.PY BACKEND DEVELOPMENT (WITH ALL APP IMAGES)
    # ══════════════════════════════════════════════════════════════════════════
    add_page_heading(doc, "Milestone 3: App.py Backend Development", level=1)
    add_body_paragraph(doc, 
        "Milestone 3 focuses on implementing the Flask micro-framework backend, setting up CORS permissions, "
        "validating request data, and linking backend routes to our core AI and history services. "
        "Below are the key code implementations with visual references:")

    add_page_heading(doc, "Activity 3.1: Generate Questions Endpoint (/api/generate)", level=2)
    add_body_paragraph(doc, 
        "The /api/generate endpoint accepts asynchronous JSON POST requests from the client. It validates input "
        "properties such as job role (required), question count range (1-30), and optional fields like skills and "
        "company context. Upon successful validation, it forwards the parameters to the Google Gemini service wrapper "
        "and saves the result to persistent history before returning the response.")
    add_screenshot(doc, "Images/app/generative questions.jpeg", 
                   "Generate Questions — The /api/generate POST route implementation showing input validation, "
                   "AI service call, and history persistence.", width=5.0)

    add_page_heading(doc, "Activity 3.2: Evaluation Criteria Structure", level=2)
    add_body_paragraph(doc, 
        "Each generated question set includes a comprehensive evaluation criteria section containing a 1-5 scoring "
        "scale, key competencies to assess, green flags indicating positive candidate signals, and red flags "
        "highlighting warning signs. This structured framework ensures recruiters conduct objective assessments.")
    add_screenshot(doc, "Images/app/Evalution criteria.jpeg", 
                   "Evaluation Criteria — Structured scoring scale, key competencies, green flags, and red flags "
                   "builder in ai_service.py.", width=5.0)

    add_page_heading(doc, "Activity 3.3: History Management Endpoints (/api/history)", level=2)
    add_body_paragraph(doc, 
        "To support history state synchronization between the client dashboard and the persistent database, "
        "we implemented RESTful CRUD endpoints in app.py. The GET /api/history endpoint returns metadata for all "
        "sessions, GET /api/history/<item_id> retrieves a specific session with full question data, and "
        "DELETE /api/history/<item_id> removes a session permanently from the JSON store.")
    add_screenshot(doc, "Images/app/history Management.jpeg", 
                   "History Management — RESTful CRUD endpoints for session listing, retrieval, and deletion.", width=5.0)

    add_page_heading(doc, "Activity 3.4: Export as DOCX & PDF (/api/export)", level=2)
    add_body_paragraph(doc, 
        "The /api/export endpoint accepts the active question set payload from the client browser and compiles it "
        "into a formatted Word (.docx) document using python-docx. The document includes structured sections for "
        "technical, behavioral, and situational questions, each with difficulty badges, follow-up prompts, and "
        "evaluation tips. The file is streamed back as a downloadable attachment.")
    add_screenshot(doc, "Images/app/Export as Doc & PDF.jpeg", 
                   "Export Service — python-docx document compiler with custom formatting and file streaming.", width=5.0)

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    # MILESTONE 4 — FRONTEND DEVELOPMENT
    # ══════════════════════════════════════════════════════════════════════════
    add_page_heading(doc, "Milestone 4: Frontend Development", level=1)
    
    add_page_heading(doc, "Activity 4.1: Designing the User Interface (HTML & CSS)", level=2)
    add_body_paragraph(doc, 
        "The frontend is built as a pure Single-Page Application (SPA) using vanilla HTML5, CSS3, and ES6 JavaScript, "
        "eliminating heavy compilation frameworks. The core page contains three dedicated content sections structured "
        "inside container divs. Navigation and view toggles are managed purely client-side by applying or removing "
        "a helper .hidden utility class:")
    add_bullet_item(doc, "Page 1: Home Landing (#page-home)", 
        "Presents a premium hero overview, active stats indicator, a top right 'Let's Start' action, an 'Explore More' button, and a visual history summary of recent sessions.")
    add_bullet_item(doc, "Page 2: Configure Form (#page-form)", 
        "Displays a gorgeous form card with fields for job role, required skills, experience level selection, question count sliders, interview type selectors, and optional company context.")
    add_bullet_item(doc, "Page 3: Results Dashboard (#page-results)", 
        "Features a two-column interactive results space containing a history navigation sidebar on the left and a main results display tab panel on the right.")

    add_page_heading(doc, "Activity 4.2: CSS Styling and Glassmorphism Aesthetics", level=2)
    add_body_paragraph(doc, 
        "To deliver a premium visual experience, the interface utilizes a dark-mode glassmorphism design. "
        "All styles are defined in styles.css using modern CSS custom properties for uniform token management:")
    add_bullet_item(doc, "Animated Glow Orbs", 
        "Three blurred gradient background orbs (.orb-1, .orb-2, .orb-3) utilize CSS @keyframes animations to shift position, scale, and opacity slowly. This adds visual depth without impacting CPU usage.")
    add_bullet_item(doc, "Glassmorphism Cards", 
        "Cards are styled with a semi-transparent background (rgba(255, 255, 255, 0.04)), custom borders, and a backing backdrop-filter: blur(20px) style. This allows the background glows to show through subtly.")
    add_bullet_item(doc, "Hover Transitions", 
        "Micro-transitions are applied to all cards and buttons. Hover actions trigger scale changes (transform: translateY(-2px)), glow highlights, and icon transitions, creating a tactile feel.")

    add_page_heading(doc, "Activity 4.3: Dynamic Client Logic in app.js", level=2)
    add_body_paragraph(doc, 
        "Client interaction is driven by app.js. It manages local application state, handles navigation transitions, "
        "and compiles Fetch requests to the server:")
    add_bullet_item(doc, "SPA Router Toggles", 
        "The showPage(pageId) function controls the visible screen by dynamically adding the .hidden class to unused views while resetting the viewport scroll to the top of the window.")
    add_bullet_item(doc, "Dynamic Card Compilers", 
        "Upon receiving JSON payloads from the server, template literal templates compile individual question elements recursively. This builds cards containing badge classes based on category and difficulty.")
    add_bullet_item(doc, "Evaluation Renderer", 
        "A dedicated function parses green flags, red flags, competencies, and interviewer tips, rendering them into a tab panel that HR professionals can review side-by-side.")
    add_bullet_item(doc, "Sidebar History Toggles", 
        "When history lists update, the sidebar and home grid synchronize, showing relative timestamps and permitting instant deletion or loading of past entries.")

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    # MILESTONE 5 — DEPLOYMENT & VERIFICATION (WITH TERMINAL OUTPUT)
    # ══════════════════════════════════════════════════════════════════════════
    add_page_heading(doc, "Milestone 5: Deployment & Verification", level=1)
    
    add_page_heading(doc, "Activity 5.1: Preparing the Application for Local Deployment", level=2)
    add_body_paragraph(doc, 
        "Milestone 5 focuses on verifying, validating, and deploying the completed InterviewAI system. "
        "First, we perform local validation to ensure all API components communicate flawlessly.")
        
    add_bullet_item(doc, "Start the Flask App", 
        "Activate the Python environment and launch the backend: python app.py. This serves the server on port 5000.")
    add_bullet_item(doc, "Browse Local Host", 
        "Open a web browser and visit http://localhost:5000 to access the application directly through the local server.")
    add_bullet_item(doc, "Verify Health API", 
        "Confirm connection by checking /api/health. This ensures that CORS headers are active and database links are online.")
    add_bullet_item(doc, "Validate Generation Flows", 
        "Test all three types (Technical, Behavioral, Mixed) across multiple experience levels to confirm that output schemas are valid and error banners are handled correctly.")

    add_body_paragraph(doc, "Below is the terminal output showing the Flask server starting successfully:")
    add_code_block(doc, 
        "PS C:\\Users\\HP\\Desktop\\Interview-Generator\\backend> .\\venv\\Scripts\\Activate.ps1\n"
        "(venv) PS C:\\Users\\HP\\Desktop\\Interview-Generator\\backend> python app.py\n\n"
        "[OK] Interview Generator API running at http://localhost:5000\n\n"
        " * Serving Flask app 'app'\n"
        " * Debug mode: on\n"
        " * Running on http://127.0.0.1:5000\n"
        "Press CTRL+C to quit\n"
        " * Restarting with stat\n"
        " * Debugger is active!\n"
        " * Debugger PIN: 123-456-789")

    add_page_heading(doc, "Activity 5.2: Expose Locally Hosted Server Publicly using Ngrok", level=2)
    add_body_paragraph(doc, 
        "To allow recruiters, managers, and clients to test the application on mobile devices or remote workstations "
        "without deploying to a cloud hosting platform, we configure a secure public Ngrok tunnel.")
        
    add_bullet_item(doc, "Step 1 — Download Ngrok", 
        "Download the Ngrok binary from the official website and authorize it using your personal authentication token.")
    add_bullet_item(doc, "Step 2 — Establish the HTTP Tunnel", 
        "In the console, run: ngrok http 5000. This creates a secure, public HTTPS forwarding URL mapping directly to local port 5000.")
    add_bullet_item(doc, "Step 3 — Cross-Device Verification", 
        "Load the HTTPS URL on mobile browsers to test responsive UI wrapping, touch sliding, and async question fetch times under cellular network conditions.")
    
    add_body_paragraph(doc, "Terminal output for Ngrok tunnel establishment:")
    add_code_block(doc, 
        "PS C:\\Users\\HP> ngrok http 5000\n\n"
        "ngrok                                                (Ctrl+C to quit)\n\n"
        "Session Status                online\n"
        "Account                       user@example.com\n"
        "Version                       3.x.x\n"
        "Region                        India (in)\n"
        "Latency                       45ms\n"
        "Web Interface                 http://127.0.0.1:4040\n"
        "Forwarding                    https://xxxx-xx-xx.ngrok-free.app -> http://localhost:5000")

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    # EXPLORING THE WEBSITE — ALL IMAGES FROM RESULTS FOLDER
    # ══════════════════════════════════════════════════════════════════════════
    add_page_heading(doc, "Exploring the Website's Web Pages", level=1)
    add_body_paragraph(doc, 
        "Below is a comprehensive visual walkthrough of all the key pages and features of the InterviewAI "
        "web application, demonstrating the complete user journey from landing to export.")
    
    # 1. Main Home Page
    add_page_heading(doc, "Home / Landing Page:", level=2)
    add_body_paragraph(doc, 
        "The landing page presents a premium, centered hero section with the application title, a brief tagline, "
        "and two call-to-action buttons — 'Let's Start' to navigate directly to the interview configuration form, "
        "and 'Explore More' to reveal additional information sections below. The header displays the InterviewAI "
        "branding with a 'Powered by Google Gemini' badge, and live statistics show 30+ questions, 3 types, and 4 levels.")
    add_screenshot(doc, "Images/Results/Main Home page.jpeg", 
                   "Main Landing Page — Premium centered hero view with CTA buttons and live statistics.")

    # 2. Exploring Home Page 1 (Simple Process)
    add_page_heading(doc, "Simple Process Section:", level=2)
    add_body_paragraph(doc, 
        "When the user clicks the 'Explore More' button, the lower information panels slide into view seamlessly. "
        "The 'Simple Process' section showcases a three-step workflow: (1) Fill in the details — enter job role, "
        "skills, and experience level; (2) AI generates questions — the system creates tailored technical, behavioral, "
        "and situational questions with follow-ups; (3) Review & export — browse questions by category, evaluate "
        "with criteria, and export to DOCX or PDF.")
    add_screenshot(doc, "Images/Results/explorimng home page1.jpeg", 
                   "Simple Process Section — Three-step workflow cards with animated reveal transition.")

    # 3. Exploring Home Page 2 (What You Get)
    add_page_heading(doc, "What You Get Section:", level=2)
    add_body_paragraph(doc, 
        "The 'What You Get' section presents a six-card feature grid highlighting the platform's core capabilities: "
        "AI-Powered Generation (Google Gemini crafts intelligent questions), Role-Specific Questions (matched to "
        "experience levels), Evaluation Framework (green flags, red flags, scoring scales), Export & Share (DOCX, PDF, "
        "clipboard), Session History (automatic saving and retrieval), and Instant Results (complete question sets "
        "generated in seconds).")
    add_screenshot(doc, "Images/Results/exploring home page-2.jpeg", 
                   "What You Get Section — Six-card feature grid displaying platform capabilities.")

    doc.add_page_break()

    # 4. Generate Questions Form Page
    add_page_heading(doc, "Configure Interview Form:", level=2)
    add_body_paragraph(doc, 
        "The configuration form page (Step 1 of 2) provides an intuitive card layout with the following input fields: "
        "Job Role (required text input), Required Skills (optional comma-separated list), Experience Level (dropdown "
        "selector with Beginner, Intermediate, Advanced, and Senior options), Questions count (interactive slider "
        "ranging from 5 to 30), Interview Type (toggle buttons for Technical, Mixed, and Behavioral), and Company "
        "Context (optional textarea for organizational details). The 'Generate Questions' button triggers the "
        "asynchronous API call with a loading spinner.")
    add_screenshot(doc, "Images/Results/Generate Questions page.jpeg", 
                   "Configure Interview Form — Responsive form card with sliders, dropdowns, and toggle buttons.")

    # 5. Results Dashboard
    add_page_heading(doc, "Results Dashboard:", level=2)
    add_body_paragraph(doc, 
        "The results dashboard features a two-column layout with a history sidebar on the left and the main "
        "results panel on the right. The top section displays statistics cards showing total questions, technical "
        "count, behavioral count, and situational count. Below, a tabbed navigation allows switching between "
        "Technical, Behavioral, Situational, and Evaluation views. Each question card displays the category badge, "
        "difficulty level, the question text, a follow-up prompt, and an evaluation tip.")
    add_screenshot(doc, "Images/Results/Dashboard.jpeg", 
                   "Results Dashboard — Two-column layout with history sidebar, stats cards, and question cards.")

    # 6. Evaluation Page
    add_page_heading(doc, "Evaluation Panel:", level=2)
    add_body_paragraph(doc, 
        "The Evaluation tab presents a structured assessment framework containing: Scoring Scale (1-5 rating from "
        "Unsatisfactory to Exceptional), Key Competencies (tagged pills showing skills to assess), Green Flags "
        "(positive candidate signals like 'Admits limits of knowledge', 'Demonstrates awareness of memory safety'), "
        "Red Flags (warning signs like 'Claims to know everything', 'Defensive when receiving feedback'), and "
        "Interviewer Tips (actionable guidance for conducting the interview effectively).")
    add_screenshot(doc, "Images/Results/evalution page.jpeg", 
                   "Evaluation Panel — Scoring scale, competencies, green flags, red flags, and interviewer tips.")

    # 7. Save as PDF Option
    add_page_heading(doc, "Print / Save as PDF:", level=2)
    add_body_paragraph(doc, 
        "The application supports browser-native printing via the 'Print / PDF' button. The print dialog displays "
        "a clean, optimized layout with proper margins, showing the evaluation criteria, key competencies, green "
        "flags, and red flags in a printer-friendly format. Users can save directly as PDF using their browser's "
        "built-in 'Save as PDF' destination option.")
    add_screenshot(doc, "Images/Results/Save as PDF option page.jpeg", 
                   "Print / Save as PDF — Browser print dialog showing clean, optimized print layout.")

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    # CONCLUSION
    # ══════════════════════════════════════════════════════════════════════════
    add_page_heading(doc, "Conclusion", level=1)
    add_body_paragraph(doc, 
        "InterviewAI is a highly professional, modern, and standard-compliant recruiter empowerment system. "
        "By integrating the speed of Google Gemini API with a sleek glassmorphism client and robust local JSON "
        "session tracking, it sets a premium standard for hiring tools. The system eliminates repetitive "
        "administrative preparation work, allowing HR professionals and technical managers to focus entirely "
        "on candidate engagement and quality assessment.")
    
    add_body_paragraph(doc, 
        "The platform demonstrates end-to-end full-stack development proficiency, covering secure API integration, "
        "RESTful backend design, modern CSS aesthetics, asynchronous JavaScript programming, and professional "
        "document compilation. It establishes a strong foundation for future enterprise features, including "
        "multi-user collaborative reviews, live audio candidate screening, and cloud-hosted deployment.")

    add_body_paragraph(doc, 
        "Key achievements of this project include: successful integration with Google Gemini API for instant "
        "AI-powered question generation, a complete evaluation framework with scoring scales and behavioral flags, "
        "persistent session history with CRUD management, professional Word document export capabilities, "
        "a premium dark-mode glassmorphism user interface, and responsive design optimized for desktop and mobile devices.")

    # ── Save compiled document ─────────────────────────────────────────────
    output_filename = "InterviewAI_Documentation.docx"
    doc.save(output_filename)
    print(f"[SUCCESS] Documentation compiled and saved at: {output_filename}")

if __name__ == "__main__":
    create_document()
