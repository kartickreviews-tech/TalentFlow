import os
import sqlite3
import re
import json
from datetime import datetime
from flask import Flask, render_template, request, jsonify

# Document Readers
try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

try:
    import docx
except ImportError:
    docx = None

app = Flask(__name__, template_folder='templates')
DB_NAME = 'talentflow.db'

def init_db():
    """Initializes SQLite database and seeds default job openings."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS jobs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            department TEXT NOT NULL
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS candidates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            email TEXT,
            phone TEXT,
            college TEXT,
            qualification TEXT,
            degree TEXT,
            cgpa REAL,
            experience REAL,
            skills TEXT,
            projects TEXT,
            certs TEXT,
            achievements TEXT,
            score INTEGER,
            status TEXT DEFAULT 'Review',
            evaluation_explanation TEXT,
            summary TEXT,
            avatar TEXT,
            initials TEXT,
            created_at TEXT,
            job_id INTEGER,
            status_history TEXT,
            FOREIGN KEY (job_id) REFERENCES jobs (id)
        )
    ''')
    
    cursor.execute('SELECT COUNT(*) FROM jobs')
    if cursor.fetchone()[0] == 0:
        default_jobs = [
            ("Cybersecurity Analyst", "Security"),
            ("Software Engineer", "Engineering"),
            ("Frontend Developer", "Engineering"),
            ("Data Scientist", "Data & AI")
        ]
        cursor.executemany('INSERT INTO jobs (title, department) VALUES (?, ?)', default_jobs)
        
    conn.commit()
    conn.close()

init_db()

def extract_text_from_file(file_path, filename):
    """Extracts raw plain text from PDF and DOCX resume files."""
    text = ""
    ext = filename.lower().split('.')[-1]
    
    if ext == 'pdf' and PdfReader:
        try:
            reader = PdfReader(file_path)
            for page in reader.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
        except Exception as e:
            print(f"Error reading PDF {filename}: {e}")
            
    elif ext in ['docx', 'doc'] and docx:
        try:
            doc = docx.Document(file_path)
            for para in doc.paragraphs:
                if para.text.strip():
                    text += para.text + "\n"
            for table in doc.tables:
                for row in table.rows:
                    text += " | ".join([cell.text.strip() for cell in row.cells if cell.text.strip()]) + "\n"
        except Exception as e:
            print(f"Error reading DOCX {filename}: {e}")
            
    else:
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                text = f.read()
        except Exception as e:
            print(f"Error reading plain text {filename}: {e}")
            
    return text

def parse_resume_content(text, filename):
    """Parses real candidate details using regex heuristics."""
    clean_lines = [l.strip() for l in text.split('\n') if l.strip()]
    
    # 1. EXTRACT EMAIL
    email_match = re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b', text)
    email = email_match.group(0) if email_match else "Not provided"
    
    # 2. EXTRACT PHONE NUMBER
    phone_match = re.search(r'(?:(?:\+?\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4})', text)
    phone = phone_match.group(0) if phone_match else "Not provided"
    
    # 3. EXTRACT CANDIDATE NAME
    name = ""
    for line in clean_lines[:5]:
        if '@' not in line and not re.search(r'\d{5,}', line) and len(line) < 40:
            if not re.search(r'\b(resume|cv|curriculum|profile|page)\b', line, re.IGNORECASE):
                name = line.strip()
                break
    if not name:
        base_filename = os.path.splitext(filename)[0]
        name = " ".join(base_filename.replace('_', ' ').replace('-', ' ').title().split()[:3])
        
    parts = name.split()
    initials = "".join([p[0].upper() for p in parts[:2]]) if parts else "C"
    
    # 4. EXTRACT CGPA / GPA / PERCENTAGE
    cgpa = None
    cgpa_patterns = [
        r'(?:cgpa|gpa|marks|grade)[:\s]*([0-9]\.[0-9]{1,2})\s*(?:\/\s*10|\/\s*4)?',
        r'([0-9]\.[0-9]{1,2})\s*(?:\/\s*10\.?0?|\/\s*4\.?0?)',
        r'\b([0-9]\.[0-9]{1,2})\s*cgpa\b',
        r'([7-9]\.[0-9]{1,2})\b'
    ]
    for pattern in cgpa_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            val = float(match.group(1))
            if val <= 10.0:
                cgpa = round(val, 2)
                break
                
    # 5. EXTRACT EXPERIENCE (Years)
    experience = 0.0
    exp_patterns = [
        r'(\d+(?:\.\d+)?)\s*(?:\+|\s*plus)?\s*(?:years|yrs|year)\s*(?:of)?\s*(?:experience|work|industry)',
        r'(?:experience|work history)[:\s]*(\d+(?:\.\d+)?)\s*(?:years|yrs)'
    ]
    for pattern in exp_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            experience = float(match.group(1))
            break
            
    if experience == 0.0:
        years = re.findall(r'\b(20[0-2][0-9]|19[9][0-9])\b', text)
        if len(years) >= 2:
            int_years = sorted([int(y) for y in years])
            calc_exp = int_years[-1] - int_years[0]
            if 0 < calc_exp <= 15:
                experience = float(calc_exp)

    # 6. EXTRACT COLLEGE / UNIVERSITY
    colleges_found = []
    college_keywords = ["university", "institute", "college", "school of", "academy", "polytechnic", "bits", "iit", "nit"]
    for line in clean_lines:
        if any(kw in line.lower() for kw in college_keywords) and len(line) < 80:
            colleges_found.append(line)
            break
    college = colleges_found[0] if colleges_found else "Not specified"

    # 7. EXTRACT DEGREE / QUALIFICATION
    degree_patterns = [
        r'\b(b\.?tech|m\.?tech|b\.?e|m\.?e|b\.?sc|m\.?sc|b\.?c\.?a|m\.?c\.?a|bachelor|master|phd|diploma)\b[^\n,.]*'
    ]
    degree = "Degree in progress"
    for pattern in degree_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            degree = match.group(0).strip().title()
            break

    # 8. EXTRACT TECHNICAL SKILLS
    skills_db = [
        "Python", "Java", "C++", "C#", "C", "JavaScript", "TypeScript", "HTML", "CSS", "SQL",
        "PostgreSQL", "MySQL", "MongoDB", "React", "Node.js", "Express", "Flask", "Django",
        "FastAPI", "Docker", "Kubernetes", "AWS", "Azure", "GCP", "Git", "GitHub", "Linux",
        "Bash", "Cybersecurity", "Network Security", "Penetration Testing", "Cryptography",
        "Wireshark", "Metasploit", "Machine Learning", "Deep Learning", "TensorFlow", "PyTorch",
        "Pandas", "NumPy", "Scikit-Learn", "REST API", "GraphQL", "CI/CD", "Agile"
    ]
    found_skills = []
    for skill in skills_db:
        if re.search(r'\b' + re.escape(skill) + r'\b', text, re.IGNORECASE):
            found_skills.append(skill)

    # 9. EXTRACT PROJECTS & CERTIFICATIONS
    projects = []
    certs = []
    in_project_section = False
    in_cert_section = False
    
    for line in clean_lines:
        lower = line.lower()
        if any(h in lower for h in ["projects", "personal projects", "academic projects"]):
            in_project_section = True
            in_cert_section = False
            continue
        elif any(h in lower for h in ["certifications", "certificates", "courses", "achievements"]):
            in_project_section = False
            in_cert_section = True
            continue
        elif any(h in lower for h in ["education", "experience", "skills", "summary"]):
            in_project_section = False
            in_cert_section = False
            
        if in_project_section and len(line) > 10 and len(projects) < 4:
            projects.append(line)
        elif in_cert_section and len(line) > 8 and len(certs) < 4:
            certs.append(line)

    # 10. COMPUTE REAL MATCH SCORE
    score_base = 50
    score_skills = min(35, len(found_skills) * 4)
    score_cgpa = (cgpa * 1.5) if cgpa else 10
    score_exp = min(15, experience * 3)
    match_score = min(99, max(45, int(score_base + score_skills + score_cgpa + score_exp - 15)))

    evaluation_explanation = (
        f"Parsed {len(found_skills)} matching skills ({', '.join(found_skills[:4]) if found_skills else 'General'}). "
        f"Extracted CGPA: {cgpa if cgpa else 'N/A'}, Experience: {experience} yrs."
    )
    
    summary = (
        f"Candidate extracted from {filename}. "
        f"Demonstrates skills in {', '.join(found_skills[:5]) if found_skills else 'various technical domains'}."
    )

    return {
        "name": name,
        "email": email,
        "phone": phone,
        "college": college,
        "qualification": "Higher Education",
        "degree": degree,
        "cgpa": cgpa,
        "experience": experience,
        "skills": found_skills,
        "projects": projects if projects else ["Extracted from uploaded resume document"],
        "certs": certs if certs else ["Extracted certifications"],
        "achievements": [],
        "score": match_score,
        "status": "Review",
        "evaluation_explanation": evaluation_explanation,
        "summary": summary,
        "avatar": "",
        "initials": initials,
        "created_at": datetime.now().isoformat(),
        "status_history": json.dumps([{"status": "Review", "at": datetime.now().isoformat()}])
    }

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/jobs', methods=['GET'])
def get_jobs():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM jobs')
    rows = cursor.fetchall()
    conn.close()
    return jsonify([dict(row) for row in rows])

@app.route('/api/candidates', methods=['GET'])
def get_candidates():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM candidates ORDER BY id DESC')
    rows = cursor.fetchall()
    conn.close()
    
    result = []
    for row in rows:
        d = dict(row)
        d['skills'] = json.loads(d['skills']) if d['skills'] else []
        d['projects'] = json.loads(d['projects']) if d['projects'] else []
        d['certs'] = json.loads(d['certs']) if d['certs'] else []
        d['achievements'] = json.loads(d['achievements']) if d['achievements'] else []
        d['status_history'] = json.loads(d['status_history']) if d['status_history'] else []
        result.append(d)
    return jsonify(result)

@app.route('/api/upload', methods=['POST'])
def upload_resumes():
    if 'files' not in request.files:
        return jsonify({"success": False, "error": "No files uploaded"}), 400
        
    files = request.files.getlist('files')
    job_id = request.form.get('job_id', 1)
    
    os.makedirs('uploads', exist_ok=True)
    processed_count = 0
    
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    for file in files:
        if file.filename == '':
            continue
        file_path = os.path.join('uploads', file.filename)
        file.save(file_path)
        
        raw_text = extract_text_from_file(file_path, file.filename)
        parsed = parse_resume_content(raw_text, file.filename)
        
        cursor.execute('''
            INSERT INTO candidates (
                name, email, phone, college, qualification, degree, cgpa, experience,
                skills, projects, certs, achievements, score, status, evaluation_explanation,
                summary, avatar, initials, created_at, job_id, status_history
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            parsed['name'], parsed['email'], parsed['phone'], parsed['college'],
            parsed['qualification'], parsed['degree'], parsed['cgpa'], parsed['experience'],
            json.dumps(parsed['skills']), json.dumps(parsed['projects']), json.dumps(parsed['certs']),
            json.dumps(parsed['achievements']), parsed['score'], parsed['status'],
            parsed['evaluation_explanation'], parsed['summary'], parsed['avatar'],
            parsed['initials'], parsed['created_at'], job_id, parsed['status_history']
        ))
        processed_count += 1
        
    conn.commit()
    conn.close()
    
    return jsonify({"success": True, "processed": processed_count})

@app.route('/api/candidates/<int:candidate_id>/status', methods=['PUT'])
def update_status(candidate_id):
    data = request.json
    new_status = data.get('status', 'Review')
    
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    cursor.execute('SELECT status_history FROM candidates WHERE id = ?', (candidate_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        return jsonify({"success": False, "error": "Candidate not found"}), 404
        
    history = json.loads(row['status_history']) if row['status_history'] else []
    history.append({"status": new_status, "at": datetime.now().isoformat()})
    
    cursor.execute('''
        UPDATE candidates SET status = ?, status_history = ? WHERE id = ?
    ''', (new_status, json.dumps(history), candidate_id))
    
    conn.commit()
    conn.close()
    
    return jsonify({"success": True})

if __name__ == '__main__':
    print("==================================================")
    print(" TalentFlow AI Resume Screening Server Running!")
    print(" Open your browser and go to: http://127.0.0.1:5000")
    print("==================================================")
    app.run(debug=True, port=5000)