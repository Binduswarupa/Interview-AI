from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from ai_service import generate_interview_questions, generate_docx
from dotenv import load_dotenv
import os
import history_service

load_dotenv()

app = Flask(__name__, static_folder="../frontend", static_url_path="/")
CORS(app)


# ── Serve Frontend ─────────────────────────────────────────────────────────
@app.route("/")
def index():
    return app.send_static_file("index.html")


# ── Health Check ───────────────────────────────────────────────────────────
@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "message": "Interview Generator API is running ✅"})


# ── Generate Questions ─────────────────────────────────────────────────────
@app.route("/api/generate", methods=["POST"])
def generate():
    data = request.get_json()

    job_role        = data.get("job_role", "").strip()
    skills          = data.get("skills", "").strip()
    experience_level = data.get("experience_level", "intermediate")
    interview_type  = data.get("interview_type", "mixed")
    num_questions   = int(data.get("num_questions", 10))
    company_context = data.get("company_context", "").strip()

    if not job_role:
        return jsonify({"error": "Job role is required."}), 400

    if num_questions < 1 or num_questions > 30:
        return jsonify({"error": "Number of questions must be between 1 and 30."}), 400

    try:
        result = generate_interview_questions(
            job_role=job_role,
            skills=skills,
            experience_level=experience_level,
            interview_type=interview_type,
            num_questions=num_questions,
            company_context=company_context,
        )
        
        # Save to persistent history
        metadata = {
            "job_role": job_role,
            "skills": skills,
            "experience_level": experience_level,
            "interview_type": interview_type,
            "num_questions": num_questions,
            "company_context": company_context
        }
        history_item = history_service.save_history_item(metadata, result)
        
        # Include reference ID in the result payload
        result["id"] = history_item["id"]
        return jsonify(result)
    except ValueError as e:
        return jsonify({"error": str(e)}), 401
    except Exception as e:
        return jsonify({"error": f"AI generation failed: {str(e)}"}), 500


# ── Export as DOCX ─────────────────────────────────────────────────────────
@app.route("/api/export", methods=["POST"])
def export():
    data = request.get_json()
    questions_data = data.get("questions_data", {})
    job_role       = data.get("job_role", "Interview")

    if not questions_data:
        return jsonify({"error": "No questions data provided."}), 400

    try:
        file_path = generate_docx(questions_data, job_role)
        safe_name = job_role.replace(" ", "_").replace("/", "-")
        return send_file(
            file_path,
            as_attachment=True,
            download_name=f"{safe_name}_Interview_Questions.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        return jsonify({"error": f"Export failed: {str(e)}"}), 500


# ── History Management ─────────────────────────────────────────────────────
@app.route("/api/history", methods=["GET"])
def get_history():
    try:
        return jsonify(history_service.get_history_list())
    except Exception as e:
        return jsonify({"error": f"Failed to retrieve history: {str(e)}"}), 500


@app.route("/api/history/<item_id>", methods=["GET"])
def get_history_item(item_id):
    try:
        item = history_service.get_history_item(item_id)
        if not item:
            return jsonify({"error": "History record not found"}), 404
        return jsonify(item)
    except Exception as e:
        return jsonify({"error": f"Failed to retrieve history item: {str(e)}"}), 500


@app.route("/api/history/<item_id>", methods=["DELETE"])
def delete_history_item(item_id):
    try:
        success = history_service.delete_history_item(item_id)
        if not success:
            return jsonify({"error": "History record not found"}), 404
        return jsonify({"status": "success", "message": "History item deleted"})
    except Exception as e:
        return jsonify({"error": f"Failed to delete history item: {str(e)}"}), 500


# ── Run ────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    print(f"\n[OK] Interview Generator API running at http://localhost:{port}\n")
    app.run(debug=True, port=port)
