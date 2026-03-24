import calendar
from datetime import date
from io import BytesIO

from flask import (Flask, render_template, request, session,
                   redirect, url_for, send_file, flash, jsonify)

from roster_engine import (generate_roster, get_roster_summary,
                           SHIFTS, DAY_NAMES, CONTENT_TYPES)
from excel_export import generate_excel
from project_engine import generate_project_coverage
from file_parser import parse_file
import database as db

app = Flask(__name__)
app.secret_key = "roster-automation-secret-key-change-in-production"

APP_NAME = "Employee Work Load Distribution"

db.init_db()


# ── Index / Employee Management ──────────────────────────

@app.route("/", methods=["GET"])
def index():
    employees = db.get_all_employees()
    all_projects = db.get_all_projects()
    saved_rosters = db.list_saved_rosters()
    today = date.today()
    month_names = [calendar.month_name[m] for m in range(1, 13)]
    imported = request.args.get("imported", type=int)
    if imported is not None:
        flash(f"Successfully imported {imported} employee(s).", "success")
    return render_template("index.html",
                           app_name=APP_NAME,
                           employees=employees,
                           all_projects=all_projects,
                           saved_rosters=saved_rosters,
                           day_names=DAY_NAMES,
                           content_types=CONTENT_TYPES,
                           now_month=today.month,
                           now_year=today.year,
                           month_names=month_names)


@app.route("/add_employee", methods=["POST"])
def add_employee():
    name = request.form.get("name", "").strip()
    content_types = request.form.getlist("content_types")
    working_days = request.form.getlist("working_days")
    project_names = request.form.getlist("project_name")
    project_types = request.form.getlist("project_type")

    if not name or not content_types or len(working_days) == 0:
        flash("Please fill in all required fields.", "danger")
        return redirect(url_for("index"))

    emp_id = db.add_employee(name, content_types, working_days)
    if emp_id is None:
        flash(f"Employee '{name}' already exists.", "warning")
        return redirect(url_for("index"))

    for pname, ptype in zip(project_names, project_types):
        pname = pname.strip()
        if pname and ptype:
            db.add_project(pname, ptype, emp_id)

    flash(f"Employee '{name}' added successfully.", "success")
    return redirect(url_for("index"))


@app.route("/edit_employee/<int:emp_id>", methods=["POST"])
def edit_employee(emp_id):
    data = request.get_json()
    if not data:
        return jsonify({"error": "Invalid data"}), 400
    content_types = data.get("content_types") or ["Content"]
    working_days = data.get("working_days") or ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    db.update_employee(emp_id, content_types, working_days)
    # Replace projects
    db.clear_projects_for_employee(emp_id)
    for proj in (data.get("projects") or []):
        name = (proj.get("name") or "").strip()
        ptype = proj.get("product_type") or "Content"
        if name:
            db.add_project(name, ptype, emp_id)
    return jsonify({"success": True})


@app.route("/remove_employee/<int:emp_id>", methods=["POST"])
def remove_employee(emp_id):
    db.remove_employee(emp_id)
    return redirect(url_for("index"))


@app.route("/clear_all", methods=["POST"])
def clear_all():
    db.clear_all_employees()
    flash("All employees cleared.", "success")
    return redirect(url_for("index"))


# ── Roster Generation ────────────────────────────────────

@app.route("/generate", methods=["POST"])
def generate():
    employees = db.get_all_employees()
    if len(employees) < 2:
        flash("Add at least 2 employees to generate a roster.", "warning")
        return redirect(url_for("index"))

    year = int(request.form.get("year", date.today().year))
    month = int(request.form.get("month", date.today().month))

    night_counts = db.get_night_shift_counts()
    roster, warnings, shift_assignments = generate_roster(
        employees, year, month, night_counts
    )

    db.save_all_rotations(shift_assignments, employees, year, month)

    all_projects = db.get_all_projects()
    proj_coverage, proj_warnings = generate_project_coverage(
        all_projects, employees, shift_assignments, year, month
    )

    excel_output = generate_excel(
        roster, warnings, shift_assignments, employees, year, month,
        proj_coverage, proj_warnings
    )
    db.save_roster_excel(year, month, excel_output.read())
    excel_output.seek(0)

    summary = get_roster_summary(employees, shift_assignments)
    month_name = calendar.month_name[month]
    num_days = calendar.monthrange(year, month)[1]

    roster_data = []
    for day in range(1, num_days + 1):
        d = date(year, month, day)
        weekday = d.weekday()
        day_info = {
            "date": d.strftime("%b %d"),
            "day_name": DAY_NAMES[weekday],
            "day_abbr": d.strftime("%a"),
            "day_num": day,
            "shifts": {}
        }
        for shift_num in [1, 2, 3]:
            emp_names = roster[d][shift_num]
            emp_details = []
            for name in emp_names:
                emp = next(e for e in employees if e["name"] == name)
                emp_details.append({
                    "name": name,
                    "content_types": emp["content_types"]
                })
            day_info["shifts"][shift_num] = emp_details
        roster_data.append(day_info)

    session["last_roster"] = {"year": year, "month": month}

    return render_template("roster.html",
                           app_name=APP_NAME,
                           roster_data=roster_data,
                           warnings=warnings,
                           summary=summary,
                           shifts=SHIFTS,
                           month_name=month_name,
                           year=year,
                           month=month)


# ── Project Coverage ─────────────────────────────────────

@app.route("/projects", methods=["POST"])
def projects():
    employees = db.get_all_employees()
    all_projects = db.get_all_projects()

    if not employees or not all_projects:
        flash("Add employees and projects first.", "warning")
        return redirect(url_for("index"))

    year = int(request.form.get("year", date.today().year))
    month = int(request.form.get("month", date.today().month))

    night_counts = db.get_night_shift_counts()
    _, _, shift_assignments = generate_roster(employees, year, month, night_counts)

    coverage, proj_warnings = generate_project_coverage(
        all_projects, employees, shift_assignments, year, month
    )

    month_name = calendar.month_name[month]
    session["last_roster"] = {"year": year, "month": month}

    return render_template("projects.html",
                           app_name=APP_NAME,
                           coverage=coverage,
                           warnings=proj_warnings,
                           month_name=month_name,
                           year=year,
                           month=month,
                           projects=all_projects,
                           shifts=SHIFTS,
                           shift_assignments=shift_assignments)


# ── Download Excel ───────────────────────────────────────

@app.route("/download", methods=["GET"])
def download():
    year = request.args.get("year", type=int)
    month = request.args.get("month", type=int)

    if not year or not month:
        last = session.get("last_roster", {})
        year = last.get("year", date.today().year)
        month = last.get("month", date.today().month)

    saved = db.get_saved_roster(year, month)
    if saved:
        output = BytesIO(saved)
        month_name = calendar.month_name[month]
        filename = f"Workload_{month_name}_{year}.xlsx"
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    employees = db.get_all_employees()
    night_counts = db.get_night_shift_counts()
    roster, warnings, shift_assignments = generate_roster(
        employees, year, month, night_counts
    )

    all_projects = db.get_all_projects()
    proj_coverage, proj_warnings = generate_project_coverage(
        all_projects, employees, shift_assignments, year, month
    )

    output = generate_excel(
        roster, warnings, shift_assignments, employees, year, month,
        proj_coverage, proj_warnings
    )

    db.save_roster_excel(year, month, output.read())
    output.seek(0)

    month_name = calendar.month_name[month]
    filename = f"Workload_{month_name}_{year}.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ── File Upload ──────────────────────────────────────────

@app.route("/preview_upload", methods=["POST"])
def preview_upload():
    """Parse uploaded file and return employee list as JSON for the preview/edit modal."""
    if "file" not in request.files:
        return jsonify({"error": "No file selected"}), 400
    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "No file selected"}), 400
    try:
        employees = parse_file(file)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    return jsonify({"employees": employees})


@app.route("/upload_confirm", methods=["POST"])
def upload_confirm():
    """Receive the (possibly edited) employee list as JSON and save to DB."""
    data = request.get_json()
    if not data or "employees" not in data:
        return jsonify({"error": "Invalid data"}), 400
    employees = data["employees"]
    if not employees:
        return jsonify({"error": "No employees provided"}), 400
    db.clear_all_employees()
    count = 0
    for emp in employees:
        name = (emp.get("name") or "").strip()
        if not name:
            continue
        content_types = emp.get("content_types") or ["Content"]
        working_days = emp.get("working_days") or ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        emp_id = db.add_employee(name, content_types, working_days)
        if emp_id:
            count += 1
            for proj in (emp.get("projects") or []):
                proj_name = (proj.get("name") or "").strip()
                proj_type = proj.get("product_type") or "Content"
                if proj_name:
                    db.add_project(proj_name, proj_type, emp_id)
    return jsonify({"success": True, "count": count})


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        flash("No file selected.", "danger")
        return redirect(url_for("index"))

    file = request.files["file"]
    if not file.filename:
        flash("No file selected.", "danger")
        return redirect(url_for("index"))

    try:
        employees = parse_file(file)
    except ValueError as e:
        flash(f"Upload failed: {e}", "danger")
        return redirect(url_for("index"))
    except Exception as e:
        flash(f"Error processing file: {e}", "danger")
        return redirect(url_for("index"))

    db.clear_all_employees()

    count = 0
    for emp in employees:
        emp_id = db.add_employee(emp["name"], emp["content_types"], emp["working_days"])
        if emp_id:
            count += 1

    flash(
        f"Successfully imported {count} employees from '{file.filename}'. "
        f"Shifts will be redistributed when you generate a roster.",
        "success"
    )
    return redirect(url_for("index"))


# ── Search API ───────────────────────────────────────────

@app.route("/search", methods=["GET"])
def search():
    q = request.args.get("q", "").strip()
    year = request.args.get("year", date.today().year, type=int)
    month = request.args.get("month", date.today().month, type=int)

    if not q or len(q) < 1:
        return jsonify({"employees": [], "projects": []})

    employees = db.get_all_employees()
    night_counts = db.get_night_shift_counts()

    shift_assignments = {}
    if employees:
        _, _, shift_assignments = generate_roster(employees, year, month, night_counts)

    emp_results = []
    matched_emps = db.search_employees(q)
    all_projects = db.get_all_projects()

    for emp in matched_emps:
        emp_projects = [p for p in all_projects if p["employee_id"] == emp["id"]]
        shift = shift_assignments.get(emp["name"])
        shift_info = None
        if shift:
            shift_info = {
                "number": shift,
                "name": SHIFTS[shift]["name"],
                "time_ist": SHIFTS[shift]["time_ist"],
                "time_est": SHIFTS[shift]["time_est"],
                "strength": SHIFTS[shift]["strength"],
            }

        num_days = calendar.monthrange(year, month)[1]
        working_count = 0
        off_count = 0
        for day in range(1, num_days + 1):
            d = date(year, month, day)
            day_name = DAY_NAMES[d.weekday()]
            if day_name in emp["working_days"]:
                working_count += 1
            else:
                off_count += 1

        emp_results.append({
            "name": emp["name"],
            "content_types": emp["content_types"],
            "working_days": emp["working_days"],
            "projects": [{"name": p["name"], "product_type": p["product_type"]}
                         for p in emp_projects],
            "shift": shift_info,
            "working_days_count": working_count,
            "off_days_count": off_count,
        })

    proj_results = []
    matched_projs = db.search_projects(q)
    for proj in matched_projs:
        owner = next((e for e in employees if e["name"] == proj["employee_name"]), None)
        shift = shift_assignments.get(proj["employee_name"])

        daily_handlers = []
        if owner:
            num_days = calendar.monthrange(year, month)[1]
            for day in range(1, num_days + 1):
                d = date(year, month, day)
                day_name = DAY_NAMES[d.weekday()]
                is_working = day_name in owner["working_days"]
                daily_handlers.append({
                    "day": day,
                    "date": d.strftime("%b %d"),
                    "day_abbr": d.strftime("%a"),
                    "handler": proj["employee_name"] if is_working else None,
                    "is_off": not is_working,
                })

        proj_results.append({
            "name": proj["name"],
            "product_type": proj["product_type"],
            "owner": proj["employee_name"],
            "shift": shift,
            "shift_name": SHIFTS[shift]["name"] if shift else None,
            "daily_handlers": daily_handlers,
        })

    return jsonify({"employees": emp_results, "projects": proj_results})


if __name__ == "__main__":
    app.run(debug=True, port=5050)
