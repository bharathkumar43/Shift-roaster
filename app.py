import calendar
from datetime import date

from flask import (Flask, render_template, request, session,
                   redirect, url_for, send_file, flash)

from roster_engine import (generate_roster, get_roster_summary,
                           SHIFTS, DAY_NAMES, CONTENT_TYPES)
from excel_export import generate_excel
from project_engine import generate_project_coverage
import database as db

app = Flask(__name__)
app.secret_key = "roster-automation-secret-key-change-in-production"

db.init_db()


# ── Index / Employee Management ──────────────────────────

@app.route("/", methods=["GET"])
def index():
    employees = db.get_all_employees()
    all_projects = db.get_all_projects()
    today = date.today()
    month_names = [calendar.month_name[m] for m in range(1, 13)]
    return render_template("index.html",
                           employees=employees,
                           all_projects=all_projects,
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
    employees = db.get_all_employees()
    last = session.get("last_roster", {})
    year = last.get("year", date.today().year)
    month = last.get("month", date.today().month)

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

    month_name = calendar.month_name[month]
    filename = f"Shift_Roster_{month_name}_{year}.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
