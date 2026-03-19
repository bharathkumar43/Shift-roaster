import json
import calendar
from datetime import date

from flask import Flask, render_template, request, session, redirect, url_for, send_file

from roster_engine import generate_roster, get_roster_summary, SHIFTS, DAY_NAMES, CONTENT_TYPES
from excel_export import generate_excel

app = Flask(__name__)
app.secret_key = "roster-automation-secret-key-change-in-production"


@app.route("/", methods=["GET"])
def index():
    employees = session.get("employees", [])
    today = date.today()
    month_names = [calendar.month_name[m] for m in range(1, 13)]
    return render_template("index.html",
                           employees=employees,
                           day_names=DAY_NAMES,
                           content_types=CONTENT_TYPES,
                           now_month=today.month,
                           now_year=today.year,
                           month_names=month_names)


@app.route("/add_employee", methods=["POST"])
def add_employee():
    name = request.form.get("name", "").strip()
    content_type = request.form.get("content_type", "")
    working_days = request.form.getlist("working_days")

    if not name or not content_type or len(working_days) == 0:
        return redirect(url_for("index"))

    employees = session.get("employees", [])

    if any(e["name"].lower() == name.lower() for e in employees):
        return redirect(url_for("index"))

    employees.append({
        "name": name,
        "content_type": content_type,
        "working_days": working_days,
    })
    session["employees"] = employees
    return redirect(url_for("index"))


@app.route("/remove_employee/<int:idx>", methods=["POST"])
def remove_employee(idx):
    employees = session.get("employees", [])
    if 0 <= idx < len(employees):
        employees.pop(idx)
        session["employees"] = employees
    return redirect(url_for("index"))


@app.route("/clear_all", methods=["POST"])
def clear_all():
    session.pop("employees", None)
    return redirect(url_for("index"))


@app.route("/generate", methods=["POST"])
def generate():
    employees = session.get("employees", [])
    if len(employees) < 2:
        return redirect(url_for("index"))

    year = int(request.form.get("year", date.today().year))
    month = int(request.form.get("month", date.today().month))

    roster, warnings, shift_assignments = generate_roster(employees, year, month)
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
                    "content_type": emp["content_type"]
                })
            day_info["shifts"][shift_num] = emp_details
        roster_data.append(day_info)

    session["last_roster"] = {
        "year": year,
        "month": month,
    }

    return render_template("roster.html",
                           roster_data=roster_data,
                           warnings=warnings,
                           summary=summary,
                           shifts=SHIFTS,
                           month_name=month_name,
                           year=year,
                           month=month)


@app.route("/download", methods=["GET"])
def download():
    employees = session.get("employees", [])
    last = session.get("last_roster", {})
    year = last.get("year", date.today().year)
    month = last.get("month", date.today().month)

    roster, warnings, shift_assignments = generate_roster(employees, year, month)
    output = generate_excel(roster, warnings, shift_assignments, employees, year, month)

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
