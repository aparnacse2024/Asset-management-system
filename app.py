
from flask import Flask, render_template, request, redirect, url_for
from openpyxl import load_workbook
import os

app = Flask(__name__)
FILE = "data.xlsx"

def sheet_map(name):
    return {
        "assets": "Assets",
        "employees": "Employees",
        "transactions": "Transactions"
    }[name]

def get_sheet(name):
    wb = load_workbook(FILE)
    sheet = wb[name]
    data = list(sheet.values)
    headers = data[0]
    rows = data[1:]
    return wb, sheet, headers, rows

@app.route('/')
def dashboard():
    wb = load_workbook(FILE)
    counts = {
        "assets": wb["Assets"].max_row - 1,
        "employees": wb["Employees"].max_row - 1,
        "transactions": wb["Transactions"].max_row - 1
    }
    return render_template("dashboard.html", counts=counts)

@app.route('/chatbot')
def chatbot():
    return render_template("chat.html")

@app.route('/<name>')
def view(name):
    wb, sheet, headers, rows = get_sheet(sheet_map(name))

    search = request.args.get("search")
    if search:
        rows = [r for r in rows if search.lower() in str(r).lower()]

    assets = []
    employees = []
    if name == "transactions":
        _, _, _, assets = get_sheet("Assets")
        _, _, _, employees = get_sheet("Employees")

    return render_template("table.html",
                           name=name,
                           headers=headers,
                           rows=rows,
                           assets=assets,
                           employees=employees)

@app.route('/add/<name>', methods=["POST"])
def add(name):
    wb = load_workbook(FILE)
    sheet = wb[sheet_map(name)]
    sheet.append(list(request.form.values()))
    wb.save(FILE)
    return redirect(url_for('view', name=name))

@app.route('/delete/<name>/<int:row_id>')
def delete(name, row_id):
    wb = load_workbook(FILE)
    sheet = wb[sheet_map(name)]
    sheet.delete_rows(row_id + 2)
    wb.save(FILE)
    return redirect(url_for('view', name=name))

@app.route('/edit/<name>/<int:row_id>')
def edit(name, row_id):
    wb, sheet, headers, rows = get_sheet(sheet_map(name))
    return render_template("edit.html",
                           name=name,
                           headers=headers,
                           row=rows[row_id],
                           row_id=row_id)

@app.route('/update/<name>/<int:row_id>', methods=["POST"])
def update(name, row_id):
    wb = load_workbook(FILE)
    sheet = wb[sheet_map(name)]
    for col, value in enumerate(request.form.values(), start=1):
        sheet.cell(row=row_id + 2, column=col).value = value
    wb.save(FILE)
    return redirect(url_for('view', name=name))


@app.route('/chat', methods=['GET', 'POST'])
def chat():
    if request.method == "GET":
        return "Chatbot is running. Use POST request."

    message = request.form.get("message").lower()
    
    wb = load_workbook(FILE)

    if "how many assets" in message:
        sheet = wb["Assets"]
        count = sheet.max_row - 1
        return {"reply": f"Total assets are {count}"}

    return {"reply": "Sorry, I didn't understand that."}

if __name__ == "__main__":
    app.run(debug=True)