import os, io, base64, traceback
from flask import Flask, request, jsonify
from openpyxl import load_workbook

app = Flask(__name__)
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "bov_template.xlsx")

@app.after_request
def add_cors(r):
    r.headers["Access-Control-Allow-Origin"] = "*"
    r.headers["Access-Control-Allow-Headers"] = "Content-Type"
    r.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
    return r

@app.route("/test", methods=["GET"])
def test():
    try:
        wb = load_workbook(TEMPLATE_PATH)
        return jsonify({"status": "ok", "sheets": wb.sheetnames})
    except Exception as e:
        return jsonify({"status": "error", "error": str(e), "trace": traceback.format_exc()})

@app.route("/generate-bov", methods=["POST","OPTIONS"])
def generate_bov():
    if request.method == "OPTIONS":
        return "", 200
    try:
        d = request.get_json(force=True)
        if not d:
            return jsonify({"success": False, "error": "No JSON received"}), 400

        wb = load_workbook(TEMPLATE_PATH)
        wb.calculation.iterate = True
        wb.calculation.iterateCount = 100
        wb.calculation.iterateDelta = 0.001
        wb["Sales Comps"].sheet_state = "visible"
        wb["Rent Comps"].sheet_state = "visible"

        prop     = str(d.get("propName","Property"))
        today    = str(d.get("today",""))
        year     = int(d.get("year", 2026))
        units    = int(d.get("units", 0))
        occupied = int(d.get("occupied", units))
        cap_rate = float(d.get("capRate", 0.05))

        def sv(ws, addr, val):
            if val is not None and val != "":
                try: ws[addr].value = val
                except: pass

        # BOV Summary
        ws = wb["BOV Summary"]
        sv(ws,"A3", prop + "  |  " + str(d.get("address","")) + "  |  " + str(units) + " Units  |  Confidential")
        sv(ws,"A4", "Prepared by: Northmarq  |  " + today)
        sv(ws,"C11", prop)
        sv(ws,"F11", units)
        sv(ws,"C12", "Manufactured Housing Community")
        sv(ws,"F12", occupied)
        sv(ws,"C13", str(d.get("address","")))
        sv(ws,"C14", str(d.get("rentRange","")))
        sv(ws,"C15", str(d.get("mgmt","")))
        sv(ws,"F15", year)
        sv(ws,"F24", cap_rate)

        # Income Statement
        ws = wb["Income Statement"]
        sv(ws,"A3", prop + "  |  January - December " + str(year-1) + "  |  Accrual Basis")
        sv(ws,"A4", "Northmarq  |  " + today)
        income_fields = [
            ("C7","lotRent"),("C8","storageFees"),("C9","appFees"),
            ("C10","lateFees"),("C11","concessions"),("C12","cableIncome"),
            ("C13","miscIncome"),("C15","gasBilled"),("C16","waterBilled"),
            ("C17","sewerBilled"),("C18","garbageBilled"),("C19","electricBilled"),
            ("C27","gasCost"),("C28","waterCost"),("C29","sewerCost"),
            ("C30","electricCost"),("C31","garbageCost"),
            ("C34","advertising"),("C35","travelAuto"),("C36","pestControl"),
            ("C37","landscaping"),("C38","insurance"),("C39","mgrInsurance"),
            ("C40","legalFees"),("C42","poolExpense"),("C43","maintenance"),
            ("C44","cleaning"),("C45","streetRepairs"),("C48","propertyTax"),
            ("C50","officeSupplies"),("C51","internet"),("C53","licensesDues"),
            ("C55","residentMgrSalary"),("C56","rmLabor"),("C57","management"),
            ("C58","payrollTax"),("C59","payrollProcessing"),
        ]
        for addr, key in income_fields:
            val = d.get(key)
            if val:
                try: sv(ws, addr, float(val))
                except: pass

        # Clear D and E columns
        rows_to_clear = [7,8,9,10,11,12,13,15,16,17,18,19,
                         27,28,29,30,31,34,35,36,37,38,39,40,41,42,43,
                         44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59]
        for row in rows_to_clear:
            for col in ["D","E"]:
                try: ws[col + str(row)].value = None
                except: pass

        # Sales Comps
        ws = wb["Sales Comps"]
        sv(ws,"B3", prop + "  |  Enter comps manually  |  " + today)
        sales_cols = ["B","C","D","E","F","G","H","I","J","K"]
        for i, row_data in enumerate(d.get("salesComps",[])[:10]):
            excel_row = 7 + i
            for j, col in enumerate(sales_cols):
                if j < len(row_data) and row_data[j] not in [None,"","nan"]:
                    val = row_data[j]
                    if col in ["D","F","G","I"]:
                        try: val = float(str(val).replace(",","").replace("$","").replace("%",""))
                        except: pass
                    try: ws[col + str(excel_row)].value = val
                    except: pass

        # Rent Comps
        ws = wb["Rent Comps"]
        sv(ws,"B3", prop + "  |  Enter rent comps manually  |  " + today)
        rent_cols = ["B","C","D","E","F","G","H","I","J"]
        for i, row_data in enumerate(d.get("rentComps",[])[:10]):
            excel_row = 7 + i
            for j, col in enumerate(rent_cols):
                if j < len(row_data) and row_data[j] not in [None,"","nan"]:
                    val = row_data[j]
                    if col in ["D","E","F","G","I"]:
                        try: val = float(str(val).replace(",","").replace("$","").replace("%",""))
                        except: pass
                    try: ws[col + str(excel_row)].value = val
                    except: pass

        # 5-Year Cash Flow
        ws = wb["5-Year Cash Flow"]
        sv(ws,"A2", prop + "  |  Projected " + str(year) + " - " + str(year+4) + "  |  Capital Markets as of " + today)

        buf = io.BytesIO()
        wb.save(buf)
        return jsonify({"success": True, "b64": base64.b64encode(buf.getvalue()).decode()})

    except Exception as e:
        tb = traceback.format_exc()
        print("ERROR:", tb)
        return jsonify({"success": False, "error": str(e), "trace": tb}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))