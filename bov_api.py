import os, io, base64, traceback, json
from flask import Flask, request, jsonify
from openpyxl import load_workbook

app = Flask(__name__)
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "bov_template.xlsx")

@app.after_request
def add_cors(r):
    r.headers["Access-Control-Allow-Origin"] = "*"
    r.headers["Access-Control-Allow-Headers"] = "Content-Type"
    r.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS, GET"
    return r

@app.route("/test", methods=["GET"])
def test():
    try:
        wb = load_workbook(TEMPLATE_PATH)
        return jsonify({"status": "ok", "sheets": wb.sheetnames})
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)})

@app.route("/research-rents", methods=["POST","OPTIONS"])
def research_rents():
    if request.method == "OPTIONS": return "", 200
    try:
        d = request.get_json(force=True)
        parks = d.get("parks", [])
        if not parks:
            return jsonify({"success": False, "error": "No parks provided"}), 400

        api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        if not api_key:
            return jsonify({"success": False, "error": "ANTHROPIC_API_KEY not set"}), 500

        park_list = "\n".join([
            f"{i+1}. {p.get('name','')} — {p.get('address','')}"
            for i, p in enumerate(parks)
        ])

        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            system="""You are a manufactured housing market analyst with deep knowledge of US MHC lot rents by region. Estimate current lot rents for each park. Respond ONLY with a JSON array, no markdown, no explanation:
[{"index":1,"avg_rent":850,"min_rent":800,"max_rent":900,"spaces":120,"utility":"Sub-metered","source":"AI Estimate","confidence":"medium"}]
confidence: high=specific knowledge, medium=regional estimate, low=rough guess.""",
            messages=[{"role": "user", "content": f"Estimate lot rents for these manufactured home parks:\n\n{park_list}"}]
        )

        text = message.content[0].text
        s = text.find("[")
        e = text.rfind("]")
        if s < 0 or e < 0:
            raise Exception("No JSON array in response: " + text[:200])
        results = json.loads(text[s:e+1])
        return jsonify({"success": True, "results": results})

    except Exception as e:
        tb = traceback.format_exc()
        print("research_rents ERROR:", tb)
        return jsonify({"success": False, "error": str(e), "trace": tb}), 500

@app.route("/generate-bov", methods=["POST","OPTIONS"])
def generate_bov():
    if request.method == "OPTIONS": return "", 200
    try:
        d = request.get_json(force=True)
        if not d: return jsonify({"success": False, "error": "No JSON received"}), 400
        wb = load_workbook(TEMPLATE_PATH)
        wb.calculation.iterate = True
        wb.calculation.iterateCount = 100
        wb.calculation.iterateDelta = 0.001
        wb["Sales Comps"].sheet_state = "visible"
        wb["Rent Comps"].sheet_state = "visible"
        prop = str(d.get("propName","Property"))
        today = str(d.get("today",""))
        year = int(d.get("year", 2026))
        units = int(d.get("units", 0))
        occupied = int(d.get("occupied", units))
        cap_rate = float(d.get("capRate", 0.05))
        def sv(ws, addr, val):
            if val is not None and val != "":
                try: ws[addr].value = val
                except: pass
        ws = wb["BOV Summary"]
        sv(ws,"A3", prop + "  |  " + str(d.get("address","")) + "  |  " + str(units) + " Units  |  Confidential")
        sv(ws,"A4", "Prepared by: Northmarq  |  " + today)
        sv(ws,"C11", prop); sv(ws,"F11", units)
        sv(ws,"C12", "Manufactured Housing Community"); sv(ws,"F12", occupied)
        sv(ws,"C13", str(d.get("address",""))); sv(ws,"C14", str(d.get("rentRange","")))
        sv(ws,"C15", str(d.get("mgmt",""))); sv(ws,"F15", year); sv(ws,"F24", cap_rate)
        ws = wb["Income Statement"]
        sv(ws,"A3", prop + "  |  January - December " + str(year-1) + "  |  Accrual Basis")
        sv(ws,"A4", "Northmarq  |  " + today)
        for addr, key in [
            ("C7","lotRent"),("C8","storageFees"),("C9","appFees"),("C10","lateFees"),
            ("C11","concessions"),("C12","cableIncome"),("C13","miscIncome"),
            ("C15","gasBilled"),("C16","waterBilled"),("C17","sewerBilled"),
            ("C18","garbageBilled"),("C19","electricBilled"),("C27","gasCost"),
            ("C28","waterCost"),("C29","sewerCost"),("C30","electricCost"),("C31","garbageCost"),
            ("C34","advertising"),("C35","travelAuto"),("C36","pestControl"),("C37","landscaping"),
            ("C38","insurance"),("C39","mgrInsurance"),("C40","legalFees"),("C42","poolExpense"),
            ("C43","maintenance"),("C44","cleaning"),("C45","streetRepairs"),("C48","propertyTax"),
            ("C50","officeSupplies"),("C51","internet"),("C53","licensesDues"),
            ("C55","residentMgrSalary"),("C56","rmLabor"),("C57","management"),
            ("C58","payrollTax"),("C59","payrollProcessing"),
        ]:
            val = d.get(key)
            if val:
                try: sv(ws, addr, float(val))
                except: pass
        for row in [7,8,9,10,11,12,13,15,16,17,18,19,27,28,29,30,31,34,35,36,37,38,
                    39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59]:
            for col in ["D","E"]:
                try: ws[col + str(row)].value = None
                except: pass
        ws = wb["Sales Comps"]
        sv(ws,"B3", prop + "  |  Enter comps manually  |  " + today)
        for i, row_data in enumerate(d.get("salesComps",[])[:10]):
            for j, col in enumerate(["B","C","D","E","F","G","H","I","J","K"]):
                if j < len(row_data) and row_data[j] not in [None,"","nan"]:
                    val = row_data[j]
                    if col in ["D","F","G","I"]:
                        try: val = float(str(val).replace(",","").replace("$","").replace("%",""))
                        except: pass
                    try: ws[col + str(7+i)].value = val
                    except: pass
        ws = wb["Rent Comps"]
        sv(ws,"B3", prop + "  |  Enter rent comps manually  |  " + today)
        for i, row_data in enumerate(d.get("rentComps",[])[:10]):
            for j, col in enumerate(["B","C","D","E","F","G","H","I","J"]):
                if j < len(row_data) and row_data[j] not in [None,"","nan"]:
                    val = row_data[j]
                    if col in ["D","E","F","G","I"]:
                        try: val = float(str(val).replace(",","").replace("$","").replace("%",""))
                        except: pass
                    try: ws[col + str(7+i)].value = val
                    except: pass
        ws = wb["5-Year Cash Flow"]
        sv(ws,"A2", prop + "  |  Projected " + str(year) + " - " + str(year+4) + "  |  Capital Markets as of " + today)
        buf = io.BytesIO()
        wb.save(buf)
        return jsonify({"success": True, "b64": base64.b64encode(buf.getvalue()).decode()})
    except Exception as e:
        return jsonify({"success": False, "error": str(e), "trace": traceback.format_exc()}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
