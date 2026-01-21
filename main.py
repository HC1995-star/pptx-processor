from flask import Flask, request, jsonify
import base64
from pptx import Presentation
from io import BytesIO

app = Flask(__name__)

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok"})

@app.route('/process-pptx', methods=['POST'])
def process_pptx():
    try:
        data = request.json
        pptx_b64 = data.get('pptx_binary')
        qbr = data.get('qbr_data', {})
        
        pptx_bytes = base64.b64decode(pptx_b64)
        prs = Presentation(BytesIO(pptx_bytes))
        
        def s(x):
            return "" if x is None else str(x)
        
        full_date = ""
        if qbr.get("Report Date") and qbr.get("Report Time"):
            full_date = f"{qbr.get('Report Date')} {qbr.get('Report Time')}"
        else:
            full_date = s(qbr.get("Report Date") or qbr.get("Report Time") or "")
        
        # Start with manual mappings for fields that need custom formatting
        replacements = {
            "{{brand_name}}": s(qbr.get("Client") or qbr.get("clientName") or ""),
            "{{period}}": s(qbr.get("Period") or qbr.get("period") or ""),
            "{{date}}": full_date,
            "{{prepared_by}}": s(qbr.get("preparedBy") or ""),
            "{{market}}": s(qbr.get("market") or ""),
            "{{clicks_recent}}": s(qbr.get("Current Clicks") or ""),
            "{{sales_recent}}": s(qbr.get("Current Sales") or ""),
            "{{cvr_recent}}": s(qbr.get("Current Conv Rate") or ""),
            "{{order_value_recent}}": s(qbr.get("Current Order Value") or ""),
            "{{aov_recent}}": s(qbr.get("Current AOV") or ""),
            "{{total_commission_recent}}": s(qbr.get("Current Commission") or ""),
            "{{roi_recent}}": s(qbr.get("Current ROI") or ""),
            "{{clicks_yoy_pct}}": s(qbr.get("YoY Clicks Change") or ""),
            "{{sales_yoy_pct}}": s(qbr.get("YoY Sales Change") or ""),
            "{{cvr_yoy_pct}}": s(qbr.get("YoY Conv Rate Change") or ""),
            "{{order_value_yoy_pct}}": s(qbr.get("YoY Order Value Change") or ""),
            "{{growth_driver_1}}": f"{s(qbr.get('Top Growth Publisher'))}: {s(qbr.get('Top Growth Amount'))} ({s(qbr.get('Top Growth YoY %'))})" 
                if (qbr.get("Top Growth Publisher") or qbr.get("Top Growth Amount")) else "",
            "{{decline_1}}": f"{s(qbr.get('Top Decline Publisher'))}: {s(qbr.get('Top Decline Amount'))} ({s(qbr.get('Top Decline YoY %'))})" 
                if (qbr.get("Top Decline Publisher") or qbr.get("Top Decline Amount")) else "",
            "{{opp_1}}": f"{s(qbr.get('Top Opportunity Domain'))} (Pos {s(qbr.get('Top Opportunity Position'))}, Score {s(qbr.get('Top Opportunity Score'))})" 
                if qbr.get("Top Opportunity Domain") else "",
            "{{supporting_note_or_context}}": f"Total opportunities identified: {s(qbr.get('Total Opportunities Identified'))}" 
                if qbr.get("Total Opportunities Identified") else "",
        }
        
        # NEW: Add default values for table placeholders (pointing to notes)
        replacements.update({
            "{{yoy_summary_table}}": "→ See detailed YoY summary table in presentation notes",
            "{{top_10_growth_table}}": "→ See top 10 growth publishers in presentation notes",
            "{{top_10_decline_table}}": "→ See top 10 declining publishers in presentation notes",
            "{{top_current_performers_table}}": "→ See top current performers in presentation notes",
            "{{search_visibility_table}}": "→ See visibility opportunities in presentation notes"
        })
        
        # Automatically add ALL other fields from qbr_data as placeholders
        # This handles all the AI-extracted fields like exec_point_1, insight_1, etc.
        for key, value in qbr.items():
            placeholder = f"{{{{{key}}}}}"
            # Only add if not already in replacements (don't override manual mappings)
            if placeholder not in replacements:
                replacements[placeholder] = s(value)
        
        def replace_in_shape(shape):
            if hasattr(shape, "text_frame"):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for key, value in replacements.items():
                            if key in run.text:
                                run.text = run.text.replace(key, value)
            
            if hasattr(shape, "table"):
                for row in shape.table.rows:
                    for cell in row.cells:
                        for key, value in replacements.items():
                            if key in cell.text:
                                cell.text = cell.text.replace(key, value)
            
            if hasattr(shape, "shapes"):
                for sub_shape in shape.shapes:
                    replace_in_shape(sub_shape)
        
        for slide in prs.slides:
            for shape in slide.shapes:
                replace_in_shape(shape)
        
        notes_parts = []
        if qbr.get("Complete QBR Report"):
            notes_parts.append(f"=== Complete QBR Report ===\n{qbr['Complete QBR Report']}")
        if qbr.get("Program Analysis (Full Text)"):
            notes_parts.append(f"=== Program Analysis ===\n{qbr['Program Analysis (Full Text)']}")
        if qbr.get("Publisher Analysis (Full Text)"):
            notes_parts.append(f"=== Publisher Analysis ===\n{qbr['Publisher Analysis (Full Text)']}")
        if qbr.get("Visibility Analysis (Full Text)"):
            notes_parts.append(f"=== Visibility Analysis ===\n{qbr['Visibility Analysis (Full Text)']}")
        
        if notes_parts and len(prs.slides) > 0:
            prs.slides[-1].notes_slide.notes_text_frame.text = "\n\n".join(notes_parts)
        
        output = BytesIO()
        prs.save(output)
        output.seek(0)
        result_b64 = base64.b64encode(output.read()).decode('utf-8')
        
        return jsonify({
            "success": True,
            "binary": result_b64,
            "filename": f"QBR-{s(qbr.get('clientName') or qbr.get('Client') or 'Client')}-{s(qbr.get('period') or qbr.get('Period') or 'Period')}.pptx"
        })
        
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8000)
