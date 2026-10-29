"""Local Ministry decision-intelligence store for the MEMD proof of concept.

All seed records are explicitly synthetic and live in SQLite so changing a reading
changes the API and visualisations. The store is a local POC, not a Ministry feed.
"""
import json
import os
import sqlite3
import threading
import uuid
from contextlib import contextmanager
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

from .bc_client import DATA_DIR

DB = DATA_DIR / "memd.sqlite3"
UPLOAD_DIR = DATA_DIR / "uploaded_templates"

@contextmanager
def connect():
    DB.parent.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(DB)
    db.row_factory = sqlite3.Row
    try:
        yield db
        db.commit()
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()

def initialize():
    with connect() as db:
        db.executescript("""
        CREATE TABLE IF NOT EXISTS indicators(id INTEGER PRIMARY KEY, sector TEXT, metric TEXT, unit TEXT, period TEXT, value REAL, source TEXT, updated_at TEXT);
        CREATE TABLE IF NOT EXISTS assets(id TEXT PRIMARY KEY, name TEXT, type TEXT, district TEXT, latitude REAL, longitude REAL, capacity REAL, unit TEXT, status TEXT, risk TEXT);
        CREATE TABLE IF NOT EXISTS telemetry(id INTEGER PRIMARY KEY, asset_id TEXT, recorded_at TEXT, throughput REAL, pressure REAL, temperature REAL, vibration REAL, power_kw REAL, status TEXT, FOREIGN KEY(asset_id) REFERENCES assets(id));
        CREATE TABLE IF NOT EXISTS report_templates(id TEXT PRIMARY KEY, filename TEXT, uploaded_at TEXT, sheets_json TEXT, formula_count INTEGER, output_path TEXT);
        CREATE TABLE IF NOT EXISTS report_jobs(id TEXT PRIMARY KEY, filename TEXT NOT NULL, status TEXT NOT NULL, created_at TEXT NOT NULL, updated_at TEXT NOT NULL, result_json TEXT, error TEXT, upload_path TEXT);
        CREATE TABLE IF NOT EXISTS audit(id INTEGER PRIMARY KEY, at TEXT, action TEXT, detail TEXT);
        """)
        if db.execute("SELECT count(*) FROM indicators").fetchone()[0] == 0:
            now=datetime.now().isoformat(timespec="seconds")
            series={
              "minerals": [("Mineral production value", "UGX bn", [378,301,326,342]),("Non-tax mineral revenue", "UGX bn", [19.4,25.9,27.1,29.3]),("Active mineral licences", "licences", [112,126,139,151]),("Inspections completed", "%", [62,68,73,79])],
              "petroleum": [("Petroleum product imports", "million litres", [2100,2230,2380,2490]),("National fuel demand", "million litres", [2040,2180,2320,2460]),("Strategic stock cover", "days", [24,21,19,17]),("Pipeline throughput", "million litres", [380,420,455,478])],
              "energy": [("Installed generation capacity", "MW", [1890,1920,1980,2050]),("Electricity access", "%", [57,59,61,64]),("Distribution losses", "%", [17.8,18.1,17.3,16.9]),("Renewable generation share", "%", [91,90,89,90])],
              "projects": [("Strategic projects on schedule", "%", [71,68,65,69]),("Projects with cost variance", "projects", [4,5,7,6]),("Budget execution", "%", [64,72,75,81]),("Milestones due this quarter", "milestones", [18,21,24,22])]
            }
            periods=["2023/24", "2024/25", "2025 Q4", "2026 Q1"]
            for sector, rows in series.items():
                for metric, unit, values in rows:
                    for period,value in zip(periods,values):
                        db.execute("INSERT INTO indicators(sector,metric,unit,period,value,source,updated_at) VALUES(?,?,?,?,?,?,?)",(sector,metric,unit,period,value,"MEMD workspace",now))
            assets=[("PS-01","Kabaale Pump Station","Pump station","Hoima",1.435,31.343,120,"MW","Normal","Low"),("PS-02","Kikube Pump Station","Pump station","Kikuube",1.478,31.122,110,"MW","Normal","Low"),("PS-03","Masindi Pump Station","Pump station","Masindi",1.674,31.715,115,"MW","Watch","Medium"),("PS-04","Lambanga Terminal","Terminal","Masindi",1.825,31.895,90,"MW","Normal","Low"),("PS-05","Namwasa Pump Station","Pump station","Rakai",-0.936,31.092,105,"MW","Alert","High")]
            for a in assets: db.execute("INSERT INTO assets VALUES(?,?,?,?,?,?,?,?,?,?)",a)
            import random
            rng=random.Random(2701)
            for a in assets:
                aid=a[0]; base=82 if aid!="PS-05" else 64
                for k in range(30):
                    day=(datetime.now()-timedelta(days=29-k)).isoformat(timespec="seconds")
                    vib=round(rng.uniform(1.1,2.8)+(1.8 if aid=="PS-05" and k>20 else 0),2)
                    stat="Alert" if vib>3.8 else "Normal"
                    db.execute("INSERT INTO telemetry(asset_id,recorded_at,throughput,pressure,temperature,vibration,power_kw,status) VALUES(?,?,?,?,?,?,?,?)",(aid,day,round(base+rng.uniform(-5,5),1),round(7.8+rng.uniform(-.6,.6),2),round(54+rng.uniform(-3,3),1),vib,round(2500+rng.uniform(-160,160),1),stat))
        db.commit()

def overview():
    with connect() as db:
        rows=[dict(r) for r in db.execute("SELECT sector,metric,unit,period,value,'MEMD workspace' AS source FROM indicators ORDER BY sector,metric, CASE period WHEN '2026 Q1' THEN 0 WHEN '2025 Q4' THEN 1 WHEN '2024/25' THEN 2 ELSE 3 END")]
        assets=[dict(r) for r in db.execute("SELECT a.*, (SELECT count(*) FROM telemetry t WHERE t.asset_id=a.id AND t.status='Alert') alerts FROM assets a")]
        series={}
        for r in rows: series.setdefault(r['sector'],{}).setdefault(r['metric'],[]).append({'period':r['period'],'value':r['value'],'unit':r['unit']})
        for k in series:
            for m in series[k]: series[k][m].sort(key=lambda x:x['period'])
        return {'indicators':rows,'series':series,'assets':assets,'asset_count':len(assets),'alert_count':sum(a['risk']=='High' for a in assets),'database':str(DB.name),'data_label':'MEMD workspace'}

def telemetry(asset_id):
    with connect() as db:
        return [dict(r) for r in db.execute("SELECT recorded_at,throughput,pressure,temperature,vibration,power_kw,status FROM telemetry WHERE asset_id=? ORDER BY recorded_at",(asset_id,))]

def set_reading(asset_id, reading):
    allowed={'throughput','pressure','temperature','vibration','power_kw'}
    if reading.get('metric') not in allowed: raise ValueError('Unsupported telemetry metric')
    value=float(reading.get('value'))
    if not -1000000 < value < 1000000: raise ValueError('Value out of range')
    with connect() as db:
        if not db.execute('SELECT 1 FROM assets WHERE id=?',(asset_id,)).fetchone(): raise KeyError('Asset not found')
        latest=db.execute('SELECT * FROM telemetry WHERE asset_id=? ORDER BY recorded_at DESC LIMIT 1',(asset_id,)).fetchone()
        values={k:latest[k] for k in allowed}; values[reading['metric']]=value
        status='Alert' if values['vibration']>=3.8 or values['pressure']<6.8 else 'Normal'
        db.execute("INSERT INTO telemetry(asset_id,recorded_at,throughput,pressure,temperature,vibration,power_kw,status) VALUES(?,?,?,?,?,?,?,?)",(asset_id,datetime.now().isoformat(timespec='seconds'),values['throughput'],values['pressure'],values['temperature'],values['vibration'],values['power_kw'],status))
        db.execute("UPDATE assets SET status=?,risk=? WHERE id=?",(status,'High' if status=='Alert' else 'Low',asset_id))
        db.execute("INSERT INTO audit(at,action,detail) VALUES(?,?,?)",(datetime.now().isoformat(timespec='seconds'),'telemetry_updated','{}: {}={}'.format(asset_id,reading['metric'],value)))
        db.commit()
    return {'asset_id':asset_id,'status':status,'reading':value}

def update_indicator(payload):
    try: value=float(payload.get('value'))
    except (TypeError,ValueError): raise ValueError('Enter a numeric value')
    if not -1e12 < value < 1e12: raise ValueError('Value out of range')
    with connect() as db:
        cur=db.execute("UPDATE indicators SET value=?,updated_at=? WHERE sector=? AND metric=? AND period=?",(value,datetime.now().isoformat(timespec='seconds'),payload.get('sector'),payload.get('metric'),payload.get('period')))
        if cur.rowcount==0: raise KeyError('Indicator not found')
        db.execute("INSERT INTO audit(at,action,detail) VALUES(?,?,?)",(datetime.now().isoformat(timespec='seconds'),'indicator_updated',f"{payload.get('sector')}/{payload.get('metric')}/{payload.get('period')}={value}"))
        db.commit()
    return {'ok':True,'value':value}

def answer(question):
    """Deterministic, evidence-bound POC analyst; no external model required."""
    data=overview(); q=question.lower(); sector=next((s for s in ('minerals','petroleum','energy','projects') if (s in q or (s=='minerals' and 'mineral' in q))),None)
    chosen=[r for r in data['indicators'] if not sector or r['sector']==sector]
    # If configured, Azure may explain only the exact local records supplied here.
    # The deterministic response below remains the offline fallback.
    if os.environ.get('AZURE_OPENAI_ENDPOINT'):
        try:
            from .command import _azure
            resp=_azure().chat.completions.create(
                model=os.environ.get('AZURE_OPENAI_DEPLOYMENT','gpt-5.4'),
                messages=[
                    {'role':'system','content':('You are the MEMD sector evidence analyst. Use only the supplied local database records. '
                     'Use only the supplied workspace records. State missing evidence. '
                     'Do not infer causes, legal non-compliance, safety status, or operational instructions. Human officials and asset operators decide. '
                     'Keep the answer concise, cite metric names and periods, and say when the evidence cannot answer the question.')},
                    {'role':'user','content':json.dumps({'question':question[:1200],'indicators':chosen,'pipeline_assets':data['assets'] if any(x in q for x in ('pipeline','pump','twin','asset')) else []},default=str)}
                ], max_completion_tokens=1200)
            answer=(resp.choices[0].message.content or '').strip()
            if answer: return {'mode':'azure-sector-agent','answer':answer,'evidence':chosen}
        except Exception as exc:
            azure_note=f' Azure analysis is temporarily unavailable ({str(exc)[:100]}); the local evidence response is shown.'
    else: azure_note=' Configure Azure OpenAI to enable natural-language evidence synthesis; the local evidence response is active.'
    if 'pipeline' in q or 'pump' in q or 'twin' in q:
        return {'mode':'sector-agent','answer':'The pipeline asset view contains five assets with daily telemetry history. PS-05 is flagged High risk; its latest readings should be reviewed with the operator before any operational conclusion.'+azure_note,'evidence':data['assets']}
    if not chosen: chosen=data['indicators']
    latest=[]
    for r in chosen:
        if r['period']=='2026 Q1': latest.append(r)
    return {'mode':'sector-agent','answer':f"I found {len(chosen)} indicator records across {len(set(r['sector'] for r in chosen))} sector(s) in the workspace database. Latest-period measures include: "+"; ".join(f"{r['metric']}: {r['value']:g} {r['unit']}" for r in latest[:8])+". "+azure_note,'evidence':chosen}

def analyze_asset(asset_id):
    """Explain one asset's recent telemetry without issuing operational instructions."""
    data=overview(); asset=next((x for x in data['assets'] if x['id']==asset_id),None)
    if not asset: raise KeyError('Asset not found')
    readings=telemetry(asset_id)[-14:]
    if os.environ.get('AZURE_OPENAI_ENDPOINT'):
        try:
            from .command import _azure
            resp=_azure().chat.completions.create(
                model=os.environ.get('AZURE_OPENAI_DEPLOYMENT','gpt-5.4-mini'),
                messages=[
                    {'role':'system','content':('You are an infrastructure monitoring analyst. Assess only the supplied asset metadata and telemetry. '
                     'Identify measured changes or threshold flags, cite values and timestamps, and state data limits. '
                     'Do not diagnose a physical cause, declare safety, or recommend control actions. Route operational decisions to the asset operator.')},
                    {'role':'user','content':json.dumps({'asset':asset,'telemetry_last_14':readings},default=str)}
                ],max_completion_tokens=1000)
            text=(resp.choices[0].message.content or '').strip()
            if text: return {'mode':'azure-twin-agent','answer':text,'asset':asset,'evidence':readings}
        except Exception:
            pass
    alerts=[r for r in readings if r['status']=='Alert']
    latest=readings[-1] if readings else {}
    return {'mode':'local-twin-agent','answer':f"{asset['name']} has {len(alerts)} alert-status reading(s) in the last {len(readings)} records. Latest telemetry: pressure {latest.get('pressure','—')} bar, vibration {latest.get('vibration','—')} mm/s, throughput {latest.get('throughput','—')} million litres/day. These readings identify a review point; they do not establish cause or safety status.",'asset':asset,'evidence':readings}

def _ai_report_matches(templates, indicators):
    """Map report labels to known latest-period indicators; the caller validates every field."""
    if not os.environ.get('AZURE_OPENAI_ENDPOINT'):
        return []
    try:
        from .command import _azure
        latest=[{'sector':r['sector'],'metric':r['metric'],'unit':r['unit'],'period':r['period']}
                for r in indicators if r['period']=='2026 Q1']
        sheets=[{'name':s['name'],'labels':s['labels'][:60]} for s in templates[:15]]
        resp=_azure().chat.completions.create(
            model=os.environ.get('AZURE_OPENAI_DEPLOYMENT','gpt-5.4-mini'),
            response_format={'type':'json_object'},
            messages=[
                {'role':'system','content':('You map workbook row labels to a fixed indicator catalogue. Return JSON only as {"matches":[{"sheet":"...","label_cell":"A1","metric":"...","period":"2026 Q1","confidence":0.0}]}. '
                 'Only return clear semantic matches; do not infer a metric from formulas, invent indicators, or map headings and totals. Use only catalogue metric names and periods.')},
                {'role':'user','content':json.dumps({'sheets':sheets,'catalogue':latest},default=str)}
            ],max_completion_tokens=2500)
        payload=json.loads(resp.choices[0].message.content or '{}')
        return payload.get('matches',[]) if isinstance(payload,dict) else []
    except Exception:
        return []

def inspect_and_generate(file_storage):
    if not file_storage or not file_storage.filename: raise ValueError('Choose an .xlsx report template')
    name=Path(file_storage.filename).name
    if not name.lower().endswith('.xlsx'): raise ValueError('POC template reproduction currently accepts .xlsx files')
    UPLOAD_DIR.mkdir(parents=True,exist_ok=True)
    source=UPLOAD_DIR/f'{uuid.uuid4().hex}_source.xlsx'
    file_storage.save(source)
    try: return _inspect_and_generate_path(source,name)
    finally:
        try: source.unlink(missing_ok=True)
        except OSError: pass

def _inspect_and_generate_path(source,name):
    try: wb=load_workbook(source)
    except Exception as e: raise ValueError(f'Workbook could not be read: {str(e)[:180]}')
    templates=[]; formula_count=0
    for ws in wb.worksheets:
        formulas=[]; labels=[]; style_samples=[]
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type=='f': formulas.append({'cell':cell.coordinate,'formula':cell.value}); formula_count+=1
                if cell.value is not None and cell.data_type!='f' and len(labels)<60: labels.append({'cell':cell.coordinate,'value':str(cell.value)[:120]})
                if cell.value is not None and cell.has_style and len(style_samples)<30:
                    style_samples.append({'cell':cell.coordinate,'fill':cell.fill.fgColor.rgb or cell.fill.fgColor.indexed,'font_color':cell.font.color.rgb if cell.font.color and cell.font.color.type=='rgb' else None,'bold':cell.font.bold,'number_format':cell.number_format})
        templates.append({'name':ws.title,'rows':ws.max_row,'columns':ws.max_column,'formulas':formulas[:300],'labels':labels,'styles':style_samples})
    # Preserve sheet layouts, formulas and cell formatting. Add a data-backed indicator sheet;
    # explicit {{METRIC:...}} tokens in the uploaded template are replaced where found.
    indicators=overview()['indicators']; mapping={}; matched=[]
    ai_suggestions=_ai_report_matches(templates,indicators)
    latest_by_key={(r['metric'],r['period']):r for r in indicators}
    valid_labels={(s['name'],label['cell']):label['value'] for s in templates for label in s['labels']}
    ai_by_label={}
    for suggestion in ai_suggestions:
        if not isinstance(suggestion,dict): continue
        key=(suggestion.get('sheet'),suggestion.get('label_cell'))
        metric=(suggestion.get('metric'),suggestion.get('period'))
        try: confidence=float(suggestion.get('confidence',0))
        except (TypeError,ValueError): confidence=0
        if key in valid_labels and metric in latest_by_key and confidence>=0.8:
            ai_by_label[key]=latest_by_key[metric]
    for r in indicators:
        mapping[f"{{{{{r['metric'].upper().replace(' ','_')}}}}}"]=r['value']
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value,str) and cell.value.strip().upper() in mapping: cell.value=mapping[cell.value.strip().upper()]
                if isinstance(cell.value,str):
                    normalized=' '.join(cell.value.casefold().split())
                    metric=next((r for r in indicators if ' '.join(r['metric'].casefold().split())==normalized and r['period']=='2026 Q1'),None)
                    metric=metric or ai_by_label.get((ws.title,cell.coordinate))
                    target=ws.cell(cell.row,cell.column+1) if cell.column<ws.max_column+2 else None
                    if metric and target is not None and (target.value is None or (isinstance(target.value,str) and target.value.startswith('{{'))):
                        target.value=metric['value']; matched.append({'sheet':ws.title,'label_cell':cell.coordinate,'value_cell':target.coordinate,'metric':metric['metric'],'period':metric['period'],'match_method':'AI-assisted' if (ws.title,cell.coordinate) in ai_by_label else 'exact label'})
    ws=wb.create_sheet('MEMD data extract')
    ws.append(['Sector','Indicator','Period','Value','Unit','Source classification'])
    for cell in ws[1]: cell.font=Font(bold=True,color='FFFFFF'); cell.fill=PatternFill('solid',fgColor='155EEF')
    for r in indicators: ws.append([r['sector'].title(),r['metric'],r['period'],r['value'],r['unit'],'MEMD workspace'])
    ws.freeze_panes='A2'; ws.auto_filter.ref=ws.dimensions
    for col,width in {'A':16,'B':36,'C':16,'D':14,'E':24,'F':24}.items(): ws.column_dimensions[col].width=width
    UPLOAD_DIR.mkdir(parents=True,exist_ok=True); rid=uuid.uuid4().hex[:12]; path=UPLOAD_DIR/f'{rid}_MEMD_report.xlsx'; wb.save(path)
    with connect() as db:
        db.execute('INSERT INTO report_templates VALUES(?,?,?,?,?,?)',(rid,name,datetime.now().isoformat(timespec='seconds'),json.dumps(templates),formula_count,str(path)))
        db.execute('INSERT INTO audit(at,action,detail) VALUES(?,?,?)',(datetime.now().isoformat(timespec='seconds'),'report_template_uploaded',f'{name}: {formula_count} formulas scanned'))
        db.commit()
    match_note='AI-assisted label mapping uses Azure OpenAI and only fills high-confidence matches to known indicators.' if ai_by_label else 'Exact indicator labels in a row and {{INDICATOR_NAME}} tokens are populated from the latest database period.'
    wb.close()
    return {'id':rid,'filename':name,'sheets':templates,'formula_count':formula_count,'matched_cells':matched,'download_url':f'/api/memd/reports/{rid}/download','notes':['Uploaded worksheets, formulas, and cell formatting are retained.',match_note,'The generated workbook includes a MEMD data extract sheet. Formula recalculation, charts, pivot tables, complex layouts, and ambiguous label mapping require review in Excel.']}

def create_report_job(file_storage):
    if not file_storage or not file_storage.filename: raise ValueError('Choose an .xlsx report template')
    name=Path(file_storage.filename).name
    if not name.lower().endswith('.xlsx'): raise ValueError('Report reproduction currently accepts .xlsx files')
    UPLOAD_DIR.mkdir(parents=True,exist_ok=True)
    job_id=uuid.uuid4().hex[:16]
    source=UPLOAD_DIR/f'{job_id}_source.xlsx'
    file_storage.save(source)
    now=datetime.now().isoformat(timespec='seconds')
    with connect() as db:
        db.execute('INSERT INTO report_jobs(id,filename,status,created_at,updated_at,upload_path) VALUES(?,?,?,?,?,?)',
                   (job_id,name,'queued',now,now,str(source)))
    threading.Thread(target=_run_report_job,args=(job_id,),daemon=True,name=f'memd-report-{job_id}').start()
    return {'job_id':job_id,'filename':name,'status':'queued','status_url':f'/api/memd/reports/jobs/{job_id}'}

def _run_report_job(job_id):
    with connect() as db:
        row=db.execute('SELECT filename,upload_path FROM report_jobs WHERE id=?',(job_id,)).fetchone()
        if not row: return
        name,path=row['filename'],Path(row['upload_path'])
        db.execute("UPDATE report_jobs SET status='processing',updated_at=? WHERE id=?",(datetime.now().isoformat(timespec='seconds'),job_id))
    try:
        result=_inspect_and_generate_path(path,name)
        with connect() as db:
            db.execute("UPDATE report_jobs SET status='ready',updated_at=?,result_json=?,upload_path=NULL WHERE id=?",
                       (datetime.now().isoformat(timespec='seconds'),json.dumps(result),job_id))
    except Exception as exc:
        with connect() as db:
            db.execute("UPDATE report_jobs SET status='failed',updated_at=?,error=?,upload_path=NULL WHERE id=?",
                       (datetime.now().isoformat(timespec='seconds'),str(exc)[:500],job_id))
    finally:
        try: path.unlink(missing_ok=True)
        except OSError: pass

def report_job(job_id):
    with connect() as db:
        row=db.execute('SELECT id,filename,status,created_at,updated_at,result_json,error FROM report_jobs WHERE id=?',(job_id,)).fetchone()
    if not row: return None
    job=dict(row)
    job['result']=json.loads(job.pop('result_json')) if job.get('result_json') else None
    return job

def download_template(rid):
    with connect() as db:
        row=db.execute('SELECT output_path FROM report_templates WHERE id=?',(rid,)).fetchone()
    if not row: return None
    path=Path(row['output_path'])
    return path if path.exists() else None

def report_history():
    with connect() as db:
        return [dict(r) for r in db.execute("SELECT id,filename,uploaded_at,formula_count FROM report_templates ORDER BY uploaded_at DESC LIMIT 50")]

initialize()
