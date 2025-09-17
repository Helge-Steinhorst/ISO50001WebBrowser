import os
import json
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify, session
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import or_, func
import pandas as pd
from datetime import datetime, date, timedelta, time, timezone
from io import BytesIO
from PIL import Image
from fpdf import FPDF
import matplotlib
matplotlib.use('Agg') # <-- WICHTIG: Diese Zeile muss VOR dem Import von pyplot stehen
import matplotlib.pyplot as plt
import re


from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from PyPDF2 import PdfReader
import io
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.utils import ImageReader
from PIL import Image # Wichtig: Fügen Sie diesen Import hinzu


# --- App Konfiguration ---
app = Flask(__name__)
app.config['SECRET_KEY'] = 'dein_super_geheimer_schluessel_12345'
basedir = os.path.abspath(os.path.dirname(__file__))
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(basedir, 'zeiterfassung.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# --- Datenbankmodelle ---
class TimeEntry(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.Date, nullable=False, default=date.today)
    start_time = db.Column(db.Time, nullable=False)
    end_time = db.Column(db.Time, nullable=False)
    category = db.Column(db.String(50), nullable=False)
    project = db.Column(db.String(100), nullable=False)
    info_text = db.Column(db.String(300), nullable=True)

    @property
    def duration(self):
        start_dt = datetime.combine(self.date, self.start_time)
        end_dt = datetime.combine(self.date, self.end_time)
        if end_dt < start_dt:
            end_dt += timedelta(days=1)
        return end_dt - start_dt

    @property
    def duration_str(self):
        total_seconds = self.duration.total_seconds()
        hours, remainder = divmod(total_seconds, 3600)
        minutes, _ = divmod(remainder, 60)
        return f"{int(hours):02}:{int(minutes):02}"

class QuestionAnswer(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.String(50), nullable=False)
    category_index = db.Column(db.Integer, nullable=False, default=1)
    question = db.Column(db.String(500), nullable=False)
    options = db.Column(db.String(500), nullable=False)
    answer = db.Column(db.String(100), nullable=True)
    sort_index = db.Column(db.Integer, default=99)
    __table_args__ = (db.UniqueConstraint('category', 'category_index', 'question', name='_category_question_uc'),)


# --- Helper-Funktionen ---
def parse_voltage_string(voltage_str):
    if not isinstance(voltage_str, str): return []
    specs = []
    for variant in voltage_str.split(','):
        numbers = [int(n) for n in re.findall(r'(\d+)', variant)]
        if not numbers: continue
        min_v, max_v = (min(numbers), max(numbers))
        v_type = 'UNKNOWN'
        if 'ac' in variant.lower() and 'dc' in variant.lower(): v_type = 'AC/DC'
        elif 'ac' in variant.lower(): v_type = 'AC'
        elif 'dc' in variant.lower(): v_type = 'DC'
        specs.append({'min_v': min_v, 'max_v': max_v, 'type': v_type})
    return specs

def create_pdf_cover(pdf_obj, bearbeiter_name, title):
    pdf_obj.add_page()
    logo_path = os.path.join(basedir, 'static', 'img', 'logo.png')
    if os.path.exists(logo_path):
        pdf_obj.image(logo_path, x=pdf_obj.w/2 - 55, y=40, w=110)
    pdf_obj.set_y(150)
    pdf_obj.set_font('Arial', 'B', 24)
    pdf_obj.cell(0, 20, title, 0, 1, 'C')
    pdf_obj.set_font('Arial', '', 12)
    pdf_obj.cell(0, 10, f'Bearbeiter: {bearbeiter_name}', 0, 1, 'C')
    pdf_obj.cell(0, 10, f'Datum: {date.today().strftime("%d.%m.%Y")}', 0, 1, 'C')

# --- Kontext-Prozessor ---
@app.context_processor
def inject_now():
    return {'now': datetime.now(timezone.utc)}

# --- Routen / Unterseiten ---
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/Ablauf_Kunde/<string:filename>')
def Ablauf_Kunde(filename):
    try:
        filepath = os.path.join(basedir, 'visuals', filename)
        with open(filepath, 'r', encoding='utf-8') as f: xml_content = f.read()
        config = { "xml": xml_content, "background": "#ffffff", "toolbar": "top", "lightbox": False, "transparent": False }
        diagram_data = json.dumps(config)
        return render_template('Ablauf_Kunde.html', diagram_data=diagram_data, drawing_name=filename)
    except FileNotFoundError:
        flash(f"Zeichnung '{filename}' nicht gefunden.", "error")
        return redirect(url_for('index'))
    except Exception as e:
        flash(f"Fehler beim Laden der Zeichnung: {e}", "error")
        return redirect(url_for('index'))



@app.route('/Messung_Ablauf/<string:filename>')
def Messung_Ablauf(filename):
    try:
        filepath = os.path.join(basedir, 'visuals', filename)
        with open(filepath, 'r', encoding='utf-8') as f: xml_content = f.read()
        config = { "xml": xml_content, "background": "#ffffff", "toolbar": "top", "lightbox": False, "transparent": False }
        diagram_data = json.dumps(config)
        return render_template('Messung_Ablauf.html', diagram_data=diagram_data, drawing_name=filename)
    except FileNotFoundError:
        flash(f"Zeichnung '{filename}' nicht gefunden.", "error")
        return redirect(url_for('index'))
    except Exception as e:
        flash(f"Fehler beim Laden der Zeichnung: {e}", "error")
        return redirect(url_for('index'))


@app.route('/Messung_1_Aufbau/<string:filename>')
def Messung_1_Aufbau(filename):
    try:
        filepath = os.path.join(basedir, 'visuals', filename)
        with open(filepath, 'r', encoding='utf-8') as f: xml_content = f.read()
        config = { "xml": xml_content, "background": "#ffffff", "toolbar": "top", "lightbox": False, "transparent": False }
        diagram_data = json.dumps(config)
        return render_template('Messung_1_Aufbau.html', diagram_data=diagram_data, drawing_name=filename)
    except FileNotFoundError:
        flash(f"Zeichnung '{filename}' nicht gefunden.", "error")
        return redirect(url_for('index'))
    except Exception as e:
        flash(f"Fehler beim Laden der Zeichnung: {e}", "error")
        return redirect(url_for('index'))
    


@app.route('/dashboard-bilder')
def dashboard_bilder():
    """Diese Funktion rendert die Seite mit den Dashboard-Bildern."""
    return render_template('dashboard_bilder.html')


@app.route('/begriffsfinder', methods=['GET', 'POST'])
def begriffsfinder():
    search_term, result = "", ""
    if request.method == 'POST':
        search_term = request.form.get('search_term', '').strip()
        if not search_term:
            flash("Bitte geben Sie einen Suchbegriff ein.", "error")
        else:
            try:
                df = pd.read_excel('Daten.xlsx', sheet_name='Begriffe', header=None)
                search_area = df.iloc[13:]
                match = search_area[search_area[3].astype(str).str.lower() == search_term.lower()]
                if not match.empty:
                    explanation = match.iloc[0, 7]
                    result = str(explanation) if pd.notna(explanation) else "Keine Erklärung für diesen Begriff vorhanden."
                else:
                    result = "Begriff nicht gefunden."
            except FileNotFoundError:
                result = "Fehler: Die Datei 'Daten.xlsx' wurde nicht im Hauptverzeichnis gefunden."
            except Exception as e:
                result = f"Ein unerwarteter Fehler ist aufgetreten: {e}"
    return render_template('begriffsfinder.html', search_term=search_term, result=result)

@app.route('/autocomplete_begriffe')
def autocomplete_begriffe():
    query = request.args.get('q', '').lower()
    suggestions = []
    if query:
        try:
            df = pd.read_excel('Daten.xlsx', sheet_name='Begriffe', header=None)
            search_area = df.iloc[13:]
            matching_terms = search_area[search_area[3].astype(str).str.lower().str.startswith(query)].iloc[:, 3].unique()
            suggestions = matching_terms.tolist()[:10]
        except Exception:
            suggestions = []
    return jsonify(suggestions)

@app.route('/dokumentation', methods=['GET', 'POST'])
def dokumentation():
    if request.method == 'POST':
        # Code zum Erstellen eines neuen Eintrags (bleibt unverändert)
        try:
            date_obj = datetime.strptime(request.form.get('date'), '%Y-%m-%d').date()
            start_time_obj = datetime.strptime(request.form.get('start_time'), '%H:%M').time()
            end_time_obj = datetime.strptime(request.form.get('end_time'), '%H:%M').time()
            new_entry = TimeEntry(
                date=date_obj, start_time=start_time_obj, end_time=end_time_obj,
                category=request.form.get('category'), project=request.form.get('project'),
                info_text=request.form.get('info_text')
            )
            db.session.add(new_entry)
            db.session.commit()
            flash('Eintrag erfolgreich gespeichert!', 'success')
        except Exception as e:
            flash(f'Fehler beim Speichern: {e}', 'error')
        return redirect(url_for('dokumentation'))

    # --- KORREKTUR: Grafik wird bei JEDEM Aufruf neu generiert ---
    # 1. Alle Einträge aus der Datenbank laden
    entries = TimeEntry.query.order_by(TimeEntry.date.desc(), TimeEntry.start_time.desc()).all()
    
    # 2. Die Funktion aufrufen, die das Bild 'category_chart.png' erstellt
    generate_category_chart(entries) 
    
    # 3. Die Seite mit den Einträgen und der (jetzt aktuellen) Grafik anzeigen
    return render_template('dokumentation.html', entries=entries)

@app.route('/delete/<int:entry_id>', methods=['POST'])
def delete_entry(entry_id):
    entry_to_delete = TimeEntry.query.get_or_404(entry_id)
    try:
        db.session.delete(entry_to_delete)
        db.session.commit()
        flash('Eintrag wurde gelöscht.', 'success')
    except Exception as e:
        flash(f'Fehler beim Löschen des Eintrags: {e}', 'error')
    return redirect(url_for('dokumentation'))

@app.route('/delete_all_entries', methods=['POST'])
def delete_all_entries():
    try:
        # Lösche alle Einträge aus der TimeEntry Tabelle
        db.session.query(TimeEntry).delete()
        db.session.commit()
        flash('Alle Zeiterfassungseinträge wurden gelöscht.', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Fehler beim Löschen aller Einträge: {e}', 'error')
    
    # Leite zurück zur Dokumentationsseite (dies löst auch eine Neuberechnung der Grafik aus)
    return redirect(url_for('dokumentation'))

@app.route('/generate_pdf')
def generate_pdf():
    pass

@app.route('/fragen', methods=['GET', 'POST'])
def fragen():
    if 'setup_submit' in request.args:
        try:
            new_config = {
                'num_trafos': int(request.args.get('num_trafos', 0)),
                'num_einspeisungen': int(request.args.get('num_einspeisungen', 0)),
                'num_abgaenge': int(request.args.get('num_abgaenge', 0)),
            }
            old_config = session.get('project_config', {})
            for cat, key in [('Trafo', 'num_trafos'), ('Einspeisung', 'num_einspeisungen'), ('Abgang', 'num_abgaenge')]:
                old_count = old_config.get(key, 0)
                new_count = new_config.get(key, 0)
                if new_count > old_count:
                    master_questions = QuestionAnswer.query.filter_by(category=cat, category_index=1).all()
                    for i in range(old_count + 1, new_count + 1):
                        for master_q in master_questions:
                            if not QuestionAnswer.query.filter_by(category=cat, category_index=i, question=master_q.question).first():
                                db.session.add(QuestionAnswer(
                                    category=cat, category_index=i, question=master_q.question,
                                    options=master_q.options, sort_index=master_q.sort_index
                                ))
            session['project_config'] = new_config
            db.session.commit()
            flash('Projektkonfiguration wurde aktualisiert.', 'success')
        except Exception as e:
            db.session.rollback()
            flash(f"Fehler beim Konfigurieren: {e}", "error")
        return redirect(url_for('fragen'))

    if request.method == 'POST':
        category = request.form.get('category', 'Allgemein')
        if 'save_answers' in request.form:
            for key, value in request.form.items():
                if key.startswith('answer_') and (entry := QuestionAnswer.query.get(int(key.split('_')[1]))):
                    entry.answer = value
            db.session.commit()
            flash('Antworten erfolgreich gespeichert!', 'success')
        elif 'new_question' in request.form:
            try:
                project_config = session.get('project_config', {})
                num_instances = {
                    'Trafo': project_config.get('num_trafos', 0),
                    'Einspeisung': project_config.get('num_einspeisungen', 0),
                    'Abgang': project_config.get('num_abgaenge', 0)
                }.get(category, 1)
                
                max_index = db.session.query(func.max(QuestionAnswer.sort_index)).filter_by(category=category).scalar() or 0
                
                for i in range(1, num_instances + 1):
                    db.session.add(QuestionAnswer(
                        question=request.form.get('new_question'), options=request.form.get('options'),
                        category=category, category_index=i, sort_index=max_index + 1
                    ))
                db.session.commit()
                flash(f'Frage erfolgreich für "{category}" erstellt!', 'success')
            except Exception as e:
                db.session.rollback()
                flash(f'Fehler beim Erstellen der Frage: {e}', 'error')
        return redirect(url_for('fragen', _anchor=f"{category}-1", **session.get('project_config', {})))

    project_config = session.get('project_config', {})
    questions_db = QuestionAnswer.query.order_by(QuestionAnswer.category, QuestionAnswer.category_index, QuestionAnswer.sort_index).all()
    grouped_questions = {}
    for q in questions_db:
        key = (q.category, q.category_index)
        if key not in grouped_questions: grouped_questions[key] = []
        grouped_questions[key].append(q)
        
    return render_template('fragen.html', project_config=project_config, grouped_questions=grouped_questions)

@app.route('/reset_fragen_config')
def reset_fragen_config():
    session.pop('project_config', None)
    flash('Konfiguration zurückgesetzt. Bitte neu einrichten.', 'info')
    return redirect(url_for('fragen'))

@app.route('/delete_question/<int:question_id>', methods=['POST'])
def delete_question(question_id):
    q_ref = QuestionAnswer.query.get_or_404(question_id)
    try:
        QuestionAnswer.query.filter_by(question=q_ref.question, category=q_ref.category).delete()
        db.session.commit()
        flash(f'Frage wurde aus allen "{q_ref.category}"-Reitern gelöscht.', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Fehler beim Löschen der Frage: {e}', 'error')
    return redirect(url_for('fragen', _anchor=f"{q_ref.category}-1", **session.get('project_config', {})))

@app.route('/edit_question/<int:question_id>', methods=['POST'])
def edit_question(question_id):
    q_ref = QuestionAnswer.query.get_or_404(question_id)
    try:
        new_text = request.form.get(f'edited_question_text_{question_id}', '').strip()
        new_options = request.form.get(f'edited_options_{question_id}', '').strip()
        if not new_text or not new_options:
            flash("Fragetext und Antwortmöglichkeiten dürfen nicht leer sein.", "error")
        else:
            questions_to_update = QuestionAnswer.query.filter_by(question=q_ref.question, category=q_ref.category).all()
            for q in questions_to_update:
                q.question, q.options = new_text, new_options
            db.session.commit()
            flash('Frage erfolgreich für alle Instanzen aktualisiert.', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f"Fehler beim Aktualisieren der Frage: {e}", "error")
    return redirect(url_for('fragen', _anchor=f"{q_ref.category}-{q_ref.category_index}", **session.get('project_config', {})))

@app.route('/update_index/<int:question_id>', methods=['POST'])
def update_index(question_id):
    try:
        if q_ref := QuestionAnswer.query.get(question_id):
            for q in QuestionAnswer.query.filter_by(question=q_ref.question, category=q_ref.category).all():
                q.sort_index = int(request.form.get('index', 99))
            db.session.commit()
            return jsonify({'success': True})
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': str(e)})
    return jsonify({'success': False, 'message': 'Frage nicht gefunden.'})


def generate_filtered_solutions_pdf(bearbeiter):
    try:
        df_excel = pd.read_excel('Daten.xlsx', sheet_name='Fragestellungen', header=None)
        project_config = session.get('project_config', {})
        
        category_config_keys = {
            'Trafo': 'num_trafos',
            'Einspeisung': 'num_einspeisungen',
            'Abgang': 'num_abgaenge'
        }

        category_configs = {
            'Trafo': {'header_row': 14, 'data_start_row': 15, 'data_end_row': 38, 'solution_start_col': 26},
            'Einspeisung': {'header_row': 40, 'data_start_row': 41, 'data_end_row': 64, 'solution_start_col': 26},
            'Abgang': {'header_row': 77, 'data_start_row': 78, 'data_end_row': 101, 'solution_start_col': 26}
        }

        pdf = FPDF(orientation='P', unit='mm', format='A4')
        create_pdf_cover(pdf, bearbeiter, "Gefilterte Lösungen")
        found_any_solution = False
        
        # NEU: Eine Liste für Diagnose-Nachrichten
        diagnostics = []

        for category, config in category_configs.items():
            config_key = category_config_keys.get(category)
            num_components = project_config.get(config_key, 0)

            if num_components == 0:
                continue
            
            header_series = df_excel.iloc[config['header_row']]
            df_category_data = df_excel.iloc[config['data_start_row']:config['data_end_row']].copy()
            df_category_data.columns = [str(h).strip() if pd.notna(h) else '' for h in header_series]

            for i in range(1, num_components + 1):
                component_name = f"{category} {i}"
                df_to_filter = df_category_data.copy()
                
                answers = QuestionAnswer.query.filter(
                    QuestionAnswer.category == category, QuestionAnswer.category_index == i,
                    QuestionAnswer.answer.isnot(None), QuestionAnswer.answer != '',
                    QuestionAnswer.answer != 'nicht Relevant'
                ).all()

                # Diagnose-Schritt 1: Prüfen, ob Antworten in der DB gefunden wurden
                if not answers:
                    diagnostics.append(f"Für '{component_name}' wurden keine relevanten Antworten in der Datenbank gefunden. Wird übersprungen.")
                    continue
                
                diagnostics.append(f"Für '{component_name}' wurden {len(answers)} Antworten gefunden. Beginne Filterung.")

                for answer_obj in answers:
                    question_text = answer_obj.question.strip()
                    user_answer = answer_obj.answer.strip()
                    
                    if question_text not in df_to_filter.columns:
                        continue
                    
                    # ... (Ihre bestehende Filterlogik bleibt hier unverändert)
                    if question_text == "Spannungsversorgung des Messgerätes?":
                        if not (match := re.search(r'(\d+)\s*v?\s*(ac|dc|ac/dc)?', user_answer.lower())): continue
                        user_voltage, user_type = int(match.group(1)), (match.group(2) or "ac/dc").upper()
                        def check_voltage(cell):
                            if pd.isna(cell): return False
                            return any(s['min_v'] <= user_voltage <= s['max_v'] and user_type in s['type'] for s in parse_voltage_string(str(cell)))
                        condition = (df_to_filter[question_text].isna()) | (df_to_filter[question_text].apply(check_voltage))
                    elif question_text == "Bis zur wie vielten Oberschwingung soll gemessen werden?":
                        if not (nums := re.findall(r'(\d+)', user_answer)): continue
                        user_max_h = max(int(n) for n in nums)
                        def check_harmonic(cell):
                            if pd.isna(cell) or not (cell_nums := re.findall(r'(\d+)', str(cell))): return False
                            return max(int(n) for n in cell_nums) >= user_max_h
                        condition = (df_to_filter[question_text].isna()) | (df_to_filter[question_text].apply(check_harmonic))
                    else:
                        condition = (df_to_filter[question_text].isna()) | (df_to_filter[question_text].astype(str).str.contains(user_answer, na=False, case=False, regex=False))
                    
                    df_to_filter = df_to_filter[condition]
                
                final_solutions = df_to_filter.iloc[:, config['solution_start_col']:]
                final_solutions = final_solutions.dropna(how='all', axis=1).dropna(how='all', axis=0)
                
                # Diagnose-Schritt 2: Prüfen, ob nach dem Filtern noch Lösungen übrig sind
                if not final_solutions.empty:
                    diagnostics.append(f"Erfolgreich! Für '{component_name}' wurden {len(final_solutions)} Lösungen gefunden und zur PDF hinzugefügt.")
                    found_any_solution = True
                    pdf.add_page()
                    pdf.set_font("Arial", 'B', 14)
                    pdf.cell(0, 10, txt=f"Lösungen für {component_name}", ln=True, align='L')
                    
                    # ... (Rest der PDF-Tabellenerstellung bleibt unverändert)
                    pdf.set_font("Arial", 'B', 10)
                    col_widths = [(pdf.w - 20) / len(final_solutions.columns)] * len(final_solutions.columns)
                    for j, col_header in enumerate(final_solutions.columns):
                        pdf.cell(col_widths[j], 10, str(col_header), 1, 0, 'C')
                    pdf.ln()
                    pdf.set_font("Arial", '', 9)
                    for _, row in final_solutions.iterrows():
                        for j, item in enumerate(row):
                            pdf.cell(col_widths[j], 10, str(item) if pd.notna(item) else "", 1, 0, 'L')
                        pdf.ln()
                else:
                    diagnostics.append(f"Keine passenden Lösungen für '{component_name}' nach der Filterung gefunden.")

        # Zeige alle Diagnose-Nachrichten als eine einzige Flash-Nachricht an
        flash("Diagnose-Bericht: \n" + "\n".join(diagnostics), "info")

        if not found_any_solution:
            flash("Insgesamt wurden keine passenden Lösungen für die aktuelle Konfiguration gefunden.", "warning")
            return redirect(url_for('fragen', **session.get('project_config', {})))

        pdf_output = pdf.output(dest='S').encode('latin1')
        return send_file(BytesIO(pdf_output), as_attachment=True, download_name='Gefilterte_Loesungen.pdf', mimetype='application/pdf')
    
    except Exception as e:
        print(f"Ein Fehler ist aufgetreten: {e}") 
        flash(f"Ein Fehler ist beim Erstellen des Lösungs-PDFs aufgetreten: {e}", "error")
        return redirect(url_for('fragen', **session.get('project_config', {})))


@app.route('/download_filtered_pdf', methods=['POST'])
def download_filtered_pdf():
    bearbeiter = request.form.get('bearbeiter', 'N/A')
    # Ruft die neue Hilfsfunktion auf
    return generate_filtered_solutions_pdf(bearbeiter)


@app.route('/import_answers_pdf', methods=['POST'])
def import_answers_pdf():
    bearbeiter = request.form.get('bearbeiter_import', 'N/A')
    if 'answers_pdf' not in request.files:
        flash('Keine Datei hochgeladen.', 'error')
        return redirect(url_for('fragen'))
        
    file = request.files['answers_pdf']

    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        flash('Bitte wählen Sie eine gültige PDF-Datei aus.', 'error')
        return redirect(url_for('fragen'))

    try:
        reader = PdfReader(file.stream)
        fields = reader.get_fields()
        
        if not fields:
            flash('Die hochgeladene PDF enthält keine ausfüllbaren Felder.', 'warning')
            return redirect(url_for('fragen'))
            
        answers_updated = 0
        for field_name, field_data in fields.items():
            if field_name.startswith('question_') and field_data.value:
                question_id = int(field_name.split('_')[1])
                answer_value = field_data.value
                
                # NEUE REGEL: Behandelt "Nein" (unabhängig von Groß-/Kleinschreibung)
                # als "nicht Relevant".
                if answer_value.strip().lower() == 'nein':
                    answer_value = 'nicht Relevant'
                
                question_entry = QuestionAnswer.query.get(question_id)
                if question_entry:
                    question_entry.answer = answer_value
                    answers_updated += 1
        
        db.session.commit()
        flash(f'{answers_updated} Antworten wurden erfolgreich aus der PDF importiert!', 'success')
        
        return generate_filtered_solutions_pdf(bearbeiter)
        
    except Exception as e:
        db.session.rollback()
        flash(f'Ein Fehler ist beim Einlesen der PDF aufgetreten: {e}', 'error')
        print(e)
        return redirect(url_for('fragen', **session.get('project_config', {})))
    
@app.route('/export_questions_pdf', methods=['POST'])
def export_questions_pdf():
    try:
        bearbeiter = request.form.get('bearbeiter', 'N/A')
        project_config = session.get('project_config', {})
        
        # 1. Fragen aus der Datenbank abrufen
        conditions = [QuestionAnswer.category == 'Allgemein']
        if (n := project_config.get('num_trafos', 0)) > 0: conditions.append((QuestionAnswer.category == 'Trafo') & (QuestionAnswer.category_index <= n))
        if (n := project_config.get('num_einspeisungen', 0)) > 0: conditions.append((QuestionAnswer.category == 'Einspeisung') & (QuestionAnswer.category_index <= n))
        if (n := project_config.get('num_abgaenge', 0)) > 0: conditions.append((QuestionAnswer.category == 'Abgang') & (QuestionAnswer.category_index <= n))
        
        fragen_db = QuestionAnswer.query.filter(or_(*conditions)).order_by(QuestionAnswer.category, QuestionAnswer.category_index, QuestionAnswer.sort_index).all()
        
        if not fragen_db:
            flash("Keine Fragen zum Exportieren für die aktuelle Konfiguration vorhanden.", "info")
            return redirect(url_for('fragen', **session.get('project_config', {})))
            
        # 2. PDF-Erstellung vorbereiten
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=A4)
        width, height = A4
        c.acroForm

        styles = getSampleStyleSheet()
        style = styles["Normal"]
        style.alignment = TA_LEFT
        style.leading = 14

        # --- Deckblatt zeichnen (JETZT MIT LOGO) ---
        logo_path = os.path.join(basedir, 'static', 'img', 'logo.png')
        if os.path.exists(logo_path):
            # Logo zentriert oben platzieren (Breite 110 Punkte, Höhe wird automatisch angepasst)
            c.drawImage(logo_path, x=(width/2 - 63*mm), y=(height - 150*mm), width=355, preserveAspectRatio=True, mask='auto')

        c.setFont("Helvetica-Bold", 24)
        c.drawCentredString(width / 2, height - 150*mm, "Fragebogen zur ISO50001") # Text nach unten verschoben
        c.setFont("Helvetica", 12)
        c.drawCentredString(width / 2, height - 180*mm, f"Bearbeiter: {bearbeiter}") # Text nach unten verschoben
        c.drawCentredString(width / 2, height - 200*mm, f"Datum: {date.today().strftime('%d.%m.%Y')}") # Text nach unten verschoben
        c.showPage()

        # 3. Layout-Variablen und Header-Funktion
        x_margin = 18 * mm
        col_widths = [80 * mm, 55 * mm, 45 * mm]
        
        def draw_header(y_start):
            c.setFont("Helvetica-Bold", 10)
            header_height = 8 * mm
            c.drawString(x_margin + 2*mm, y_start - (header_height / 1.5), "Frage")
            c.drawString(x_margin + col_widths[0] + 2*mm, y_start - (header_height / 1.5), "Antwortmöglichkeiten")
            c.drawString(x_margin + col_widths[0] + col_widths[1] + 2*mm, y_start - (header_height / 1.5), "Antwort")
            c.grid([x_margin, x_margin + col_widths[0], x_margin + col_widths[0] + col_widths[1], x_margin + sum(col_widths)], [y_start, y_start - header_height])
            return y_start - header_height

        # 4. Fragen-Tabellen zeichnen
        from itertools import groupby
        category_order = ['Allgemein', 'Trafo', 'Einspeisung', 'Abgang']
        keyfunc = lambda q: (q.category, q.category_index)
        grouped_data = {k: list(v) for k, v in groupby(fragen_db, key=keyfunc)}

        y_cursor = height - 25 * mm

        for category in category_order:
            instanzen = sorted([k[1] for k in grouped_data.keys() if k[0] == category])
            if not instanzen: continue

            for i in instanzen:
                questions = grouped_data.get((category, i), [])
                if not questions: continue

                if y_cursor < 100 * mm:
                    c.showPage()
                    y_cursor = height - 25 * mm

                title = category if category == 'Allgemein' else f'{category} {i}'
                c.setFont("Helvetica-Bold", 14)
                c.drawString(x_margin, y_cursor, title)
                y_cursor -= 10*mm
                
                y_cursor = draw_header(y_cursor)

                for q in questions:
                    display_options = q.options.replace(',', ', ')
                    if 'ja' in display_options.lower() and 'nein' not in display_options.lower():
                        display_options += ", Nein"
                    
                    question_p = Paragraph(q.question, style)
                    options_p = Paragraph(display_options, style)

                    q_height = question_p.wrap(col_widths[0] - 4*mm, height)[1]
                    o_height = options_p.wrap(col_widths[1] - 4*mm, height)[1]
                    
                    row_height = max(10 * mm, q_height + 4*mm, o_height + 4*mm)

                    if y_cursor - row_height < 40 * mm:
                        c.showPage()
                        y_cursor = height - 25 * mm
                        y_cursor = draw_header(y_cursor)

                    question_p.drawOn(c, x_margin + 2*mm, y_cursor - row_height + 2*mm)
                    options_p.drawOn(c, x_margin + col_widths[0] + 2*mm, y_cursor - row_height + 2*mm)
                    
                    c.acroForm.textfield(
                        name=f'question_{q.id}',
                        x=x_margin + col_widths[0] + col_widths[1] + 2*mm,
                        y=y_cursor - row_height + 2*mm,
                        width=col_widths[2] - 4*mm,
                        height=row_height - 4*mm,
                        borderStyle='solid', borderWidth=1, borderColor=colors.black,
                    )

                    c.grid([x_margin, x_margin + col_widths[0], x_margin + col_widths[0] + col_widths[1], x_margin + sum(col_widths)], [y_cursor, y_cursor - row_height])
                    y_cursor -= row_height
                
                y_cursor -= 10 * mm

        # 5. PDF speichern und senden
        c.save()
        pdf_output = buffer.getvalue()
        buffer.close()
        
        return send_file(io.BytesIO(pdf_output), as_attachment=True, download_name='Fragebogen_ausfuellbar.pdf', mimetype='application/pdf')

    except Exception as e:
        flash(f"Ein Fehler ist beim Erstellen des Fragebogen-PDFs aufgetreten: {e}", "error")
        import traceback
        traceback.print_exc()
        return redirect(url_for('fragen', **session.get('project_config', {})))


def generate_category_chart(entries):
    """ Erstellt ein Kuchendiagramm mit transparentem Hintergrund und speichert es als PNG."""
    
    img_dir = os.path.join(basedir, 'static', 'img')
    os.makedirs(img_dir, exist_ok=True)
    filepath = os.path.join(img_dir, 'category_chart.png')

    # Fall 1: Keine Einträge vorhanden
    if not entries:
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.text(0.5, 0.5, 'Keine Daten für die Auswertung vorhanden.', 
                ha='center', va='center', fontsize=14, color='gray')
        ax.axis('off')
        
        # Hinzugefügt: Transparenter Hintergrund
        plt.savefig(filepath, bbox_inches='tight', transparent=True)
        plt.close(fig)
        return

    # Fall 2: Einträge sind vorhanden
    try:
        data = [{'category': entry.category, 'duration': entry.duration.total_seconds() / 3600} for entry in entries]
        df = pd.DataFrame(data)
        category_totals = df.groupby('category')['duration'].sum()

        fig, ax = plt.subplots(figsize=(10, 7))
        
        wedges, texts, autotexts = ax.pie(
            category_totals.values, 
            labels=category_totals.index, 
            autopct='%1.1f%%',
            startangle=90,
            pctdistance=0.85,
            explode=[0.05] * len(category_totals)
        )

        # Hinzugefügt: Textfarbe auf Grau geändert für bessere Kompatibilität
        plt.setp(autotexts, size=10, weight="bold", color="white")
        plt.setp(texts, size=12, color="dimgray") # Ein dunkles Grau

        # Hinzugefügt: Textfarbe für Titel
        ax.set_title('Zeitverteilung nach Kategorien', size=16, color="dimgray")
        ax.axis('equal')

        plt.tight_layout()
        
        # Hinzugefügt: Transparenter Hintergrund
        plt.savefig(filepath, transparent=True)
        plt.close(fig)

    except Exception as e:
        print(f"Fehler beim Erstellen der Grafik: {e}")


# --- App Start ---
if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(host='0.0.0.0', port=5050, debug=True)