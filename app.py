import os
import json
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify, session
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import or_, func, and_, not_
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
from reportlab.pdfbase.pdfdoc import PDFName, PDFString, PDFArray, PDFDictionary
from PIL import Image # Wichtig: Fügen Sie diesen Import hinzu

basedir = os.path.abspath(os.path.dirname(__file__))

# --- App Konfiguration ---
app = Flask(__name__)
app.config['SECRET_KEY'] = 'dein_super_geheimer_schluessel_12345'
basedir = os.path.abspath(os.path.dirname(__file__))
# Standard-Datenbank für die Zeiterfassung
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(basedir, 'zeiterfassung.db')
# Zusätzliche Datenbank für die Fragen
app.config['SQLALCHEMY_BINDS'] = {
    'fragen': 'sqlite:///' + os.path.join(basedir, 'fragen.db')
}
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# --- Datenbankmodelle ---
class TimeEntry(db.Model):
    __tablename__ = 'time_entry'
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
    __bind_key__ = 'fragen'
    __tablename__ = 'question_answer'
    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.String(50), nullable=False)
    category_index = db.Column(db.Integer, nullable=False, default=1)
    question = db.Column(db.String(500), nullable=False)
    options = db.Column(db.String(500), nullable=False)
    answer = db.Column(db.String(100), nullable=True)
    sort_index = db.Column(db.Integer, default=99)
    sasil_abgang_index = db.Column(db.Integer, nullable=True)
    __table_args__ = (db.UniqueConstraint('category', 'category_index', 'question', 'sasil_abgang_index', name='_category_question_uc'),)


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

    entries = TimeEntry.query.order_by(TimeEntry.date.desc(), TimeEntry.start_time.desc()).all()
    generate_category_chart(entries) 
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
        db.session.query(TimeEntry).delete()
        db.session.commit()
        flash('Alle Zeiterfassungseinträge wurden gelöscht.', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Fehler beim Löschen aller Einträge: {e}', 'error')
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
                'num_sasil': int(request.args.get('num_sasil', 0)),
                'sasil_abgaenge_counts': {}
            }
            old_config = session.get('project_config', {})
            
            for i in range(1, new_config['num_sasil'] + 1):
                new_config['sasil_abgaenge_counts'][str(i)] = 1

            categories_to_process = {
                'Trafo': 'num_trafos',
                'Einspeisung': 'num_einspeisungen',
                'Abgang': 'num_abgaenge',
                'SASIL': 'num_sasil'
            }
            for cat, key in categories_to_process.items():
                old_count = old_config.get(key, 0)
                new_count = new_config.get(key, 0)

                if new_count > old_count:
                    master_questions = QuestionAnswer.query.filter_by(category=cat, category_index=1, sasil_abgang_index=None if cat != 'SASIL' else 1).all()
                    for i in range(old_count + 1, new_count + 1):
                        for master_q in master_questions:
                            sasil_abgang_index = 1 if cat == 'SASIL' else None
                            if not QuestionAnswer.query.filter_by(category=cat, category_index=i, sasil_abgang_index=sasil_abgang_index, question=master_q.question).first():
                                db.session.add(QuestionAnswer(
                                    category=cat, category_index=i, question=master_q.question,
                                    options=master_q.options, sort_index=master_q.sort_index,
                                    sasil_abgang_index=sasil_abgang_index
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
                category = request.form.get('category')
                category_index = int(request.form.get('category_index'))
                sasil_abgang_index_str = request.form.get('sasil_abgang_index')
                sasil_abgang_index = int(sasil_abgang_index_str) if sasil_abgang_index_str else None
                
                max_index = db.session.query(func.max(QuestionAnswer.sort_index)).filter_by(category=category).scalar() or 0

                db.session.add(QuestionAnswer(
                    question=request.form.get('new_question'), 
                    options=request.form.get('options'),
                    category=category, 
                    category_index=category_index, 
                    sasil_abgang_index=sasil_abgang_index,
                    sort_index=max_index + 1
                ))

                db.session.commit()
                flash(f'Frage erfolgreich für "{category}" erstellt!', 'success')
            except Exception as e:
                db.session.rollback()
                flash(f'Fehler beim Erstellen der Frage: {e}', 'error')
        
        anchor = f"{request.form.get('category')}-{request.form.get('category_index')}"
        if request.form.get('sasil_abgang_index'):
            anchor += f"-{request.form.get('sasil_abgang_index')}"
        return redirect(url_for('fragen', _anchor=anchor))

    project_config = session.get('project_config', {})
    if 'sasil_abgaenge_counts' not in project_config:
        project_config['sasil_abgaenge_counts'] = {}

    questions_db = QuestionAnswer.query.order_by(
        QuestionAnswer.category, 
        QuestionAnswer.category_index, 
        QuestionAnswer.sasil_abgang_index.nullslast(),
        QuestionAnswer.sort_index
    ).all()
    
    grouped_questions = {}
    for q in questions_db:
        key = (q.category, q.category_index, q.sasil_abgang_index)
        if key not in grouped_questions: grouped_questions[key] = []
        grouped_questions[key].append(q)
        
    return render_template('fragen.html', project_config=project_config, grouped_questions=grouped_questions)


@app.route('/synchronize_questions', methods=['POST'])
def synchronize_questions():
    try:
        source_category = request.form.get('source_category')
        source_category_index = int(request.form.get('source_category_index'))
        source_sasil_abgang_index_str = request.form.get('source_sasil_abgang_index')
        source_sasil_abgang_index = int(source_sasil_abgang_index_str) if source_sasil_abgang_index_str else None

        source_questions = QuestionAnswer.query.filter_by(
            category=source_category,
            category_index=source_category_index,
            sasil_abgang_index=source_sasil_abgang_index
        ).all()

        if not source_questions:
            flash('Keine Fragen in der Quelle zum Synchronisieren gefunden.', 'warning')
            return redirect(url_for('fragen'))

        project_config = session.get('project_config', {})
        target_categories_map = {
            'Trafo': 'num_trafos',
            'Einspeisung': 'num_einspeisungen',
            'Abgang': 'num_abgaenge',
            'SASIL': 'num_sasil'
        }
        questions_added_count = 0

        for cat, num_key in target_categories_map.items():
            num_instances = project_config.get(num_key, 0)
            
            for i in range(1, num_instances + 1):
                if cat == 'SASIL':
                    sasil_counts = project_config.get('sasil_abgaenge_counts', {})
                    num_abgaenge_sasil = sasil_counts.get(str(i), 1)
                    for j in range(1, num_abgaenge_sasil + 1):
                        if cat == source_category and i == source_category_index and j == source_sasil_abgang_index:
                            continue
                        
                        for q in source_questions:
                            exists = QuestionAnswer.query.filter_by(
                                category=cat, category_index=i, sasil_abgang_index=j, question=q.question
                            ).first()
                            if not exists:
                                db.session.add(QuestionAnswer(
                                    category=cat, category_index=i, sasil_abgang_index=j,
                                    question=q.question, options=q.options, sort_index=q.sort_index
                                ))
                                questions_added_count += 1
                else:
                    if cat == source_category and i == source_category_index and source_sasil_abgang_index is None:
                        continue
                        
                    for q in source_questions:
                        exists = QuestionAnswer.query.filter_by(
                            category=cat, category_index=i, sasil_abgang_index=None, question=q.question
                        ).first()
                        if not exists:
                            db.session.add(QuestionAnswer(
                                category=cat, category_index=i,
                                question=q.question, options=q.options, sort_index=q.sort_index
                            ))
                            questions_added_count += 1

        db.session.commit()
        if questions_added_count > 0:
            flash(f'{questions_added_count} neue Fragen wurden zu den anderen Bereichen hinzugefügt.', 'success')
        else:
            flash('Alle Fragen waren bereits in den anderen Bereichen vorhanden. Nichts wurde hinzugefügt.', 'info')

    except Exception as e:
        db.session.rollback()
        flash(f'Fehler bei der Synchronisierung: {e}', 'error')
    
    anchor = f"{request.form.get('source_category')}-{request.form.get('source_category_index')}"
    if request.form.get('source_sasil_abgang_index'):
        anchor += f"-{request.form.get('source_sasil_abgang_index')}"
    return redirect(url_for('fragen', _anchor=anchor))


@app.route('/configure_sasil_abgaenge', methods=['POST'])
def configure_sasil_abgaenge():
    try:
        sasil_index_str = request.form.get('sasil_index')
        sasil_index = int(sasil_index_str)
        new_abgaenge_count = int(request.form.get('num_abgaenge'))

        project_config = session.get('project_config', {})
        sasil_counts = project_config.get('sasil_abgaenge_counts', {})
        old_abgaenge_count = sasil_counts.get(sasil_index_str, 0)

        if new_abgaenge_count > old_abgaenge_count:
            master_questions = QuestionAnswer.query.filter_by(category='SASIL', category_index=sasil_index, sasil_abgang_index=1).all()
            for i in range(old_abgaenge_count + 1, new_abgaenge_count + 1):
                for master_q in master_questions:
                    if not QuestionAnswer.query.filter_by(category='SASIL', category_index=sasil_index, sasil_abgang_index=i, question=master_q.question).first():
                        db.session.add(QuestionAnswer(
                            category='SASIL', category_index=sasil_index, sasil_abgang_index=i,
                            question=master_q.question, options=master_q.options, sort_index=master_q.sort_index
                        ))
        elif new_abgaenge_count < old_abgaenge_count:
            QuestionAnswer.query.filter(
                and_(
                    QuestionAnswer.category == 'SASIL',
                    QuestionAnswer.category_index == sasil_index,
                    QuestionAnswer.sasil_abgang_index > new_abgaenge_count
                )
            ).delete(synchronize_session=False)

        project_config['sasil_abgaenge_counts'][sasil_index_str] = new_abgaenge_count
        session['project_config'] = project_config
        
        db.session.commit()
        flash(f'Anzahl der Abgänge für SASIL {sasil_index} wurde auf {new_abgaenge_count} aktualisiert.', 'success')

    except Exception as e:
        db.session.rollback()
        flash(f'Fehler bei der Konfiguration der SASIL-Abgänge: {e}', 'error')

    return redirect(url_for('fragen', _anchor=f"SASIL-{request.form.get('sasil_index')}-1", **session.get('project_config', {})))


@app.route('/copy_sasil_answers', methods=['POST'])
def copy_sasil_answers():
    try:
        sasil_index = int(request.form.get('sasil_index'))
        project_config = session.get('project_config', {})
        sasil_counts = project_config.get('sasil_abgaenge_counts', {})
        num_abgaenge_sasil = sasil_counts.get(str(sasil_index), 0)

        source_answers = QuestionAnswer.query.filter_by(category='SASIL', category_index=sasil_index, sasil_abgang_index=1).all()
        
        for i in range(2, num_abgaenge_sasil + 1):
            for source_answer in source_answers:
                target_question = QuestionAnswer.query.filter_by(
                    category='SASIL', 
                    category_index=sasil_index, 
                    sasil_abgang_index=i,
                    question=source_answer.question
                ).first()
                if target_question:
                    target_question.answer = source_answer.answer

        db.session.commit()
        flash(f'Antworten für SASIL {sasil_index} wurden erfolgreich auf alle Abgänge übertragen.', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Fehler beim Kopieren der Antworten: {e}', 'error')
    
    return redirect(url_for('fragen', _anchor=f"SASIL-{request.form.get('sasil_index')}-1", **session.get('project_config', {})))


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
    anchor = f"{q_ref.category}-{q_ref.category_index}"
    if q_ref.category == 'SASIL':
        anchor += f"-{q_ref.sasil_abgang_index}"
    return redirect(url_for('fragen', _anchor=anchor, **session.get('project_config', {})))

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
    anchor = f"{q_ref.category}-{q_ref.category_index}"
    if q_ref.category == 'SASIL':
        anchor += f"-{q_ref.sasil_abgang_index}"
    return redirect(url_for('fragen', _anchor=anchor, **session.get('project_config', {})))


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
        project_config = session.get('project_config', {})
        
        category_config_keys = {
            'Trafo': 'num_trafos',
            'Einspeisung': 'num_einspeisungen',
            'Abgang': 'num_abgaenge',
            'SASIL': 'num_sasil',
        }

        category_configs = {
            'Trafo': {'sheet_name': 'Filter_Trafo', 'header_row': 14, 'data_start_row': 15, 'solution_start_col': 26},
            'Einspeisung': {'sheet_name': 'Filter_Einspeisung', 'header_row': 14, 'data_start_row': 15, 'solution_start_col': 26},
            'Abgang': {'sheet_name': 'Filter_Abgang', 'header_row': 14, 'data_start_row': 15, 'solution_start_col': 26},
            'SASIL': {'sheet_name': 'Filter_Sasil', 'header_row': 14, 'data_start_row': 15, 'solution_start_col': 26}
        }

        pdf = FPDF(orientation='P', unit='mm', format='A4')
        create_pdf_cover(pdf, bearbeiter, "Gefilterte Lösungen")
        found_any_solution = False
        diagnostics = []
        pdf.add_page()

        for category, config in category_configs.items():
            num_components = project_config.get(category_config_keys.get(category), 0)
            if num_components == 0: continue
            
            try:
                df_sheet = pd.read_excel('Daten.xlsx', sheet_name=config['sheet_name'], header=None)
            except ValueError:
                diagnostics.append(f"FEHLER: Das Tabellenblatt '{config['sheet_name']}' wurde in 'Daten.xlsx' nicht gefunden.")
                continue

            header_series = df_sheet.iloc[config['header_row']]
            df_category_data = df_sheet.iloc[config['data_start_row']:].copy()
            df_category_data.columns = [str(h).strip() if pd.notna(h) else '' for h in header_series]

            for i in range(1, num_components + 1):
                sasil_counts = project_config.get('sasil_abgaenge_counts', {})
                num_abgaenge_loop = sasil_counts.get(str(i), 1) if category == 'SASIL' else 1
                
                for j in range(1, num_abgaenge_loop + 1):
                    component_name = f"{category} {i}"
                    if category == 'SASIL': component_name += f" Abgang {j}"
                    
                    df_to_filter = df_category_data.copy()
                    
                    query_filter = {'category': category, 'category_index': i}
                    if category == 'SASIL': query_filter['sasil_abgang_index'] = j
                    
                    answers = QuestionAnswer.query.filter_by(**query_filter).filter(
                        QuestionAnswer.answer.isnot(None),
                        QuestionAnswer.answer != '',
                        QuestionAnswer.answer != 'nicht Relevant'
                    ).all()

                    if not answers:
                        diagnostics.append(f"Für '{component_name}' wurden keine relevanten Antworten gefunden.")
                        continue
                    
                    diagnostics.append(f"Für '{component_name}' wurden {len(answers)} Antworten gefunden. Beginne Filterung.")
                    
                    for answer_obj in answers:
                        question_text, user_answer = answer_obj.question.strip(), answer_obj.answer.strip()
                        if question_text not in df_to_filter.columns: continue
                        
                        if question_text == "Spannungsversorgung des Messgerätes?":
                            match = re.search(r'(\d+)\s*v?\s*(ac|dc|ac/dc)?', user_answer.lower())
                            if not match: continue
                            user_voltage, user_type = int(match.group(1)), (match.group(2) or "ac/dc").upper()
                            def check_voltage(cell):
                                return pd.notna(cell) and any(s['min_v'] <= user_voltage <= s['max_v'] and user_type in s['type'] for s in parse_voltage_string(str(cell)))
                            condition = df_to_filter[question_text].apply(check_voltage) | df_to_filter[question_text].isna()
                        elif question_text == "Bis zur wie vielten Oberschwingung soll gemessen werden?":
                            nums = re.findall(r'(\d+)', user_answer)
                            if not nums: continue
                            user_max_h = max(int(n) for n in nums)
                            def check_harmonic(cell):
                                cell_nums = re.findall(r'(\d+)', str(cell))
                                return pd.notna(cell) and cell_nums and max(int(n) for n in cell_nums) >= user_max_h
                            condition = df_to_filter[question_text].apply(check_harmonic) | df_to_filter[question_text].isna()
                        else:
                            condition = df_to_filter[question_text].astype(str).str.contains(user_answer, na=True, case=False, regex=False)
                        
                        df_to_filter = df_to_filter[condition]
                    
                    final_solutions = df_to_filter.iloc[:, config['solution_start_col']:].dropna(how='all', axis=1).dropna(how='all', axis=0)
                    
                    if not final_solutions.empty:
                        diagnostics.append(f"Erfolgreich! Für '{component_name}' wurden {len(final_solutions)} Lösungen gefunden.")
                        found_any_solution = True

                        if pdf.get_y() + (len(final_solutions) + 2) * 10 > (pdf.h - pdf.b_margin): pdf.add_page()
                        
                        pdf.set_font("Arial", 'B', 14); pdf.cell(0, 10, txt=f"Lösungen für {component_name}", ln=True, align='L')
                        pdf.set_font("Arial", 'B', 10)
                        
                        col_widths = [(pdf.w - 20) / len(final_solutions.columns)] * len(final_solutions.columns)
                        for k, col_header in enumerate(final_solutions.columns): pdf.cell(col_widths[k], 10, str(col_header), 1, 0, 'C')
                        pdf.ln()

                        pdf.set_font("Arial", '', 9)
                        for _, row in final_solutions.iterrows():
                            for k, item in enumerate(row): pdf.cell(col_widths[k], 10, str(item) if pd.notna(item) else "", 1, 0, 'L')
                            pdf.ln()
                    else:
                        diagnostics.append(f"Keine passenden Lösungen für '{component_name}' gefunden.")

        flash("Diagnose-Bericht: \n" + "\n".join(diagnostics), "info")

        if not found_any_solution:
            flash("Insgesamt wurden keine passenden Lösungen gefunden.", "warning")
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
    return generate_filtered_solutions_pdf(bearbeiter)


@app.route('/import_answers_pdf', methods=['POST'])
def import_answers_pdf():
    bearbeiter = request.form.get('bearbeiter_import', 'N/A')
    if 'answers_pdf' not in request.files or not request.files['answers_pdf'].filename:
        flash('Keine Datei hochgeladen.', 'error')
        return redirect(url_for('fragen'))
        
    file = request.files['answers_pdf']
    if not file.filename.lower().endswith('.pdf'):
        flash('Bitte wählen Sie eine gültige PDF-Datei aus.', 'error')
        return redirect(url_for('fragen'))

    try:
        project_config = session.get('project_config', {})
        if not project_config:
            flash('Keine Projektkonfiguration in der Session gefunden.', 'error')
            return redirect(url_for('fragen'))

        # Bestehende Antworten für das Projekt zurücksetzen
        conditions_to_reset = [QuestionAnswer.category == 'Allgemein']
        if project_config.get('num_trafos', 0) > 0: conditions_to_reset.append(QuestionAnswer.category == 'Trafo')
        if project_config.get('num_einspeisungen', 0) > 0: conditions_to_reset.append(QuestionAnswer.category == 'Einspeisung')
        if project_config.get('num_abgaenge', 0) > 0: conditions_to_reset.append(QuestionAnswer.category == 'Abgang')
        if project_config.get('num_sasil', 0) > 0: conditions_to_reset.append(QuestionAnswer.category == 'SASIL')
        QuestionAnswer.query.filter(or_(*conditions_to_reset)).update({QuestionAnswer.answer: None}, synchronize_session=False)
        
        reader = PdfReader(file.stream)
        if not (fields := reader.get_fields()):
            flash('Die PDF enthält keine ausfüllbaren Felder.', 'warning')
            return redirect(url_for('fragen'))
            
        updated_count = 0
        sasil_to_sync = set() # Speichert die Indizes der zu synchronisierenden SASILs

        # 1. Durchlaufe alle Felder aus dem PDF
        for name, data in fields.items():
            value = data.get("/V")
            
            # Reguläre Antworten importieren
            if name.startswith('question_') and value and str(value).strip():
                # Der Name kann jetzt z.B. 'question_123_abgang_1' sein
                parts = name.split('_')
                q_id = int(parts[1])
                
                # Konvertiere den PDF-Wert in einen speicherbaren String
                # /Off ist der Wert für eine nicht angekreuzte Box, /Yes (oder ein anderer Name) für eine angekreuzte
                answer_val = str(value)
                if answer_val.startswith('/'):
                    answer_val = answer_val[1:] # Entferne das '/'
                
                answer = 'nicht Relevant' if answer_val.lower() == 'nein' else answer_val
                
                QuestionAnswer.query.filter_by(id=q_id).update({QuestionAnswer.answer: answer})
                updated_count += 1
            
            # Prüfe, ob es sich um eine Sync-Checkbox handelt und ob sie angekreuzt ist
            if name.startswith('sync_sasil_') and value and value != '/Off':
                sasil_index = int(name.split('_')[-1])
                sasil_to_sync.add(sasil_index)

        # 2. Führe die Synchronisierung für die markierten SASILs durch
        if sasil_to_sync:
            flash(f"Synchronisierung für SASIL-Felder {list(sasil_to_sync)} wird durchgeführt.", "info")
            for sasil_index in sasil_to_sync:
                # Hole alle Antworten von Abgang 1 für dieses SASIL-Feld
                source_answers = QuestionAnswer.query.filter_by(
                    category='SASIL',
                    category_index=sasil_index,
                    sasil_abgang_index=1
                ).all()

                sasil_counts = project_config.get('sasil_abgaenge_counts', {})
                num_abgaenge = sasil_counts.get(str(sasil_index), 1)

                # Kopiere die Antworten auf alle anderen Abgänge (2, 3, ...)
                for i in range(2, num_abgaenge + 1):
                    for source_q in source_answers:
                        # Finde die Zielfrage und update sie
                        QuestionAnswer.query.filter_by(
                            category='SASIL',
                            category_index=sasil_index,
                            sasil_abgang_index=i,
                            question=source_q.question # Finde die Frage mit demselben Text
                        ).update({QuestionAnswer.answer: source_q.answer})
        
        db.session.commit()
        flash(f'{updated_count} Antworten wurden aus der PDF importiert!', 'success')
        # Nach dem Import direkt die Lösungs-PDF generieren und herunterladen
        return generate_filtered_solutions_pdf(bearbeiter)
        
    except Exception as e:
        db.session.rollback()
        flash(f'Fehler beim Einlesen der PDF: {e}', 'error')
        import traceback
        traceback.print_exc()
        return redirect(url_for('fragen', **session.get('project_config', {})))
    
@app.route('/export_questions_pdf', methods=['POST'])
def export_questions_pdf():
    # Wir benötigen diese speziellen Imports nicht mehr
    # from reportlab.pdfbase.pdfdoc import PDFName, PDFString, PDFArray, PDFDictionary

    try:
        bearbeiter = request.form.get('bearbeiter', 'N/A')
        kunde = request.form.get('kunde', 'N/A')
        project_config = session.get('project_config', {})
        
        # ... (Datenbankabfrage bleibt unverändert) ...
        conditions = [QuestionAnswer.category == 'Allgemein']
        if (n := project_config.get('num_trafos', 0)) > 0: conditions.append(QuestionAnswer.category == 'Trafo')
        if (n := project_config.get('num_einspeisungen', 0)) > 0: conditions.append(QuestionAnswer.category == 'Einspeisung')
        if (n := project_config.get('num_abgaenge', 0)) > 0: conditions.append(QuestionAnswer.category == 'Abgang')
        if (n := project_config.get('num_sasil', 0)) > 0: conditions.append(QuestionAnswer.category == 'SASIL')
        
        fragen_db = QuestionAnswer.query.filter(or_(*conditions)).order_by(
            QuestionAnswer.category, QuestionAnswer.category_index, 
            QuestionAnswer.sasil_abgang_index.nullslast(), QuestionAnswer.sort_index
        ).all()
        
        if not fragen_db:
            flash("Keine Fragen zum Exportieren vorhanden.", "info")
            return redirect(url_for('fragen', **project_config))
            
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=A4)
        width, height = A4; c.acroForm; styles = getSampleStyleSheet()
        style = styles["Normal"]; style.alignment = TA_LEFT; style.leading = 14

        # ... (Deckblatt- und Header-Logik bleibt unverändert) ...
        logo_path = os.path.join(basedir, 'static', 'img', 'logo.png')
        if os.path.exists(logo_path):
            c.drawImage(logo_path, x=(width/2 - 63*mm), y=(height - 150*mm), width=355, preserveAspectRatio=True, mask='auto')
        c.setFont("Helvetica-Bold", 24); c.drawCentredString(width/2, height - 150*mm, "Fragebogen zur ISO50001")
        c.setFont("Helvetica", 12)
        c.drawCentredString(width/2, height - 180*mm, f"Kunde: {kunde}")
        c.drawCentredString(width/2, height - 200*mm, f"Bearbeiter: {bearbeiter}")
        c.drawCentredString(width/2, height - 220*mm, f"Datum: {date.today().strftime('%d.%m.%Y')}")
        c.showPage()

        x_margin, col_widths = 18 * mm, [80 * mm, 55 * mm, 45 * mm]
        def draw_page_header(canvas, pg_width, pg_height):
            canvas.setFont("Helvetica-Bold", 18)
            canvas.drawString(pg_width/2 - 30*mm, pg_height - 15*mm, 'ISO50001 Fragebogen')
            if os.path.exists(logo_path):
                canvas.drawImage(logo_path, x=160*mm, y=207*mm, width=105, preserveAspectRatio=True, mask='auto')
        def draw_table_header(y_pos):
            c.setFont("Helvetica-Bold", 10); header_h = 8*mm
            c.drawString(x_margin + 2*mm, y_pos - (header_h/1.5), "Frage")
            c.drawString(x_margin + col_widths[0] + 2*mm, y_pos - (header_h/1.5), "Antwortmöglichkeiten")
            c.drawString(x_margin + sum(col_widths[:2]) + 2*mm, y_pos - (header_h/1.5), "Antwort")
            c.grid([x_margin, x_margin + col_widths[0], x_margin + sum(col_widths[:2]), x_margin + sum(col_widths)], [y_pos, y_pos - header_h])
            return y_pos - header_h

        from itertools import groupby
        keyfunc = lambda q: (q.category, q.category_index, q.sasil_abgang_index)
        grouped_data = {k: list(v) for k, v in groupby(fragen_db, key=keyfunc)}
        
        # KEIN JAVASCRIPT MEHR NÖTIG

        draw_page_header(c, width, height)
        y_cursor = height - 25 * mm

        for category, cat_index, abgang_index in sorted(grouped_data.keys(), key=lambda k: (['Allgemein', 'Trafo', 'Einspeisung', 'Abgang', 'SASIL'].index(k[0]), k[1], k[2] or 0)):
            questions = grouped_data.get((category, cat_index, abgang_index), [])
            if not questions: continue
            if y_cursor < 100 * mm:
                c.showPage(); draw_page_header(c, width, height); y_cursor = height - 25*mm
            title = category if category == 'Allgemein' else f'{category} {cat_index}'
            
            if category == 'SASIL' and abgang_index == 1:
                c.setFont("Helvetica-Bold", 14); c.drawString(x_margin, y_cursor, title); y_cursor -= 8*mm
                
                sasil_counts = project_config.get('sasil_abgaenge_counts', {})
                num_abgaenge = sasil_counts.get(str(cat_index), 1)
                
                if num_abgaenge > 1:
                    # Einfache Checkbox ohne Aktion. Der Name ist wichtig für den Import.
                    c.acroForm.checkbox(
                        name=f'sync_sasil_{cat_index}',
                        x=x_margin,
                        y=y_cursor - 4*mm,
                        tooltip="Wenn angekreuzt, werden die Antworten von Abgang 1 für alle anderen Abgänge dieses Feldes übernommen."
                    )
                    
                    c.setFont("Helvetica", 10)
                    c.drawString(x_margin + 10*mm, y_cursor - 1.75*mm, "Alle Messinstrumente sind gleich zu wählen (Wenn diese Aktion ausgewählt ist bitte nur Abgang 1 ausfüllen)")
                    y_cursor -= 10*mm

            if category == 'SASIL':
                title = f"Abgang {abgang_index}"
            
            # ... (Rest der Funktion zum Zeichnen der Tabellen bleibt unverändert) ...
            c.setFont("Helvetica-Bold", 12); c.drawString(x_margin, y_cursor, title); y_cursor -= 10*mm
            y_cursor = draw_table_header(y_cursor)
            for q in questions:
                field_name = f'question_{q.id}'
                if category == 'SASIL':
                    field_name += f'_abgang_{abgang_index}'
                display_options = "1. - 63." if "Oberschwingung" in q.question else "SpannungV AC oder DC" if "Spannungsversorgung" in q.question else q.options.replace(',', ', ') + (", Nein" if 'ja' in q.options.lower() and 'nein' not in q.options.lower() else "")
                q_p, o_p = Paragraph(q.question, style), Paragraph(display_options, style)
                q_h, o_h = q_p.wrap(col_widths[0] - 4*mm, height)[1], o_p.wrap(col_widths[1] - 4*mm, height)[1]
                row_h = max(10 * mm, q_h + 4*mm, o_h + 4*mm)
                if y_cursor - row_h < 40 * mm:
                    c.showPage(); draw_page_header(c, width, height); y_cursor = height - 25*mm
                    y_cursor = draw_table_header(y_cursor)
                q_p.drawOn(c, x_margin + 2*mm, y_cursor - row_h + 2*mm)
                o_p.drawOn(c, x_margin + col_widths[0] + 2*mm, y_cursor - row_h + 2*mm)
                c.acroForm.textfield(name=field_name, x=x_margin + sum(col_widths[:2]) + 2*mm, y=y_cursor - row_h + 2*mm, width=col_widths[2] - 4*mm, height=row_h - 4*mm, borderStyle='solid', borderWidth=1, borderColor=colors.black)
                c.grid([x_margin, x_margin + col_widths[0], x_margin + sum(col_widths[:2]), x_margin + sum(col_widths)], [y_cursor, y_cursor - row_h])
                y_cursor -= row_h
            y_cursor -= 10*mm

        c.save()
        buffer.seek(0)
        return send_file(buffer, as_attachment=True, download_name='Fragebogen_ausfuellbar.pdf', mimetype='application/pdf')

    except Exception as e:
        flash(f"Fehler beim Erstellen des PDFs: {e}", "error")
        import traceback
        traceback.print_exc()
        return redirect(url_for('fragen', **session.get('project_config', {})))


def generate_category_chart(entries):
    img_dir = os.path.join(basedir, 'static', 'img')
    os.makedirs(img_dir, exist_ok=True)
    filepath = os.path.join(img_dir, 'category_chart.png')

    if not entries:
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.text(0.5, 0.5, 'Keine Daten für die Auswertung vorhanden.', ha='center', va='center', fontsize=14, color='gray')
        ax.axis('off')
        plt.savefig(filepath, bbox_inches='tight', transparent=True); plt.close(fig)
        return

    try:
        df = pd.DataFrame([{'category': e.category, 'duration': e.duration.total_seconds() / 3600} for e in entries])
        category_totals = df.groupby('category')['duration'].sum()

        fig, ax = plt.subplots(figsize=(10, 7))
        wedges, texts, autotexts = ax.pie(category_totals.values, labels=category_totals.index, autopct='%1.1f%%',
                                          startangle=90, pctdistance=0.85, explode=[0.05] * len(category_totals))
        plt.setp(autotexts, size=10, weight="bold", color="white"); plt.setp(texts, size=12, color="dimgray")
        ax.set_title('Zeitverteilung nach Kategorien', size=16, color="dimgray"); ax.axis('equal')
        plt.tight_layout()
        plt.savefig(filepath, transparent=True); plt.close(fig)
    except Exception as e:
        print(f"Fehler beim Erstellen der Grafik: {e}")

# --- App Start ---
if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(host='0.0.0.0', port=5050, debug=True)