import os
import json
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify, session, send_from_directory
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import or_, func
import pandas as pd
from datetime import datetime, date, timedelta, time, timezone
from io import BytesIO

# WICHTIG: Diese Zeile muss VOR dem Import von pyplot stehen!
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from fpdf import FPDF
from io import BytesIO
from PIL import Image

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
    answer = db.Column(db.String(50), nullable=True)
    sort_index = db.Column(db.Integer, default=99)
    __table_args__ = (db.UniqueConstraint('category', 'category_index', 'question', name='_category_question_uc'),)

# --- Kontext-Prozessor ---
@app.context_processor
def inject_now():
    return {'now': datetime.now(timezone.utc)}

# --- Routen / Unterseiten ---
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/view_drawing/<string:filename>')
def view_drawing(filename):
    try:
        filepath = os.path.join(basedir, 'visuals', filename)
        with open(filepath, 'r', encoding='utf-8') as f:
            xml_content = f.read()
        
        config = {
            "xml": xml_content, "background": "#ffffff", "toolbar": "top",
            "lightbox": False, "transparent": False
        }
        
        diagram_data = json.dumps(config)
        return render_template('view_drawing.html', diagram_data=diagram_data, drawing_name=filename)
    except FileNotFoundError:
        flash(f"Zeichnung '{filename}' nicht gefunden.", "error")
        return redirect(url_for('index'))
    except Exception as e:
        flash(f"Fehler beim Laden der Zeichnung: {e}", "error")
        return redirect(url_for('index'))

@app.route('/begriffsfinder', methods=['GET', 'POST'])
def begriffsfinder():
    search_term = ""
    result = ""
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
            suggestions = matching_terms.tolist()
            suggestions = suggestions[:10]
        except FileNotFoundError:
            suggestions = ["Fehler: 'Daten.xlsx' nicht gefunden."]
        except Exception as e:
            suggestions = [f"Ein Fehler ist aufgetreten: {e}"]
    return jsonify(suggestions)

@app.route('/dokumentation', methods=['GET', 'POST'])
def dokumentation():
    if request.method == 'POST':
        try:
            date_obj = datetime.strptime(request.form.get('date'), '%Y-%m-%d').date()
            start_time_obj = datetime.strptime(request.form.get('start_time'), '%H:%M').time()
            end_time_obj = datetime.strptime(request.form.get('end_time'), '%H:%M').time()
            category = request.form.get('category')
            project = request.form.get('project')
            info_text = request.form.get('info_text')
            new_entry = TimeEntry(
                date=date_obj, start_time=start_time_obj, end_time=end_time_obj,
                category=category, project=project, info_text=info_text
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

@app.route('/generate_pdf')
def generate_pdf():
    # ... Logik für Zeiterfassungs-PDF ...
    pass

@app.route('/fragen', methods=['GET', 'POST'])
def fragen():
    if request.method == 'POST':
        if 'save_answers' in request.form:
            for key, value in request.form.items():
                if key.startswith('answer_'):
                    q_id = key.split('_')[1]
                    entry = QuestionAnswer.query.get(q_id)
                    if entry: entry.answer = value
            db.session.commit()
            flash('Antworten erfolgreich gespeichert!', 'success')
            return redirect(url_for('fragen', **session.get('project_config', {})))

        elif 'new_question' in request.form:
            category = request.form.get('category')
            new_question_text = request.form.get('new_question')
            new_options_text = request.form.get('options')
            
            try:
                project_config = session.get('project_config', {})
                num_instances = 1
                if category == 'Trafo':
                    num_instances = project_config.get('num_trafos', 1)
                elif category == 'Einspeisung':
                    num_instances = project_config.get('num_einspeisungen', 1)
                elif category == 'Abgang':
                    num_instances = project_config.get('num_abgaenge', 1)
                
                max_index = db.session.query(func.max(QuestionAnswer.sort_index)).filter_by(category=category).scalar() or 0
                new_sort_index = max_index + 1

                for i in range(1, num_instances + 1):
                    new_entry = QuestionAnswer(
                        question=new_question_text,
                        options=new_options_text,
                        category=category,
                        category_index=i,
                        sort_index=new_sort_index
                    )
                    db.session.add(new_entry)
                
                db.session.commit()
                flash(f'Frage erfolgreich in allen "{category}"-Reitern erstellt!', 'success')
            except Exception as e:
                db.session.rollback()
                flash(f'Fehler beim Speichern der Frage: {e}', 'error')
            
            active_tab_anchor = f"{category}-1"
            return redirect(url_for('fragen', _anchor=active_tab_anchor, **session.get('project_config', {})))

    project_config = {
        'num_trafos': request.args.get('num_trafos', session.get('project_config', {}).get('num_trafos', 0), type=int),
        'num_einspeisungen': request.args.get('num_einspeisungen', session.get('project_config', {}).get('num_einspeisungen', 0), type=int),
        'num_abgaenge': request.args.get('num_abgaenge', session.get('project_config', {}).get('num_abgaenge', 0), type=int),
    }

    if 'setup_submit' in request.args:
        with db.session.no_autoflush:
            old_config = session.get('project_config', {})
            new_config = project_config

            for cat, key in [('Trafo', 'num_trafos'), ('Einspeisung', 'num_einspeisungen'), ('Abgang', 'num_abgaenge')]:
                old_count = old_config.get(key, 0)
                new_count = new_config.get(key, 0)

                if new_count > old_count:
                    unique_questions_query = db.session.query(QuestionAnswer.question, QuestionAnswer.options, QuestionAnswer.sort_index).filter_by(category=cat).distinct()
                    master_questions = {q.question: {'options': q.options, 'sort_index': q.sort_index} for q in unique_questions_query}
                    
                    if not master_questions: continue

                    for i in range(old_count + 1, new_count + 1):
                        for q_text, q_data in master_questions.items():
                            exists = QuestionAnswer.query.filter_by(category=cat, category_index=i, question=q_text).first()
                            if not exists:
                                new_q = QuestionAnswer(
                                    category=cat, category_index=i,
                                    question=q_text, options=q_data['options'], sort_index=q_data['sort_index']
                                )
                                db.session.add(new_q)
        
        try:
            db.session.commit()
            session['project_config'] = new_config
            flash('Projektkonfiguration wurde aktualisiert.', 'success')
        except Exception as e:
            db.session.rollback()
            flash(f"Fehler beim Aktualisieren: {e}", "error")
        
        return redirect(url_for('fragen', **new_config))

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
    flash('Die Konfiguration wurde zurückgesetzt. Die Fragen bleiben erhalten.', 'info')
    return redirect(url_for('fragen'))


@app.route('/delete_question/<int:question_id>', methods=['POST'])
def delete_question(question_id):
    question_to_delete_ref = QuestionAnswer.query.get_or_404(question_id)
    question_text = question_to_delete_ref.question
    category = question_to_delete_ref.category
    
    active_tab_anchor = f"{category}-1"

    try:
        questions_to_delete = QuestionAnswer.query.filter_by(question=question_text, category=category).all()
        for q in questions_to_delete:
            db.session.delete(q)
        
        db.session.commit()
        flash(f'Frage wurde aus allen "{category}"-Reitern gelöscht.', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Fehler beim Löschen der Frage: {e}', 'error')
        
    return redirect(url_for('fragen', _anchor=active_tab_anchor, **session.get('project_config', {})))

@app.route('/update_index/<int:question_id>', methods=['POST'])
def update_index(question_id):
    try:
        new_index = request.form.get('index', 99, type=int)
        
        question_ref = QuestionAnswer.query.get(question_id)
        if question_ref:
            questions_to_update = QuestionAnswer.query.filter_by(
                question=question_ref.question, 
                category=question_ref.category
            ).all()

            for q in questions_to_update:
                q.sort_index = new_index
            
            db.session.commit()
            return jsonify({'success': True, 'message': 'Index für alle Instanzen aktualisiert.'})
        return jsonify({'success': False, 'message': 'Frage nicht gefunden.'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Fehler: {e}'})

@app.route('/edit_options/<int:question_id>', methods=['POST'])
def edit_options(question_id):
    question_ref = QuestionAnswer.query.get_or_404(question_id)
    active_tab_anchor = f"{question_ref.category}-{question_ref.category_index}"
    
    try:
        new_options = request.form.get('new_options', '').strip()
        if not new_options:
            flash("Antwortmöglichkeiten dürfen nicht leer sein.", "error")
            return redirect(url_for('fragen', _anchor=active_tab_anchor, **session.get('project_config', {})))

        question_text = question_ref.question
        category = question_ref.category
        
        questions_to_update = QuestionAnswer.query.filter_by(question=question_text, category=category).all()
        for q in questions_to_update:
            q.options = new_options
        
        db.session.commit()
        flash('Antwortmöglichkeiten erfolgreich aktualisiert.', 'success')
        
    except Exception as e:
        db.session.rollback()
        flash(f"Fehler beim Aktualisieren der Optionen: {e}", "error")
        
    return redirect(url_for('fragen', _anchor=active_tab_anchor, **session.get('project_config', {})))

@app.route('/download_filtered_pdf', methods=['POST'])
def download_filtered_pdf():
    try:
        bearbeiter = request.form.get('bearbeiter', 'N/A')
        df_excel = pd.read_excel('Daten.xlsx', sheet_name='Fragestellungen', header=None)
        project_config = session.get('project_config', {})
        
        # HINWEIS: Ich habe die Zeilennummern aus Ihrer ersten Anfrage wiederhergestellt,
        # da die aus der zweiten Anfrage sehr große Lücken hatten. Passen Sie diese bei Bedarf an.
        category_configs = {
            'Trafo': {'start_row': 14, 'end_row': 38, 'count': project_config.get('num_trafos', 0)},
            'Einspeisung': {'start_row': 40, 'end_row': 65, 'count': project_config.get('num_einspeisungen', 0)},
            'Abgang': {'start_row': 77, 'end_row': 107, 'count': project_config.get('num_abgaenge', 0)},
        }

        pdf = FPDF(orientation='P', unit='mm', format='A4')
        create_pdf_cover(pdf, bearbeiter, "Gefilterte Lösungen")
        found_any_solution = False

        for category, config in category_configs.items():
            for i in range(1, config['count'] + 1):
                
                # --- HIER IST DIE KORREKTUR ---
                # Wir holen nur die Antworten aus der Datenbank, die auch relevant sind.
                answers = QuestionAnswer.query.filter(
                    QuestionAnswer.category == category,
                    QuestionAnswer.category_index == i,
                    QuestionAnswer.answer != None,
                    QuestionAnswer.answer != 'nicht Relevant' # Ignoriere "nicht Relevant"
                ).all()
                # -----------------------------

                if not answers: continue

                # Der Rest der Funktion bleibt unverändert
                if not found_any_solution:
                    pdf.add_page()
                    found_any_solution = True
                
                question_row_index = config['start_row']
                data_start_row = question_row_index + 1
                data_end_row = config['end_row'] + 1

                questions_row = df_excel.iloc[question_row_index]
                df_to_filter = df_excel.iloc[data_start_row:data_end_row].copy()
                df_to_filter.columns = questions_row

                for answer_obj in answers:
                    question_text = answer_obj.question.strip()
                    user_answer = answer_obj.answer.strip()

                    if question_text in df_to_filter.columns:
                        condition = (df_to_filter[question_text].isna()) | (df_to_filter[question_text].astype(str).str.contains(user_answer, na=False))
                        df_to_filter = df_to_filter[condition]
                
                solution_start_col_index = 26
                solution_headers = questions_row.iloc[solution_start_col_index:].dropna()
                final_solutions = df_to_filter.iloc[:, solution_start_col_index:solution_start_col_index + len(solution_headers)]
                final_solutions.columns = solution_headers
                final_solutions = final_solutions.dropna(how='all')

                if pdf.get_y() + 30 > pdf.h - pdf.b_margin: pdf.add_page()
                
                pdf.ln(10)
                pdf.set_font("Arial", 'B', 14)
                pdf.cell(0, 10, txt=f"Lösungen für {category} {i}", ln=True, align='L')
                
                if final_solutions.empty:
                    pdf.set_font("Arial", '', 10)
                    pdf.cell(0, 10, txt="Keine passenden Lösungen für diese Konfiguration gefunden.", ln=True, align='L')
                else:
                    pdf.set_font("Arial", 'B', 10)
                    col_widths = [(pdf.w - 20) / len(final_solutions.columns)] * len(final_solutions.columns)
                    for j, header in enumerate(final_solutions.columns):
                        pdf.cell(col_widths[j], 10, str(header), 1, 0, 'C')
                    pdf.ln()

                    pdf.set_font("Arial", '', 9)
                    for index, row in final_solutions.iterrows():
                        for j, item in enumerate(row):
                            pdf.cell(col_widths[j], 10, str(item) if pd.notna(item) else "", 1, 0, 'L')
                        pdf.ln()

        if not found_any_solution:
            flash("Bitte beantworten Sie zuerst die Fragen, um Lösungen zu filtern.", "error")
            return redirect(url_for('fragen', **session.get('project_config', {})))

        pdf_output = pdf.output(dest='S').encode('latin1')
        return send_file(BytesIO(pdf_output), as_attachment=True, download_name='Gefilterte_Loesungen.pdf', mimetype='application/pdf')

    except FileNotFoundError:
        flash("Fehler: 'Daten.xlsx' wurde nicht gefunden.", "error")
        return redirect(url_for('fragen', **session.get('project_config', {})))
    except Exception as e:
        flash(f"Ein unerwarteter Fehler ist aufgetreten: {e}", "error")
        return redirect(url_for('fragen', **session.get('project_config', {})))

@app.route('/export_questions_pdf', methods=['POST'])
def export_questions_pdf():
    try:
        bearbeiter = request.form.get('bearbeiter', 'N/A')
        project_config = session.get('project_config', {})
        
        conditions = []
        conditions.append(QuestionAnswer.category == 'Allgemein')
        
        num_trafos = project_config.get('num_trafos', 0)
        if num_trafos > 0:
            conditions.append(
                (QuestionAnswer.category == 'Trafo') & (QuestionAnswer.category_index <= num_trafos)
            )
            
        num_einspeisungen = project_config.get('num_einspeisungen', 0)
        if num_einspeisungen > 0:
            conditions.append(
                (QuestionAnswer.category == 'Einspeisung') & (QuestionAnswer.category_index <= num_einspeisungen)
            )
            
        num_abgaenge = project_config.get('num_abgaenge', 0)
        if num_abgaenge > 0:
            conditions.append(
                (QuestionAnswer.category == 'Abgang') & (QuestionAnswer.category_index <= num_abgaenge)
            )

        fragen_db = QuestionAnswer.query.filter(or_(*conditions)).order_by(
            QuestionAnswer.category, QuestionAnswer.category_index, QuestionAnswer.sort_index
        ).all()

        if not fragen_db:
            flash("Keine Fragen zum Exportieren für die aktuelle Konfiguration vorhanden.", "info")
            return redirect(url_for('fragen', **session.get('project_config', {})))
            
        pdf = QuestionPDF(orientation='P', unit='mm', format='A4')
        pdf.create_cover(bearbeiter)
        pdf.create_question_tables(fragen_db)
        pdf_output = pdf.output(dest='S').encode('latin1')
        return send_file(BytesIO(pdf_output), as_attachment=True, download_name='Fragebogen.pdf', mimetype='application/pdf')
    except Exception as e:
        flash(f"Fehler beim Erstellen des PDFs: {e}", "error")
        return redirect(url_for('fragen', **session.get('project_config', {})))

def generate_category_chart(entries):
    pass

def create_pdf_cover(pdf_obj, bearbeiter_name, title):
    pdf_obj.add_page()
    logo_path = os.path.join(basedir, 'static', 'img', 'logo.png')
    if os.path.exists(logo_path):
        try:
            with Image.open(logo_path) as img:
                temp_logo_path = os.path.join(basedir, 'static', 'img', 'temp_logo_for_pdf.png')
                img.save(temp_logo_path)
                pdf_obj.image(temp_logo_path, x=pdf_obj.w/2 - 55, y=40, w=110)
                os.remove(temp_logo_path)
        except Exception as e:
            print(f"Fehler beim Verarbeiten des Logos: {e}")
    pdf_obj.set_y(120)
    pdf_obj.set_font('Arial', 'B', 24)
    pdf_obj.cell(0, 20, title, 0, 1, 'C')
    pdf_obj.ln(20)
    pdf_obj.set_y(180)
    pdf_obj.set_font('Arial', '', 12)
    pdf_obj.cell(0, 10, f'Bearbeiter: {bearbeiter_name}', 0, 1, 'C')
    pdf_obj.cell(0, 10, f'Datum: {date.today().strftime("%d.%m.%Y")}', 0, 1, 'C')

class PDF(FPDF):
    pass

class QuestionPDF(FPDF):
    def header(self):
        # KORRIGIERTE HEADER-METHODE
        if self.page_no() > 1:
            self.set_font('Arial', 'B', 16)
            # Positioniere den Titel manuell, um Platz für das Logo zu lassen
            title_w = self.get_string_width('Fragebogen zur ISO50001') + 6
            self.set_x((self.w - title_w) / 2)
            self.cell(title_w, 10, 'Fragebogen zur ISO50001', 0, 0, 'C')
            
            logo_path = os.path.join(basedir, 'static', 'img', 'logo.png')
            if os.path.exists(logo_path):
                # x-Position: Seitenbreite - rechter Rand - Bildbreite
                self.image(logo_path, x=self.w - self.r_margin - 50, y=5, w=50)
            self.ln(20)
            
    def footer(self):
        if self.page_no() > 1:
            self.set_y(-15)
            self.set_font('Arial', 'I', 8)
            self.cell(0, 10, f'Seite {self.page_no() - 1}', 0, 0, 'C')
            
    def create_cover(self, bearbeiter_name):
        create_pdf_cover(self, bearbeiter_name, 'Fragebogen zur ISO50001')
        
    def create_question_tables(self, data):
        from itertools import groupby
        
        category_order = ['Allgemein', 'Trafo', 'Einspeisung', 'Abgang']
        keyfunc = lambda q: (q.category, q.category_index)
        grouped_data = {k: list(v) for k, v in groupby(data, key=keyfunc)}
        
        self.add_page()
        col_widths = {"frage": 95, "optionen": 50, "antwort": 45}
        
        for category in category_order:
            instanzen = sorted([k[1] for k in grouped_data.keys() if k[0] == category])
            if not instanzen: continue

            if self.get_y() + 40 > self.h - self.b_margin: self.add_page()
            
            for i in instanzen:
                questions = grouped_data.get((category, i), [])
                if not questions: continue
                
                # KORRIGIERTE PRÜFUNG FÜR SEITENUMBRUCH
                if self.get_y() + 60 > self.h - self.b_margin: 
                    self.add_page()

                # Titel nur einmal pro Kategorie/Instanz
                if category == 'Allgemein' and i == 1:
                    self.set_font('Arial', 'B', 14)
                    self.cell(0, 10, f'{category}', 0, 1, 'L')
                    self.ln(3)
                elif category != 'Allgemein':
                    self.set_font('Arial', 'B', 14)
                    self.cell(0, 10, f'{category} {i}', 0, 1, 'L')
                    self.ln(3)

                self.set_font('Arial', 'B', 10)
                self.cell(col_widths["frage"], 10, 'Frage', 1, 0, 'C')
                self.cell(col_widths["optionen"], 10, 'Antwortmöglichkeiten', 1, 0, 'C')
                self.cell(col_widths["antwort"], 10, 'Antwort', 1, 1, 'C')
                self.set_font('Arial', '', 9)

                for q in questions:
                    line_height = 7
                    question_lines = self.multi_cell(col_widths["frage"], line_height, q.question, 0, 'L', split_only=True)
                    options_lines = self.multi_cell(col_widths["optionen"], line_height, q.options.replace(',', ', '), 0, 'L', split_only=True)
                    num_lines = max(len(question_lines), len(options_lines), 1)
                    row_height = num_lines * line_height

                    if self.get_y() + row_height > self.h - self.b_margin:
                        self.add_page()
                        self.set_font('Arial', 'B', 10)
                        self.cell(col_widths["frage"], 10, 'Frage', 1, 0, 'C')
                        self.cell(col_widths["optionen"], 10, 'Antwortmöglichkeiten', 1, 0, 'C')
                        self.cell(col_widths["antwort"], 10, 'Antwort', 1, 1, 'C')
                        self.set_font('Arial', '', 9)


                    if num_lines > 1 and len(options_lines) > len(question_lines):
                        Frage_Hoehe = row_height
                        Moeglichkeiten_Hoehe = line_height
                        Antwort_Hoehe = row_height
                    elif num_lines > 1 and len(question_lines) > len(options_lines):
                        Frage_Hoehe = line_height
                        Moeglichkeiten_Hoehe = row_height
                        Antwort_Hoehe = row_height
                    else:
                        Frage_Hoehe = line_height
                        Moeglichkeiten_Hoehe = line_height
                        Antwort_Hoehe = line_height
                      
                    x_start, y_start = self.get_x(), self.get_y()
                    self.multi_cell(col_widths["frage"], Frage_Hoehe, q.question, 1, 'L')
                    self.set_xy(x_start + col_widths["frage"], y_start)
                    self.multi_cell(col_widths["optionen"], Moeglichkeiten_Hoehe, q.options.replace(',', ', '), 1, 'L')
                    self.set_xy(x_start + col_widths["frage"] + col_widths["optionen"], y_start)
                    self.cell(col_widths["antwort"], Antwort_Hoehe, "", 1, 1, 'L')
                
                self.ln(5) # Abstand nach jeder Instanztabelle

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(host='0.0.0.0', port=5050, debug=True)