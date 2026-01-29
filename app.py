APP_NAME = "Локальная информационная система гигиенического обучения"
APP_VERSION = "v1.0 (stable)"
APP_YEAR = "2025"
APP_AUTHOR = "А. Бисеналин"

import threading
import webbrowser
import tkinter as tk
from werkzeug.serving import make_server
import os
import sys
import base64
import hashlib
from io import BytesIO
from datetime import datetime, timedelta

import qrcode
from openpyxl import Workbook

from flask import Flask, render_template, request, redirect, url_for, send_file
from flask_sqlalchemy import SQLAlchemy

import logging
from logging.handlers import RotatingFileHandler


# Папка, где лежит exe/скрипт (для portable-версии)
BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, "templates"),
    static_folder=os.path.join(BASE_DIR, "static"),
)

# БД рядом с exe (portable)
db_path = os.path.join(BASE_DIR, "hygiene.db")
app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{db_path}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

# Секрет для генерации контрольного кода (держать в секрете, можно поменять)
SECRET_KEY = "POL2KST-SECRET-2025"


# ----------------- ЛОГ ФАЙЛ ------------------


LOG_PATH = os.path.join(BASE_DIR, "error.log")

handler = RotatingFileHandler(
    LOG_PATH,
    maxBytes=2 * 1024 * 1024,  # 2 МБ
    backupCount=3,
    encoding="utf-8"
)
handler.setLevel(logging.ERROR)

formatter = logging.Formatter(
    "%(asctime)s | %(levelname)s | %(message)s"
)
handler.setFormatter(formatter)

app.logger.addHandler(handler)
app.logger.setLevel(logging.ERROR)


# ---------------- ОБРАБОТЧИК -----------------


@app.errorhandler(Exception)
def handle_exception(e):
    app.logger.error("Unhandled exception", exc_info=e)
    return "Произошла внутренняя ошибка. Обратитесь к администратору.", 500


# ----------------- МОДЕЛИ БД -----------------


class Program(db.Model):
    __tablename__ = "programs"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255), nullable=False)          # название программы
    category = db.Column(db.String(255), nullable=True)       # категория/группа (можно не использовать)
    theory_hours = db.Column(db.Integer, nullable=True)       # часы теории
    exam_hours = db.Column(db.Integer, nullable=True)         # часы экзамена

    trainings = db.relationship(
        "Training",
        back_populates="program",
        cascade="all, delete-orphan",
    )


class Participant(db.Model):
    __tablename__ = "participants"

    id = db.Column(db.Integer, primary_key=True)
    iin = db.Column(db.String(12), nullable=True)
    full_name = db.Column(db.String(255), nullable=False)
    birth_date = db.Column(db.Date, nullable=True)
    sex = db.Column(db.String(10), nullable=True)

    lmk_number = db.Column(db.String(100), nullable=True)     # № ЛМК
    workplace = db.Column(db.String(255), nullable=True)      # место работы
    position = db.Column(db.String(255), nullable=True)       # должность
    activity_type = db.Column(db.String(255), nullable=True)  # вид деятельности/услуг

    trainings = db.relationship(
        "Training",
        back_populates="participant",
        cascade="all, delete-orphan",
    )


class Training(db.Model):
    __tablename__ = "trainings"

    id = db.Column(db.Integer, primary_key=True)
    participant_id = db.Column(db.Integer, db.ForeignKey("participants.id"), nullable=False)
    program_id = db.Column(db.Integer, db.ForeignKey("programs.id"), nullable=False)

    training_start_date = db.Column(db.Date, nullable=True)
    training_end_date = db.Column(db.Date, nullable=True)
    exam_date = db.Column(db.Date, nullable=False)

    questions_total = db.Column(db.Integer, nullable=False, default=0)
    correct_answers = db.Column(db.Integer, nullable=False, default=0)

    exam_percent = db.Column(db.Float, nullable=False, default=0.0)
    exam_result = db.Column(db.String(50), nullable=False, default="Отрицательный")
    next_exam_date = db.Column(db.Date, nullable=True)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    participant = db.relationship("Participant", back_populates="trainings")
    program = db.relationship("Program", back_populates="trainings")


# -------------- УТИЛИТЫ ----------------------


def generate_control_hash(training: Training) -> str:
    """
    Создаёт криптографический хэш SHA-256 для свидетельства.
    Включаем ID, ФИО, дату экзамена, результат и секретный ключ.
    """
    data_string = (
        f"{training.id}|"
        f"{training.participant.full_name}|"
        f"{training.exam_date.strftime('%Y-%m-%d')}|"
        f"{training.exam_result}|"
        f"POL2KST|"
        f"{SECRET_KEY}"
    )

    return hashlib.sha256(data_string.encode("utf-8")).hexdigest()


def generate_qr_base64(data: str) -> str:
    """
    Генерирует QR-код по строке data и возвращает base64-строку
    для вставки в <img src="data:image/png;base64,...">
    """
    qr = qrcode.QRCode(
        version=2,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=4,
        border=2,
    )
    qr.add_data(data)
    qr.make(fit=True)

    img = qr.make_image(fill_color="black", back_color="white")

    buf = BytesIO()
    img.save(buf, format="PNG")
    img_bytes = buf.getvalue()
    return base64.b64encode(img_bytes).decode("utf-8")


# ----------- ИНИЦИАЛИЗАЦИЯ БД --------------

with app.app_context():
    db.create_all()


# ----------------- ГЛАВНАЯ ------------------


@app.route("/")
def index():
    return redirect(url_for("list_trainings"))


# -------- СПРАВОЧНИК ПРОГРАММ ---------------


@app.route("/programs")
def list_programs():
    programs = Program.query.order_by(Program.name).all()
    return render_template("programs.html", programs=programs)


@app.route("/programs/new", methods=["GET", "POST"])
def new_program():
    if request.method == "POST":
        name = request.form.get("name")
        category = request.form.get("category")
        theory_hours = request.form.get("theory_hours") or None
        exam_hours = request.form.get("exam_hours") or None

        prog = Program(
            name=name,
            category=category,
            theory_hours=int(theory_hours) if theory_hours else None,
            exam_hours=int(exam_hours) if exam_hours else None,
        )
        db.session.add(prog)
        db.session.commit()
        return redirect(url_for("list_programs"))

    return render_template(
        "program_form.html",
        program=None,
        action_url=url_for("new_program"),
        submit_label="Сохранить",
    )


@app.route("/programs/edit/<int:program_id>", methods=["GET", "POST"])
def edit_program(program_id):
    prog = Program.query.get_or_404(program_id)

    if request.method == "POST":
        prog.name = request.form.get("name")
        prog.category = request.form.get("category")
        theory_hours = request.form.get("theory_hours") or None
        exam_hours = request.form.get("exam_hours") or None
        prog.theory_hours = int(theory_hours) if theory_hours else None
        prog.exam_hours = int(exam_hours) if exam_hours else None
        db.session.commit()
        return redirect(url_for("list_programs"))

    return render_template(
        "program_form.html",
        program=prog,
        action_url=url_for("edit_program", program_id=program_id),
        submit_label="Обновить",
    )


@app.route("/programs/delete/<int:program_id>", methods=["POST"])
def delete_program(program_id):
    prog = Program.query.get_or_404(program_id)
    db.session.delete(prog)
    db.session.commit()
    return redirect(url_for("list_programs"))


# --------- СПРАВОЧНИК СЛУШАТЕЛЕЙ ------------


@app.route("/participants")
def list_participants():
    participants = Participant.query.order_by(Participant.full_name).all()
    return render_template("participants.html", participants=participants)


@app.route("/participants/new", methods=["GET", "POST"])
def new_participant():
    if request.method == "POST":
        full_name = request.form.get("full_name")
        iin = request.form.get("iin")
        birth_date_str = request.form.get("birth_date")
        sex = request.form.get("sex")
        lmk_number = request.form.get("lmk_number")
        workplace = request.form.get("workplace")
        position = request.form.get("position")
        activity_type = request.form.get("activity_type")

        birth_date = None
        if birth_date_str:
            birth_date = datetime.strptime(birth_date_str, "%Y-%m-%d").date()

        p = Participant(
            full_name=full_name,
            iin=iin,
            birth_date=birth_date,
            sex=sex,
            lmk_number=lmk_number,
            workplace=workplace,
            position=position,
            activity_type=activity_type,
        )
        db.session.add(p)
        db.session.commit()
        return redirect(url_for("list_participants"))

    return render_template(
        "participant_form.html",
        participant=None,
        action_url=url_for("new_participant"),
        submit_label="Сохранить",
    )


@app.route("/participants/edit/<int:participant_id>", methods=["GET", "POST"])
def edit_participant(participant_id):
    p = Participant.query.get_or_404(participant_id)

    if request.method == "POST":
        p.full_name = request.form.get("full_name")
        p.iin = request.form.get("iin")
        birth_date_str = request.form.get("birth_date")
        p.sex = request.form.get("sex")
        p.lmk_number = request.form.get("lmk_number")
        p.workplace = request.form.get("workplace")
        p.position = request.form.get("position")
        p.activity_type = request.form.get("activity_type")

        if birth_date_str:
            p.birth_date = datetime.strptime(birth_date_str, "%Y-%m-%d").date()
        else:
            p.birth_date = None

        db.session.commit()
        return redirect(url_for("list_participants"))

    return render_template(
        "participant_form.html",
        participant=p,
        action_url=url_for("edit_participant", participant_id=participant_id),
        submit_label="Обновить",
    )


@app.route("/participants/delete/<int:participant_id>", methods=["POST"])
def delete_participant(participant_id):
    p = Participant.query.get_or_404(participant_id)
    db.session.delete(p)
    db.session.commit()
    return redirect(url_for("list_participants"))


# ------------ ОБУЧЕНИЯ / ЭКЗАМЕНЫ ------------


@app.route("/trainings")
def list_trainings():
    trainings = (
        Training.query
        .order_by(Training.exam_date.desc(), Training.id.desc())
        .all()
    )
    return render_template("trainings.html", trainings=trainings)


@app.route("/trainings/new", methods=["GET", "POST"])
def new_training():
    participants = Participant.query.order_by(Participant.full_name).all()
    programs = Program.query.order_by(Program.name).all()

    if request.method == "POST":
        participant_id = int(request.form.get("participant_id"))
        program_id = int(request.form.get("program_id"))

        training_start_date_str = request.form.get("training_start_date")
        training_end_date_str = request.form.get("training_end_date")
        exam_date_str = request.form.get("exam_date")

        questions_total = int(request.form.get("questions_total"))
        correct_answers = int(request.form.get("correct_answers"))

        training_start_date = (
            datetime.strptime(training_start_date_str, "%Y-%m-%d").date()
            if training_start_date_str else None
        )
        training_end_date = (
            datetime.strptime(training_end_date_str, "%Y-%m-%d").date()
            if training_end_date_str else None
        )
        exam_date = datetime.strptime(exam_date_str, "%Y-%m-%d").date()

        if questions_total > 0:
            exam_percent = round(correct_answers / questions_total * 100, 1)
        else:
            exam_percent = 0.0

        exam_result = "Положительный" if exam_percent >= 80.0 else "Отрицательный"

        next_exam_date = None
        if exam_result == "Положительный":
            try:
                next_exam_date = exam_date.replace(year=exam_date.year + 1)
            except ValueError:
                next_exam_date = exam_date + timedelta(days=365)

        t = Training(
            participant_id=participant_id,
            program_id=program_id,
            training_start_date=training_start_date,
            training_end_date=training_end_date,
            exam_date=exam_date,
            questions_total=questions_total,
            correct_answers=correct_answers,
            exam_percent=exam_percent,
            exam_result=exam_result,
            next_exam_date=next_exam_date,
        )
        db.session.add(t)
        db.session.commit()
        return redirect(url_for("list_trainings"))

    return render_template(
        "training_form.html",
        training=None,
        participants=participants,
        programs=programs,
        action_url=url_for("new_training"),
        submit_label="Сохранить",
    )


@app.route("/trainings/edit/<int:training_id>", methods=["GET", "POST"])
def edit_training(training_id):
    training = Training.query.get_or_404(training_id)
    participants = Participant.query.order_by(Participant.full_name).all()
    programs = Program.query.order_by(Program.name).all()

    if request.method == "POST":
        training.participant_id = int(request.form.get("participant_id"))
        training.program_id = int(request.form.get("program_id"))

        training_start_date_str = request.form.get("training_start_date")
        training_end_date_str = request.form.get("training_end_date")
        exam_date_str = request.form.get("exam_date")

        training.training_start_date = (
            datetime.strptime(training_start_date_str, "%Y-%m-%d").date()
            if training_start_date_str else None
        )
        training.training_end_date = (
            datetime.strptime(training_end_date_str, "%Y-%m-%d").date()
            if training_end_date_str else None
        )
        training.exam_date = datetime.strptime(exam_date_str, "%Y-%m-%d").date()

        training.questions_total = int(request.form.get("questions_total"))
        training.correct_answers = int(request.form.get("correct_answers"))

        if training.questions_total > 0:
            training.exam_percent = round(
                training.correct_answers / training.questions_total * 100, 1
            )
        else:
            training.exam_percent = 0.0

        training.exam_result = (
            "Положительный" if training.exam_percent >= 80.0 else "Отрицательный"
        )

        training.next_exam_date = None
        if training.exam_result == "Положительный":
            try:
                training.next_exam_date = training.exam_date.replace(
                    year=training.exam_date.year + 1
                )
            except ValueError:
                training.next_exam_date = training.exam_date + timedelta(days=365)

        db.session.commit()
        return redirect(url_for("list_trainings"))

    return render_template(
        "training_form.html",
        training=training,
        participants=participants,
        programs=programs,
        action_url=url_for("edit_training", training_id=training_id),
        submit_label="Обновить",
    )


@app.route("/trainings/delete/<int:training_id>", methods=["POST"])
def delete_training(training_id):
    training = Training.query.get_or_404(training_id)
    db.session.delete(training)
    db.session.commit()
    return redirect(url_for("list_trainings"))


# ----------------- ЖУРНАЛ -------------------


@app.route("/journal")
def journal():
    trainings = (
        Training.query
        .order_by(Training.exam_date.asc(), Training.id.asc())
        .all()
    )
    return render_template("journal.html", trainings=trainings)


@app.route("/journal/excel")
def journal_excel():
    trainings = (
        Training.query
        .order_by(Training.exam_date.asc(), Training.id.asc())
        .all()
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "Журнал"

    headers = [
        "№",
        "ФИО слушателя",
        "Место работы, должность, вид деятельности/услуг",
        "Период обучения",
        "Дата экзамена",
        "Результат экзамена",
        "Дата очередного экзамена",
    ]
    ws.append(headers)

    for idx, t in enumerate(trainings, start=1):
        if t.training_start_date and t.training_end_date:
            period = (
                f"{t.training_start_date.strftime('%d.%m.%Y')} – "
                f"{t.training_end_date.strftime('%d.%m.%Y')}"
            )
        else:
            period = ""

        workplace_block = t.participant.workplace or ""
        position_block = t.participant.position or ""
        activity_block = t.participant.activity_type or ""

        place_full = workplace_block
        if position_block:
            place_full += f", {position_block}"
        if activity_block:
            place_full += f", {activity_block}"

        exam_date_str = t.exam_date.strftime("%d.%m.%Y") if t.exam_date else ""
        result_str = f"{t.exam_result} ({t.exam_percent:.1f} %)"
        next_exam_str = (
            t.next_exam_date.strftime("%d.%m.%Y") if t.next_exam_date else ""
        )

        row = [
            idx,
            t.participant.full_name,
            place_full,
            period,
            exam_date_str,
            result_str,
            next_exam_str,
        ]
        ws.append(row)

    for column_cells in ws.columns:
        length = max(
            len(str(cell.value)) if cell.value is not None else 0
            for cell in column_cells
        )
        ws.column_dimensions[column_cells[0].column_letter].width = length + 2

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)

    return send_file(
    stream,
    as_attachment=True,
    download_name="journal.xlsx",
    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)



# -------------- СВИДЕТЕЛЬСТВО ---------------


@app.route("/certificate/<int:training_id>")
def certificate(training_id):
    training = Training.query.get_or_404(training_id)

    # Контрольный код (аналог "цифровой подписи" документа)
    control_hash = generate_control_hash(training)

    # Данные, включаемые в QR-код
    qr_data = (
        f"КГП \"Поликлиника № 2 города Костанай\"\n"
        f"Свидетельство № {training.id}\n"
        f"ФИО: {training.participant.full_name}\n"
        f"Дата экзамена: {training.exam_date.strftime('%d.%m.%Y')}\n"
        f"Результат: {training.exam_result} ({training.exam_percent:.1f} %)\n"
        f"Контрольный код: {control_hash}"
    )

    qr_image = generate_qr_base64(qr_data)

    return render_template(
        "certificate.html",
        training=training,
        qr_image=qr_image,
        control_hash=control_hash,
    )


# -------------- КЛАСС СЕРВЕРА  --------------


class ServerThread(threading.Thread):
    def __init__(self, flask_app, host="127.0.0.1", port=5000):
        super().__init__(daemon=True)
        self.host = host
        self.port = port
        self.srv = make_server(host, port, flask_app)
        self.ctx = flask_app.app_context()
        self.ctx.push()

    def run(self):
        self.srv.serve_forever()

    def shutdown(self):
        self.srv.shutdown()


# ----------------- ЗАПУСК -------------------


if __name__ == "__main__":
    HOST = "127.0.0.1"
    PORT = 5000
    URL = f"http://{HOST}:{PORT}"

    # Запускаем сервер в отдельном потоке
    server = ServerThread(app, host=HOST, port=PORT)
    server.start()

    # Окно управления
    root = tk.Tk()
    root.title("Гигиеническое обучение — ИС")
    root.geometry("420x170")
    root.resizable(False, False)

    lbl1 = tk.Label(root, text="Информационная система запущена", font=("Segoe UI", 12, "bold"))
    lbl1.pack(pady=(18, 6))

    lbl2 = tk.Label(root, text=URL, font=("Segoe UI", 10))
    lbl2.pack(pady=(0, 14))

    btn_frame = tk.Frame(root)
    btn_frame.pack()

    def open_site():
        webbrowser.open(URL)

    def stop_app():
        try:
            server.shutdown()
        finally:
            root.destroy()

    btn_open = tk.Button(btn_frame, text="Открыть в браузере", width=18, command=open_site)
    btn_open.grid(row=0, column=0, padx=8)

    btn_stop = tk.Button(btn_frame, text="Завершить работу", width=18, command=stop_app)
    btn_stop.grid(row=0, column=1, padx=8)

    # Автооткрытие браузера через 1 сек (по желанию)
    root.after(1000, open_site)

    # Если пользователь закрывает окно крестиком — тоже корректно выключаем
    root.protocol("WM_DELETE_WINDOW", stop_app)

    root.mainloop()
