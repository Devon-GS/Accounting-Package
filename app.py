from __future__ import annotations

import hashlib
import os
import re
import sqlite3
from contextlib import closing
from datetime import date, datetime
from io import StringIO
from pathlib import Path
from typing import Iterable

from flask import (
	Flask,
	flash,
	g,
	redirect,
	render_template,
	request,
	url_for,
)
from ofxparse import OfxParser
from openpyxl import load_workbook
from werkzeug.utils import secure_filename


BASE_DIR = Path(__file__).resolve().parent
INSTANCE_DIR = BASE_DIR / "instance"
UPLOAD_DIR = INSTANCE_DIR / "uploads"
DB_PATH = INSTANCE_DIR / "accounting.db"
ACCOUNT_NAMES_PATH = BASE_DIR / "Account Names.xlsx"

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}
PAYMENT_COLUMNS = (2, 4)
PAYMENT_RANGE = range(18, 34)

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-accounting-secret")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024


def normalize_text(value: object) -> str:
	text = str(value or "").strip().upper()
	return re.sub(r"[^A-Z0-9]+", " ", text).strip()


def parse_money(value: object) -> float | None:
	if value in (None, ""):
		return None
	if isinstance(value, str):
		cleaned = value.replace(" ", "").replace(",", "")
		if not cleaned:
			return None
		try:
			number = float(cleaned)
		except ValueError:
			return None
	else:
		try:
			number = float(value)
		except (TypeError, ValueError):
			return None
	if abs(number) < 1e-9:
		return None
	return round(number, 2)


def parse_filename_date(filename: str) -> str:
	name = Path(filename).stem
	patterns = [
		r"(?P<day>\d{1,2})[\s_\-\.]+(?P<month>[A-Za-z]+)[\s_\-\.]+(?P<year>\d{4})",
		r"(?P<day>\d{1,2})[\s_\-\.]+(?P<month>\d{1,2})[\s_\-\.]+(?P<year>\d{4})",
	]
	month_lookup = {
		"JAN": 1,
		"JANUARY": 1,
		"FEB": 2,
		"FEBRUARY": 2,
		"MAR": 3,
		"MARCH": 3,
		"APR": 4,
		"APRIL": 4,
		"MAY": 5,
		"JUN": 6,
		"JUNE": 6,
		"JUL": 7,
		"JULY": 7,
		"AUG": 8,
		"AUGUST": 8,
		"SEP": 9,
		"SEPT": 9,
		"SEPTEMBER": 9,
		"OCT": 10,
		"OCTOBER": 10,
		"NOV": 11,
		"NOVEMBER": 11,
		"DEC": 12,
		"DECEMBER": 12,
	}
	for pattern in patterns:
		match = re.search(pattern, name, flags=re.IGNORECASE)
		if not match:
			continue
		day = int(match.group("day"))
		year = int(match.group("year"))
		month_text = match.group("month")
		if month_text.isdigit():
			month = int(month_text)
		else:
			month = month_lookup.get(month_text.strip().upper())
		if month is None:
			continue
		return datetime(year, month, day).date().isoformat()
	raise ValueError(
		f"Could not find a date in the filename '{filename}'. "
		"Expected something like 'Control Sheet Daily 01 JULY 2026.xlsx'."
	)


def allowed_file(filename: str) -> bool:
	return Path(filename).suffix.lower() in SUPPORTED_EXTENSIONS


def get_connection() -> sqlite3.Connection:
	if "db" not in g:
		conn = sqlite3.connect(DB_PATH)
		conn.row_factory = sqlite3.Row
		conn.execute("PRAGMA foreign_keys = ON")
		g.db = conn
	return g.db


def query_one(sql: str, params: tuple = ()) -> sqlite3.Row | None:
	return get_connection().execute(sql, params).fetchone()


def query_all(sql: str, params: tuple = ()) -> list[sqlite3.Row]:
	return get_connection().execute(sql, params).fetchall()


def execute(sql: str, params: tuple = ()) -> sqlite3.Cursor:
	conn = get_connection()
	cur = conn.execute(sql, params)
	conn.commit()
	return cur


def initialize_storage() -> None:
	INSTANCE_DIR.mkdir(exist_ok=True)
	UPLOAD_DIR.mkdir(exist_ok=True)
	with closing(sqlite3.connect(DB_PATH)) as conn:
		conn.row_factory = sqlite3.Row
		conn.execute("PRAGMA foreign_keys = ON")
		conn.executescript(
			"""
			CREATE TABLE IF NOT EXISTS accounts (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				name TEXT NOT NULL UNIQUE,
				source_sheet TEXT,
				created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
			);

			CREATE TABLE IF NOT EXISTS supplier_rules (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				supplier_key TEXT NOT NULL UNIQUE,
				supplier_name TEXT NOT NULL,
				account_id INTEGER NOT NULL,
				created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
				FOREIGN KEY (account_id) REFERENCES accounts (id) ON DELETE CASCADE
			);

			CREATE TABLE IF NOT EXISTS import_batches (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
				source_folder TEXT,
				source_type TEXT NOT NULL DEFAULT 'cash',
				status TEXT NOT NULL DEFAULT 'open'
			);

			CREATE TABLE IF NOT EXISTS import_files (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				batch_id INTEGER NOT NULL,
				original_filename TEXT NOT NULL,
				stored_filename TEXT NOT NULL,
				file_hash TEXT NOT NULL UNIQUE,
				file_date TEXT NOT NULL,
				created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
				FOREIGN KEY (batch_id) REFERENCES import_batches (id) ON DELETE CASCADE
			);

			CREATE TABLE IF NOT EXISTS staged_payments (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				batch_id INTEGER NOT NULL,
				file_id INTEGER NOT NULL,
				payment_date TEXT NOT NULL,
				supplier_name TEXT NOT NULL,
				supplier_key TEXT NOT NULL,
				amount_paid REAL NOT NULL,
				source_sheet TEXT NOT NULL,
				source_row INTEGER NOT NULL,
				source_column INTEGER NOT NULL,
				account_id INTEGER,
				created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
				UNIQUE (file_id, source_sheet, source_row, source_column),
				FOREIGN KEY (batch_id) REFERENCES import_batches (id) ON DELETE CASCADE,
				FOREIGN KEY (file_id) REFERENCES import_files (id) ON DELETE CASCADE,
				FOREIGN KEY (account_id) REFERENCES accounts (id) ON DELETE SET NULL
			);

			CREATE TABLE IF NOT EXISTS payments (
				id INTEGER PRIMARY KEY AUTOINCREMENT,
				payment_date TEXT NOT NULL,
				supplier_name TEXT NOT NULL,
				amount_paid REAL NOT NULL,
				account_id INTEGER NOT NULL,
				source_file_name TEXT NOT NULL,
				source_file_hash TEXT NOT NULL,
				source_sheet TEXT NOT NULL,
				source_row INTEGER NOT NULL,
				source_column INTEGER NOT NULL,
				created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
				UNIQUE (source_file_hash, source_sheet, source_row, source_column),
				FOREIGN KEY (account_id) REFERENCES accounts (id) ON DELETE RESTRICT
			);
			"""
		)
		columns = {row["name"] for row in conn.execute("PRAGMA table_info(import_batches)")}
		if "source_type" not in columns:
			conn.execute("ALTER TABLE import_batches ADD COLUMN source_type TEXT NOT NULL DEFAULT 'cash'")
		conn.execute(
			"UPDATE import_batches SET source_type = 'cash' WHERE source_type IS NULL OR TRIM(source_type) = ''"
		)
		conn.commit()


def sync_account_names() -> dict[str, int]:
    created: dict[str, int] = {}
    if not ACCOUNT_NAMES_PATH.exists():
        return created

    wb = load_workbook(ACCOUNT_NAMES_PATH, data_only=True, read_only=True)
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_col=1, max_col=1):
            cell = row[0]
            if cell.value in (None, "", "."):
                continue
            name = str(cell.value).strip()
            if not name:
                continue
            existing = find_account_by_normalized_name(name)
            if existing:
                created[name] = existing["id"]
                continue
            cursor = execute(
                "INSERT OR IGNORE INTO accounts (name, source_sheet) VALUES (?, ?)",
                (name, ws.title),
            )
            account_id = cursor.lastrowid
            if not account_id:
                row = query_one("SELECT id FROM accounts WHERE name = ?", (name,))
                account_id = row["id"] if row else None
            if account_id:
                created[name] = account_id
    return created


def find_account_by_normalized_name(name: str) -> sqlite3.Row | None:
    target = normalize_text(name)
    if not target:
        return None
    for account in query_all("SELECT id, name FROM accounts"):
        if normalize_text(account["name"]) == target:
            return account
    return None


def canonical_account_name(name: str) -> str:
    return str(name or "").strip().upper()


def ensure_rule(supplier_name: str, account_id: int) -> None:
	key = normalize_text(supplier_name)
	execute(
		"""
		INSERT INTO supplier_rules (supplier_key, supplier_name, account_id)
		VALUES (?, ?, ?)
		ON CONFLICT(supplier_key) DO UPDATE SET
			supplier_name = excluded.supplier_name,
			account_id = excluded.account_id
		""",
		(key, supplier_name.strip(), account_id),
	)


def resolve_account_for_supplier(supplier_name: str) -> sqlite3.Row | None:
	supplier_key = normalize_text(supplier_name)
	rule = query_one(
		"""
		SELECT supplier_rules.account_id, accounts.name
		FROM supplier_rules
		JOIN accounts ON accounts.id = supplier_rules.account_id
		WHERE supplier_rules.supplier_key = ?
		""",
		(supplier_key,),
	)
	if rule:
		return rule

	exact = find_account_by_normalized_name(supplier_name)
	if exact:
		ensure_rule(supplier_name, exact["id"])
		return exact
	return None


def sha256_file(path: Path) -> str:
	digest = hashlib.sha256()
	with path.open("rb") as handle:
		for chunk in iter(lambda: handle.read(8192), b""):
			digest.update(chunk)
	return digest.hexdigest()


def pick_payment_sheet(workbook):
	for worksheet in workbook.worksheets:
		if normalize_text(worksheet.title) == "CASH UP":
			return worksheet
	return workbook.worksheets[0]


def stage_workbook(
    batch_id: int,
    file_storage,
    original_filename: str,
    file_date: str,
) -> tuple[int | None, int, bool]:
    source_path = Path(file_storage.filename)
    safe_name = secure_filename(source_path.name) or "upload.xlsx"
    temp_path = UPLOAD_DIR / f"batch{batch_id}_{safe_name}"
    file_storage.save(temp_path)

    try:
        file_hash = sha256_file(temp_path)
        existing = query_one("SELECT id FROM import_files WHERE file_hash = ?", (file_hash,))
        if existing:
            temp_path.unlink(missing_ok=True)
            return None, 0, True

        stored_name = f"{file_hash[:12]}_{safe_name}"
        stored_path = UPLOAD_DIR / stored_name
        temp_path.replace(stored_path)

        cursor = execute(
            """
            INSERT INTO import_files (batch_id, original_filename, stored_filename, file_hash, file_date)
            VALUES (?, ?, ?, ?, ?)
            """,
            (batch_id, original_filename, stored_name, file_hash, file_date),
        )
        file_id = cursor.lastrowid
        workbook = load_workbook(stored_path, data_only=True, read_only=True)
        worksheet = pick_payment_sheet(workbook)
        staged_rows = 0

        for row_number in PAYMENT_RANGE:
            supplier_name = worksheet.cell(row=row_number, column=1).value
            if supplier_name in (None, ""):
                continue
            supplier_name = str(supplier_name).strip()
            if not supplier_name:
                continue
            supplier_key = normalize_text(supplier_name)
            for column_number in PAYMENT_COLUMNS:
                amount = parse_money(worksheet.cell(row=row_number, column=column_number).value)
                if amount is None:
                    continue
                account = resolve_account_for_supplier(supplier_name)
                account_id = account["account_id"] if account and "account_id" in account.keys() else None
                if account is not None and "id" in account.keys():
                    account_id = account["id"]
                execute(
                    """
                    INSERT OR IGNORE INTO staged_payments (
                        batch_id, file_id, payment_date, supplier_name, supplier_key,
                        amount_paid, source_sheet, source_row, source_column, account_id
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        batch_id,
                        file_id,
                        file_date,
                        supplier_name,
                        supplier_key,
                        amount,
                        worksheet.title,
                        row_number,
                        column_number,
                        account_id,
                    ),
                )
                staged_rows += 1

        finalize_resolved_rows(batch_id)
        return file_id, staged_rows, False
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def date_to_iso(value: object) -> str | None:
	if value is None:
		return None
	if isinstance(value, datetime):
		return value.date().isoformat()
	if isinstance(value, date):
		return value.isoformat()
	if isinstance(value, str):
		for fmt in ("%Y-%m-%d", "%Y-%m-%dT%H:%M:%S", "%Y%m%d", "%Y%m%d%H%M%S"):
			try:
				return datetime.strptime(value, fmt).date().isoformat()
			except ValueError:
				continue
	return None


def parse_ofx_transactions(file_path: Path) -> list[dict[str, object]]:
	raw_text = file_path.read_text(encoding="latin-1")
	raw_text = re.sub(r"CHARSET:NONE", "CHARSET:1252", raw_text, flags=re.IGNORECASE)
	with StringIO(raw_text) as handle:
		ofx = OfxParser.parse(handle)

	transactions = []

	account_list = getattr(ofx, "accounts", None) or []
	if account_list:
		for account in account_list:
			statement = getattr(account, "statement", None)
			if statement:
				transactions.extend(getattr(statement, "transactions", []) or [])
	else:
		statement = getattr(ofx, "statement", None)
		if statement:
			transactions.extend(getattr(statement, "transactions", []) or [])

	parsed = []
	for index, txn in enumerate(transactions, start=1):
		description = ""
		for attr in ("payee", "memo", "name", "type"):
			value = getattr(txn, attr, None)
			if value:
				description = str(value).strip()
				break
		amount = parse_money(getattr(txn, "amount", None))
		payment_date = date_to_iso(getattr(txn, "date", None)) or date.today().isoformat()
		if not description or amount is None:
			continue
		if amount >= 0:
			continue
		parsed.append(
			{
				"index": index,
				"description": description,
				"amount": abs(amount),
				"payment_date": payment_date,
			}
		)
	return parsed


def stage_ofx_file(
	batch_id: int,
	file_storage,
	original_filename: str,
) -> tuple[int | None, int, bool]:
	source_path = Path(file_storage.filename)
	safe_name = secure_filename(source_path.name) or "statement.ofx"
	temp_path = UPLOAD_DIR / f"batch{batch_id}_{safe_name}"
	file_storage.save(temp_path)

	try:
		file_hash = sha256_file(temp_path)
		existing = query_one("SELECT id FROM import_files WHERE file_hash = ?", (file_hash,))
		if existing:
			temp_path.unlink(missing_ok=True)
			return None, 0, True

		transactions = parse_ofx_transactions(temp_path)
		file_date = transactions[0]["payment_date"] if transactions else date.today().isoformat()

		stored_name = f"{file_hash[:12]}_{safe_name}"
		stored_path = UPLOAD_DIR / stored_name
		temp_path.replace(stored_path)

		cursor = execute(
			"""
			INSERT INTO import_files (batch_id, original_filename, stored_filename, file_hash, file_date)
			VALUES (?, ?, ?, ?, ?)
			""",
			(batch_id, original_filename, stored_name, file_hash, file_date),
		)
		file_id = cursor.lastrowid
		staged_rows = 0

		for txn in transactions:
			description = str(txn["description"]).strip()
			amount = parse_money(txn["amount"])
			payment_date = str(txn["payment_date"])
			if not description or amount is None:
				continue
			account = resolve_account_for_supplier(description)
			account_id = account["id"] if account is not None and "id" in account.keys() else None
			if account is not None and "account_id" in account.keys():
				account_id = account["account_id"]
			execute(
				"""
				INSERT OR IGNORE INTO staged_payments (
					batch_id, file_id, payment_date, supplier_name, supplier_key,
					amount_paid, source_sheet, source_row, source_column, account_id
				)
				VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
				""",
				(
					batch_id,
					file_id,
					payment_date,
					description,
					normalize_text(description),
					amount,
					"OFX",
					int(txn["index"]),
					1,
					account_id,
				),
			)
			staged_rows += 1

		finalize_resolved_rows(batch_id)
		return file_id, staged_rows, False
	except Exception:
		temp_path.unlink(missing_ok=True)
		raise


def finalize_resolved_rows(batch_id: int) -> int:
	conn = get_connection()
	cursor = conn.execute(
		"""
		INSERT OR IGNORE INTO payments (
			payment_date, supplier_name, amount_paid, account_id,
			source_file_name, source_file_hash, source_sheet, source_row, source_column
		)
		SELECT
			staged_payments.payment_date,
			staged_payments.supplier_name,
			staged_payments.amount_paid,
			staged_payments.account_id,
			import_files.original_filename,
			import_files.file_hash,
			staged_payments.source_sheet,
			staged_payments.source_row,
			staged_payments.source_column
		FROM staged_payments
		JOIN import_files ON import_files.id = staged_payments.file_id
		WHERE staged_payments.batch_id = ?
		  AND staged_payments.account_id IS NOT NULL
		""",
		(batch_id,),
	)
	conn.commit()
	return cursor.rowcount if cursor.rowcount != -1 else 0


def latest_batch_id(source_type: str | None = None) -> int | None:
	if source_type:
		row = query_one(
			"SELECT id FROM import_batches WHERE source_type = ? ORDER BY id DESC LIMIT 1",
			(source_type,),
		)
	else:
		row = query_one("SELECT id FROM import_batches ORDER BY id DESC LIMIT 1")
	return row["id"] if row else None


def normalize_month(value: str | None) -> str:
	if not value:
		return ""
	try:
		return datetime.strptime(value, "%Y-%m").strftime("%Y-%m")
	except ValueError:
		return ""


def latest_payment_month() -> str | None:
	row = query_one(
		"""
		SELECT strftime('%Y-%m', payment_date) AS month_value
		FROM payments
		WHERE payment_date IS NOT NULL
		ORDER BY payment_date DESC
		LIMIT 1
		"""
	)
	return row["month_value"] if row and row["month_value"] else None


def available_payment_months() -> list[sqlite3.Row]:
	return query_all(
		"""
		SELECT strftime('%Y-%m', payment_date) AS month_value
		FROM payments
		WHERE payment_date IS NOT NULL
		GROUP BY month_value
		ORDER BY month_value DESC
		"""
	)


def month_label(month_value: str) -> str:
	return datetime.strptime(month_value, "%Y-%m").strftime("%B %Y")


def batch_summary(batch_id: int) -> dict:
	counts = query_one(
		"""
		SELECT
			COUNT(*) AS staged_total,
			SUM(CASE WHEN account_id IS NULL THEN 1 ELSE 0 END) AS unresolved_total,
			SUM(CASE WHEN account_id IS NOT NULL THEN 1 ELSE 0 END) AS resolved_total
		FROM staged_payments
		WHERE batch_id = ?
		""",
		(batch_id,),
	)
	files = query_all(
		"""
		SELECT original_filename, file_date, stored_filename
		FROM import_files
		WHERE batch_id = ?
		ORDER BY id DESC
		""",
		(batch_id,),
	)
	unresolved = query_all(
		"""
		SELECT
			staged_payments.id AS staged_id,
			staged_payments.batch_id,
			staged_payments.payment_date,
			staged_payments.supplier_name,
			staged_payments.supplier_key,
			staged_payments.amount_paid,
			staged_payments.source_sheet,
			staged_payments.source_row,
			staged_payments.source_column,
			import_files.original_filename,
			import_files.file_date
		FROM staged_payments
		JOIN import_files ON import_files.id = staged_payments.file_id
		WHERE staged_payments.account_id IS NULL
		ORDER BY import_files.file_date DESC, import_files.original_filename DESC,
				 staged_payments.source_row ASC, staged_payments.source_column ASC
		""",
	)
	return {
		"counts": counts,
		"files": files,
		"unresolved": unresolved,
	}


def account_summary_for_month(month_value: str) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT
            accounts.id AS account_id,
            accounts.name AS account_name,
            SUM(payments.amount_paid) AS amount_paid
        FROM payments
        JOIN accounts ON accounts.id = payments.account_id
        WHERE strftime('%Y-%m', payments.payment_date) = ?
        GROUP BY accounts.id, accounts.name
        ORDER BY accounts.name
        """,
        (month_value,),
    )


def account_transactions_for_month(month_value: str, account_id: int) -> list[sqlite3.Row]:
    return query_all(
        """
        SELECT
            payments.payment_date,
            payments.supplier_name,
            payments.amount_paid,
            accounts.name AS account_name,
            payments.source_file_name,
            payments.source_sheet,
            payments.source_row,
            payments.source_column
        FROM payments
        JOIN accounts ON accounts.id = payments.account_id
        WHERE strftime('%Y-%m', payments.payment_date) = ?
          AND payments.account_id = ?
        ORDER BY payments.payment_date, payments.id
        """,
        (month_value, account_id),
    )


@app.before_request
def prepare_app() -> None:
	initialize_storage()
	if not getattr(app, "_seeded_accounts", False):
		sync_account_names()
		app._seeded_accounts = True


@app.teardown_appcontext
def close_connection(_: Exception | None = None) -> None:
	db = g.pop("db", None)
	if db is not None:
		db.close()


@app.route("/")
def index():
	account_count = query_one("SELECT COUNT(*) AS total FROM accounts")["total"]
	payment_count = query_one("SELECT COUNT(*) AS total FROM payments")["total"]
	batch_count = query_one("SELECT COUNT(*) AS total FROM import_batches")["total"]
	return render_template(
		"index.html",
		account_count=account_count,
		payment_count=payment_count,
		batch_count=batch_count,
	)


@app.route("/cash-payments", methods=["GET"])
def cash_payments():
	batch_id = request.args.get("batch_id", type=int) or latest_batch_id("cash")
	summary = batch_summary(batch_id) if batch_id else None
	accounts = query_all("SELECT id, name FROM accounts ORDER BY name")
	return render_template(
		"cash_payments.html",
		batch_id=batch_id,
		summary=summary,
		accounts=accounts,
	)


@app.route("/cash-payments/upload", methods=["POST"])
def upload_cash_payments():
    files = request.files.getlist("files")
    if not files:
        flash("Please choose one or more Excel files.", "warning")
        return redirect(url_for("cash_payments"))

    batch_id = execute(
        "INSERT INTO import_batches (source_folder, source_type, status) VALUES (?, ?, ?)",
        (request.form.get("source_folder", ""), "cash", "open"),
    ).lastrowid

    imported_files = 0
    staged_rows = 0
    duplicates = 0

    for file_storage in files:
        if not file_storage or not file_storage.filename:
            continue
        if not allowed_file(file_storage.filename):
            flash(f"Skipped {file_storage.filename}: only .xlsx and .xlsm files are supported.", "warning")
            continue
        try:
            file_date = parse_filename_date(Path(file_storage.filename).name)
        except ValueError as exc:
            flash(str(exc), "danger")
            continue

        file_id, staged_count, duplicate_flag = stage_workbook(
            batch_id=batch_id,
            file_storage=file_storage,
            original_filename=Path(file_storage.filename).name,
            file_date=file_date,
        )
        if duplicate_flag:
            duplicates += 1
            continue
        if file_id is not None:
            imported_files += 1
            staged_rows += staged_count

    if imported_files == 0 and staged_rows == 0:
        if duplicates:
            flash(f"All {duplicates} uploaded file(s) were already imported.", "info")
        else:
            flash("No new workbook rows were imported.", "info")
    else:
        message = f"Imported {imported_files} new file(s) and staged {staged_rows} payment row(s)."
        if duplicates:
            message += f" {duplicates} file(s) were already in the database."
        flash(message, "success")

    return redirect(url_for("cash_payments", batch_id=batch_id))


@app.route("/eft-payments", methods=["GET"])
def eft_payments():
	batch_id = request.args.get("batch_id", type=int) or latest_batch_id("eft")
	summary = batch_summary(batch_id) if batch_id else None
	accounts = query_all("SELECT id, name FROM accounts ORDER BY name")
	return render_template(
		"eft_payments.html",
		batch_id=batch_id,
		summary=summary,
		accounts=accounts,
	)


@app.route("/eft-payments/upload", methods=["POST"])
def upload_eft_payments():
	files = request.files.getlist("files")
	if not files:
		flash("Please choose one or more OFX files.", "warning")
		return redirect(url_for("eft_payments"))

	batch_id = execute(
		"INSERT INTO import_batches (source_folder, source_type, status) VALUES (?, ?, ?)",
		(request.form.get("source_folder", ""), "eft", "open"),
	).lastrowid

	imported_files = 0
	staged_rows = 0
	duplicates = 0

	for file_storage in files:
		if not file_storage or not file_storage.filename:
			continue
		if Path(file_storage.filename).suffix.lower() != ".ofx":
			flash(f"Skipped {file_storage.filename}: only .ofx files are supported.", "warning")
			continue
		file_id, staged_count, duplicate_flag = stage_ofx_file(
			batch_id=batch_id,
			file_storage=file_storage,
			original_filename=Path(file_storage.filename).name,
		)
		if duplicate_flag:
			duplicates += 1
			continue
		if file_id is not None:
			imported_files += 1
			staged_rows += staged_count

	if imported_files == 0 and staged_rows == 0:
		if duplicates:
			flash(f"All {duplicates} uploaded file(s) were already imported.", "info")
		else:
			flash("No new OFX transactions were imported.", "info")
	else:
		message = f"Imported {imported_files} new file(s) and staged {staged_rows} transaction(s)."
		if duplicates:
			message += f" {duplicates} file(s) were already in the database."
		flash(message, "success")

	return redirect(url_for("eft_payments", batch_id=batch_id))


def _resolve_payment_item(staged_id: int, redirect_endpoint: str):
	staged = query_one(
		"""
		SELECT staged_payments.*, import_files.original_filename, import_files.file_date
		FROM staged_payments
		JOIN import_files ON import_files.id = staged_payments.file_id
		WHERE staged_payments.id = ?
		""",
		(staged_id,),
	)
	if not staged:
		flash("That transaction could not be found.", "warning")
		return redirect(url_for("cash_payments"))

	if staged["account_id"] is not None:
		flash("That transaction is already resolved.", "info")
		return redirect(url_for(redirect_endpoint))

	selected_account_id = request.form.get("account_id", type=int)
	new_account_name = request.form.get("new_account_name", "").strip()
	account_id = None
	account_name = None

	if selected_account_id:
		account = query_one("SELECT id FROM accounts WHERE id = ?", (selected_account_id,))
		if account:
			account_id = selected_account_id
			account_name = query_one("SELECT name FROM accounts WHERE id = ?", (selected_account_id,))["name"]
	elif new_account_name:
		existing = find_account_by_normalized_name(new_account_name)
		if existing:
			account_id = existing["id"]
			account_name = existing["name"]
		else:
			upper_name = canonical_account_name(new_account_name)
			cursor = execute("INSERT INTO accounts (name) VALUES (?)", (upper_name,))
			account_id = cursor.lastrowid
			account_name = upper_name

	if not account_id:
		flash("Choose an existing account or type a new account name before saving.", "warning")
		return redirect(url_for(redirect_endpoint))

	ensure_rule(staged["supplier_name"], account_id)
	execute(
		"""
		UPDATE staged_payments
		SET account_id = ?
		WHERE batch_id = ? AND supplier_key = ? AND account_id IS NULL
		""",
		(account_id, staged["batch_id"], staged["supplier_key"]),
	)
	finalize_resolved_rows(staged["batch_id"])
	flash(f"Saved '{staged['supplier_name']}' to '{account_name}'.", "success")
	return redirect(url_for(redirect_endpoint))


@app.route("/cash-payments/resolve/<int:staged_id>", methods=["POST"])
def resolve_cash_payment_item(staged_id: int):
	return _resolve_payment_item(staged_id, "cash_payments")


@app.route("/eft-payments/resolve/<int:staged_id>", methods=["POST"])
def resolve_eft_payment_item(staged_id: int):
	return _resolve_payment_item(staged_id, "eft_payments")


@app.route("/accounts", methods=["GET", "POST"])
def accounts():
    month_value = normalize_month(request.args.get("month")) or latest_payment_month()
    month_options = available_payment_months()
    if not month_value and month_options:
        month_value = month_options[0]["month_value"]
    rows = account_summary_for_month(month_value) if month_value else []
    return render_template(
        "accounts.html",
        month_value=month_value,
        month_label=month_label(month_value) if month_value else None,
        month_options=month_options,
        rows=rows,
    )


@app.route("/accounts/<int:account_id>", methods=["GET"])
def account_transactions(account_id: int):
    account = query_one("SELECT id, name FROM accounts WHERE id = ?", (account_id,))
    if not account:
        flash("That account could not be found.", "warning")
        return redirect(url_for("accounts"))

    month_value = normalize_month(request.args.get("month")) or latest_payment_month()
    month_options = available_payment_months()
    if not month_value and month_options:
        month_value = month_options[0]["month_value"]
    transactions = account_transactions_for_month(month_value, account_id) if month_value else []
    total_amount = sum(row["amount_paid"] for row in transactions)
    return render_template(
        "account_transactions.html",
        account_name=account["name"],
        account_id=account_id,
        month_value=month_value,
        month_label=month_label(month_value) if month_value else None,
        month_options=month_options,
        transactions=transactions,
        total_amount=total_amount,
    )


@app.route("/settings", methods=["GET", "POST"])
def settings():
    if request.method == "POST":
        account_name = request.form.get("account_name", "").strip()
        if account_name:
            existing = find_account_by_normalized_name(account_name)
            if existing:
                flash(f"Account '{account_name}' already exists.", "info")
            else:
                upper_name = canonical_account_name(account_name)
                execute("INSERT INTO accounts (name) VALUES (?)", (upper_name,))
                flash(f"Created account '{upper_name}'.", "success")
        return redirect(url_for("settings"))

    account_rows = query_all(
        """
        SELECT accounts.id, accounts.name, accounts.source_sheet,
               COUNT(DISTINCT supplier_rules.id) AS linked_suppliers
        FROM accounts
        LEFT JOIN supplier_rules ON supplier_rules.account_id = accounts.id
        GROUP BY accounts.id
        ORDER BY accounts.name
        """
    )
    rules = query_all(
        """
        SELECT supplier_rules.supplier_name, accounts.name AS account_name
        FROM supplier_rules
        JOIN accounts ON accounts.id = supplier_rules.account_id
        ORDER BY supplier_rules.supplier_name
        """
    )
    return render_template("settings.html", accounts=account_rows, rules=rules)


if __name__ == "__main__":
    with app.app_context():
        initialize_storage()
        sync_account_names()
    app.run(debug=True)
