import io
import os
import sqlite3
from datetime import datetime
from functools import wraps

import pandas as pd
from flask import (Flask, abort, flash, redirect, render_template, request,
                   send_from_directory, session, url_for)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'change-this-in-prod-a8f3k2p9')

# ── Storage paths ────────────────────────────────────────────────────────────
DATA_DIR = os.environ.get('DATA_DIR', os.path.join(os.path.dirname(__file__), 'data'))
DB_PATH = os.path.join(DATA_DIR, 'procurement.db')
UPLOAD_FOLDER = os.path.join(DATA_DIR, 'uploads')
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ALLOWED_EXT = {'xlsx', 'xls', 'csv'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT


# ── Database ─────────────────────────────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    with get_db() as conn:
        conn.executescript('''
            CREATE TABLE IF NOT EXISTS users (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                username      TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS items (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                name        TEXT NOT NULL,
                category    TEXT,
                description TEXT,
                qty         REAL,
                unit        TEXT,
                created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS suppliers (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                name       TEXT NOT NULL,
                phone      TEXT,
                address    TEXT,
                remarks    TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS quote_requests (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                title      TEXT NOT NULL,
                notes      TEXT,
                status     TEXT DEFAULT 'open',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS quote_request_items (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                quote_request_id INTEGER NOT NULL,
                item_id          INTEGER NOT NULL,
                FOREIGN KEY (quote_request_id) REFERENCES quote_requests(id) ON DELETE CASCADE,
                FOREIGN KEY (item_id)          REFERENCES items(id)          ON DELETE CASCADE,
                UNIQUE(quote_request_id, item_id)
            );

            CREATE TABLE IF NOT EXISTS quote_request_suppliers (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                quote_request_id INTEGER NOT NULL,
                supplier_id      INTEGER NOT NULL,
                FOREIGN KEY (quote_request_id) REFERENCES quote_requests(id) ON DELETE CASCADE,
                FOREIGN KEY (supplier_id)      REFERENCES suppliers(id)      ON DELETE CASCADE,
                UNIQUE(quote_request_id, supplier_id)
            );

            CREATE TABLE IF NOT EXISTS supplier_quotes (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                quote_request_id INTEGER NOT NULL,
                supplier_id      INTEGER NOT NULL,
                file_path        TEXT,
                notes            TEXT,
                uploaded_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (quote_request_id) REFERENCES quote_requests(id) ON DELETE CASCADE,
                FOREIGN KEY (supplier_id)      REFERENCES suppliers(id)      ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS quote_prices (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                supplier_quote_id INTEGER NOT NULL,
                item_id          INTEGER NOT NULL,
                price            REAL,
                notes            TEXT,
                FOREIGN KEY (supplier_quote_id) REFERENCES supplier_quotes(id) ON DELETE CASCADE,
                FOREIGN KEY (item_id)           REFERENCES items(id)           ON DELETE CASCADE
            );
        ''')

        # Seed default admin account
        admin_username = os.environ.get('ADMIN_USERNAME', 'admin')
        admin_password = os.environ.get('ADMIN_PASSWORD', 'admin123')
        row = conn.execute('SELECT id FROM users WHERE username = ?', (admin_username,)).fetchone()
        if not row:
            conn.execute(
                'INSERT INTO users (username, password_hash) VALUES (?, ?)',
                (admin_username, generate_password_hash(admin_password))
            )


# ── Auth helpers ──────────────────────────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login', next=request.url))
        return f(*args, **kwargs)
    return decorated


# ── Auth routes ───────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return redirect(url_for('dashboard') if 'user_id' in session else url_for('login'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        with get_db() as conn:
            user = conn.execute('SELECT * FROM users WHERE username = ?', (username,)).fetchone()
        if user and check_password_hash(user['password_hash'], password):
            session['user_id'] = user['id']
            session['username'] = user['username']
            return redirect(request.args.get('next') or url_for('dashboard'))
        flash('Invalid username or password.', 'danger')
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


# ── Dashboard ─────────────────────────────────────────────────────────────────
@app.route('/dashboard')
@login_required
def dashboard():
    with get_db() as conn:
        item_count     = conn.execute('SELECT COUNT(*) FROM items').fetchone()[0]
        supplier_count = conn.execute('SELECT COUNT(*) FROM suppliers').fetchone()[0]
        qr_count       = conn.execute('SELECT COUNT(*) FROM quote_requests').fetchone()[0]
        open_count     = conn.execute("SELECT COUNT(*) FROM quote_requests WHERE status='open'").fetchone()[0]
        recent = conn.execute('''
            SELECT sq.uploaded_at, s.name AS supplier_name, qr.title, qr.id AS qr_id
            FROM   supplier_quotes sq
            JOIN   suppliers      s  ON s.id  = sq.supplier_id
            JOIN   quote_requests qr ON qr.id = sq.quote_request_id
            ORDER  BY sq.uploaded_at DESC LIMIT 6
        ''').fetchall()
    return render_template('dashboard.html',
                           item_count=item_count, supplier_count=supplier_count,
                           qr_count=qr_count, open_count=open_count, recent=recent)


# ── Items ─────────────────────────────────────────────────────────────────────
@app.route('/items')
@login_required
def items_list():
    with get_db() as conn:
        items = conn.execute('SELECT * FROM items ORDER BY category, name').fetchall()
    return render_template('items/index.html', items=items)


@app.route('/items/import', methods=['GET', 'POST'])
@login_required
def items_import():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file or file.filename == '':
            flash('No file selected.', 'danger')
            return redirect(request.url)
        if not allowed_file(file.filename):
            flash('Unsupported file type. Use .xlsx, .xls, or .csv', 'danger')
            return redirect(request.url)
        try:
            raw = file.read()
            buf = io.BytesIO(raw)
            df = pd.read_csv(buf) if file.filename.lower().endswith('.csv') else pd.read_excel(buf)
            df.columns = [str(c).lower().strip().replace(' ', '_') for c in df.columns]

            if 'name' not in df.columns:
                flash('File must contain a "name" column.', 'danger')
                return redirect(request.url)

            mode = request.form.get('mode', 'append')
            with get_db() as conn:
                if mode == 'replace':
                    conn.execute('DELETE FROM items')
                count = 0
                for _, row in df.iterrows():
                    name = str(row.get('name', '')).strip()
                    if not name or name.lower() == 'nan':
                        continue
                    def clean(col):
                        v = str(row.get(col, '') or '').strip()
                        return None if v in ('', 'nan') else v
                    qty_val = row.get('qty', row.get('quantity', None))
                    try:
                        qty = float(qty_val) if qty_val is not None and str(qty_val).lower() != 'nan' else None
                    except (ValueError, TypeError):
                        qty = None
                    conn.execute(
                        'INSERT INTO items (name, category, description, qty, unit) VALUES (?,?,?,?,?)',
                        (name, clean('category'), clean('description'), qty, clean('unit'))
                    )
                    count += 1
            flash(f'Imported {count} items successfully.', 'success')
            return redirect(url_for('items_list'))
        except Exception as e:
            flash(f'Error processing file: {e}', 'danger')
            return redirect(request.url)
    return render_template('items/import.html')


@app.route('/items/edit/<int:item_id>', methods=['GET', 'POST'])
@login_required
def items_edit(item_id):
    with get_db() as conn:
        item = conn.execute('SELECT * FROM items WHERE id = ?', (item_id,)).fetchone()
        if not item:
            abort(404)
        if request.method == 'POST':
            conn.execute(
                'UPDATE items SET name=?,category=?,description=?,qty=?,unit=? WHERE id=?',
                (
                    request.form['name'].strip(),
                    request.form.get('category', '').strip() or None,
                    request.form.get('description', '').strip() or None,
                    request.form.get('qty') or None,
                    request.form.get('unit', '').strip() or None,
                    item_id,
                )
            )
            flash('Item updated.', 'success')
            return redirect(url_for('items_list'))
    return render_template('items/edit.html', item=item)


@app.route('/items/delete/<int:item_id>', methods=['POST'])
@login_required
def items_delete(item_id):
    with get_db() as conn:
        conn.execute('DELETE FROM items WHERE id = ?', (item_id,))
    flash('Item deleted.', 'success')
    return redirect(url_for('items_list'))


# ── Suppliers ─────────────────────────────────────────────────────────────────
@app.route('/suppliers')
@login_required
def suppliers_list():
    with get_db() as conn:
        suppliers = conn.execute('SELECT * FROM suppliers ORDER BY name').fetchall()
    return render_template('suppliers/index.html', suppliers=suppliers)


@app.route('/suppliers/add', methods=['GET', 'POST'])
@login_required
def suppliers_add():
    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        if not name:
            flash('Supplier name is required.', 'danger')
            return redirect(request.url)
        with get_db() as conn:
            conn.execute(
                'INSERT INTO suppliers (name, phone, address, remarks) VALUES (?,?,?,?)',
                (name,
                 request.form.get('phone', '').strip() or None,
                 request.form.get('address', '').strip() or None,
                 request.form.get('remarks', '').strip() or None)
            )
        flash('Supplier added.', 'success')
        return redirect(url_for('suppliers_list'))
    return render_template('suppliers/form.html', supplier=None, action='Add')


@app.route('/suppliers/edit/<int:supplier_id>', methods=['GET', 'POST'])
@login_required
def suppliers_edit(supplier_id):
    with get_db() as conn:
        supplier = conn.execute('SELECT * FROM suppliers WHERE id = ?', (supplier_id,)).fetchone()
        if not supplier:
            abort(404)
        if request.method == 'POST':
            name = request.form.get('name', '').strip()
            if not name:
                flash('Supplier name is required.', 'danger')
                return redirect(request.url)
            conn.execute(
                'UPDATE suppliers SET name=?,phone=?,address=?,remarks=? WHERE id=?',
                (name,
                 request.form.get('phone', '').strip() or None,
                 request.form.get('address', '').strip() or None,
                 request.form.get('remarks', '').strip() or None,
                 supplier_id)
            )
            flash('Supplier updated.', 'success')
            return redirect(url_for('suppliers_list'))
    return render_template('suppliers/form.html', supplier=supplier, action='Edit')


@app.route('/suppliers/delete/<int:supplier_id>', methods=['POST'])
@login_required
def suppliers_delete(supplier_id):
    with get_db() as conn:
        conn.execute('DELETE FROM suppliers WHERE id = ?', (supplier_id,))
    flash('Supplier deleted.', 'success')
    return redirect(url_for('suppliers_list'))


# ── Quote Requests ────────────────────────────────────────────────────────────
@app.route('/quotes')
@login_required
def quotes_list():
    with get_db() as conn:
        rows = conn.execute('''
            SELECT qr.*,
                   COUNT(DISTINCT qri.item_id)     AS item_count,
                   COUNT(DISTINCT qrs.supplier_id) AS supplier_count,
                   COUNT(DISTINCT sq.id)           AS received_count
            FROM   quote_requests qr
            LEFT JOIN quote_request_items     qri ON qri.quote_request_id = qr.id
            LEFT JOIN quote_request_suppliers qrs ON qrs.quote_request_id = qr.id
            LEFT JOIN supplier_quotes         sq  ON sq.quote_request_id  = qr.id
            GROUP BY qr.id
            ORDER BY qr.created_at DESC
        ''').fetchall()
    return render_template('quotes/index.html', requests=rows)


@app.route('/quotes/create', methods=['GET', 'POST'])
@login_required
def quotes_create():
    with get_db() as conn:
        items     = conn.execute('SELECT * FROM items ORDER BY category, name').fetchall()
        suppliers = conn.execute('SELECT * FROM suppliers ORDER BY name').fetchall()

        if request.method == 'POST':
            title       = request.form.get('title', '').strip()
            notes       = request.form.get('notes', '').strip()
            item_ids    = request.form.getlist('item_ids')
            supplier_ids = request.form.getlist('supplier_ids')

            errors = []
            if not title:       errors.append('Title is required.')
            if not item_ids:    errors.append('Select at least one item.')
            if not supplier_ids: errors.append('Select at least one supplier.')
            for e in errors:
                flash(e, 'danger')
            if errors:
                return render_template('quotes/create.html', items=items, suppliers=suppliers)

            cur = conn.execute('INSERT INTO quote_requests (title, notes) VALUES (?,?)', (title, notes or None))
            qr_id = cur.lastrowid
            for iid in item_ids:
                conn.execute('INSERT INTO quote_request_items (quote_request_id, item_id) VALUES (?,?)', (qr_id, int(iid)))
            for sid in supplier_ids:
                conn.execute('INSERT INTO quote_request_suppliers (quote_request_id, supplier_id) VALUES (?,?)', (qr_id, int(sid)))

            flash('Quote request created.', 'success')
            return redirect(url_for('quotes_detail', qr_id=qr_id))

    return render_template('quotes/create.html', items=items, suppliers=suppliers)


@app.route('/quotes/<int:qr_id>')
@login_required
def quotes_detail(qr_id):
    with get_db() as conn:
        qr = conn.execute('SELECT * FROM quote_requests WHERE id = ?', (qr_id,)).fetchone()
        if not qr:
            abort(404)
        items = conn.execute('''
            SELECT i.* FROM items i
            JOIN quote_request_items qri ON qri.item_id = i.id
            WHERE qri.quote_request_id = ?
            ORDER BY i.category, i.name
        ''', (qr_id,)).fetchall()
        suppliers = conn.execute('''
            SELECT s.* FROM suppliers s
            JOIN quote_request_suppliers qrs ON qrs.supplier_id = s.id
            WHERE qrs.quote_request_id = ?
            ORDER BY s.name
        ''', (qr_id,)).fetchall()
        sq_map = {}
        for s in suppliers:
            sq = conn.execute('''
                SELECT * FROM supplier_quotes
                WHERE quote_request_id = ? AND supplier_id = ?
                ORDER BY uploaded_at DESC LIMIT 1
            ''', (qr_id, s['id'])).fetchone()
            sq_map[s['id']] = sq
    return render_template('quotes/detail.html', qr=qr, items=items, suppliers=suppliers, sq_map=sq_map)


@app.route('/quotes/<int:qr_id>/upload/<int:supplier_id>', methods=['GET', 'POST'])
@login_required
def quotes_upload(qr_id, supplier_id):
    with get_db() as conn:
        qr       = conn.execute('SELECT * FROM quote_requests WHERE id = ?', (qr_id,)).fetchone()
        supplier = conn.execute('SELECT * FROM suppliers WHERE id = ?', (supplier_id,)).fetchone()
        if not qr or not supplier:
            abort(404)
        items = conn.execute('''
            SELECT i.* FROM items i
            JOIN quote_request_items qri ON qri.item_id = i.id
            WHERE qri.quote_request_id = ?
            ORDER BY i.category, i.name
        ''', (qr_id,)).fetchall()

        if request.method == 'POST':
            notes      = request.form.get('notes', '').strip()
            file_path  = None
            prices_map = {}   # item_id -> price

            # ── Parse uploaded file ──────────────────────────────────────────
            file = request.files.get('file')
            if file and file.filename:
                if not allowed_file(file.filename):
                    flash('Unsupported file type. Use .xlsx, .xls, or .csv', 'danger')
                    return redirect(request.url)
                try:
                    raw = file.read()
                    buf = io.BytesIO(raw)
                    df  = pd.read_csv(buf) if file.filename.lower().endswith('.csv') else pd.read_excel(buf)
                    df.columns = [str(c).lower().strip().replace(' ', '_') for c in df.columns]

                    # Save original file
                    ts = datetime.now().strftime('%Y%m%d%H%M%S')
                    fname = secure_filename(f"q{qr_id}_s{supplier_id}_{ts}_{file.filename}")
                    with open(os.path.join(UPLOAD_FOLDER, fname), 'wb') as fh:
                        fh.write(raw)
                    file_path = fname

                    # Build name → id map (case-insensitive)
                    name_to_id = {i['name'].lower().strip(): i['id'] for i in items}

                    # Detect name and price columns flexibly
                    name_col  = next((c for c in ['item_name', 'name', 'item', 'description'] if c in df.columns), None)
                    price_col = next((c for c in ['price', 'unit_price', 'quoted_price', 'amount', 'rate'] if c in df.columns), None)

                    if name_col and price_col:
                        for _, row in df.iterrows():
                            key = str(row.get(name_col, '')).lower().strip()
                            if key in name_to_id and row.get(price_col) is not None:
                                try:
                                    prices_map[name_to_id[key]] = float(str(row[price_col]).replace(',', ''))
                                except (ValueError, TypeError):
                                    pass
                except Exception as e:
                    flash(f'Error reading file: {e}', 'danger')
                    return redirect(request.url)

            # ── Manual price overrides from form ─────────────────────────────
            for item in items:
                key = f'price_{item["id"]}'
                val = request.form.get(key, '').strip()
                if val:
                    try:
                        prices_map[item['id']] = float(val.replace(',', ''))
                    except ValueError:
                        pass

            if not prices_map:
                flash('No prices found. Please check your file columns or enter prices manually.', 'warning')
                return redirect(request.url)

            # Save quote + prices
            cur = conn.execute(
                'INSERT INTO supplier_quotes (quote_request_id, supplier_id, file_path, notes) VALUES (?,?,?,?)',
                (qr_id, supplier_id, file_path, notes or None)
            )
            sq_id = cur.lastrowid
            for item_id, price in prices_map.items():
                conn.execute(
                    'INSERT INTO quote_prices (supplier_quote_id, item_id, price) VALUES (?,?,?)',
                    (sq_id, item_id, price)
                )
            flash(f'Quote from {supplier["name"]} saved with {len(prices_map)} prices.', 'success')
            return redirect(url_for('quotes_detail', qr_id=qr_id))

    return render_template('quotes/upload.html', qr=qr, supplier=supplier, items=items)


@app.route('/quotes/<int:qr_id>/comparison')
@login_required
def quotes_comparison(qr_id):
    with get_db() as conn:
        qr = conn.execute('SELECT * FROM quote_requests WHERE id = ?', (qr_id,)).fetchone()
        if not qr:
            abort(404)
        items = conn.execute('''
            SELECT i.* FROM items i
            JOIN quote_request_items qri ON qri.item_id = i.id
            WHERE qri.quote_request_id = ?
            ORDER BY i.category, i.name
        ''', (qr_id,)).fetchall()
        suppliers = conn.execute('''
            SELECT s.* FROM suppliers s
            JOIN quote_request_suppliers qrs ON qrs.supplier_id = s.id
            WHERE qrs.quote_request_id = ?
            ORDER BY s.name
        ''', (qr_id,)).fetchall()

        # price_matrix[item_id][supplier_id] = price
        price_matrix = {item['id']: {} for item in items}
        for s in suppliers:
            sq = conn.execute('''
                SELECT id FROM supplier_quotes
                WHERE quote_request_id = ? AND supplier_id = ?
                ORDER BY uploaded_at DESC LIMIT 1
            ''', (qr_id, s['id'])).fetchone()
            if sq:
                for p in conn.execute('SELECT item_id, price FROM quote_prices WHERE supplier_quote_id = ?', (sq['id'],)).fetchall():
                    if p['item_id'] in price_matrix:
                        price_matrix[p['item_id']][s['id']] = p['price']

    return render_template('quotes/comparison.html',
                           qr=qr, items=items, suppliers=suppliers, price_matrix=price_matrix)


@app.route('/quotes/<int:qr_id>/toggle-status', methods=['POST'])
@login_required
def quotes_toggle_status(qr_id):
    with get_db() as conn:
        qr = conn.execute('SELECT status FROM quote_requests WHERE id = ?', (qr_id,)).fetchone()
        if qr:
            new = 'closed' if qr['status'] == 'open' else 'open'
            conn.execute('UPDATE quote_requests SET status = ? WHERE id = ?', (new, qr_id))
    return redirect(url_for('quotes_detail', qr_id=qr_id))


@app.route('/quotes/<int:qr_id>/delete', methods=['POST'])
@login_required
def quotes_delete(qr_id):
    with get_db() as conn:
        conn.execute('DELETE FROM quote_requests WHERE id = ?', (qr_id,))
    flash('Quote request deleted.', 'success')
    return redirect(url_for('quotes_list'))


# ── Uploaded file download ────────────────────────────────────────────────────
@app.route('/uploads/<path:filename>')
@login_required
def uploaded_file(filename):
    return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=True)


# ── Bootstrap ─────────────────────────────────────────────────────────────────
init_db()

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
