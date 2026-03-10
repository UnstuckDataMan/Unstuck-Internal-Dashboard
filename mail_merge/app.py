import os
import uuid
import json
import requests as http_req
from flask import Flask, request, jsonify, send_file
from werkzeug.utils import secure_filename

from utils.excel_reader import parse_prospect_file
from utils.merge import validate_templates, perform_merge
from utils.scheduler import generate_schedule, get_working_days
from utils.excel_writer import write_merge_output

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(os.path.dirname(__file__), 'outputs')
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32 MB

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# ── Supabase config (set these as environment variables) ─────────────────────
SUPABASE_URL     = os.environ.get('SUPABASE_URL', '').rstrip('/')
SUPABASE_ANON_KEY = os.environ.get('SUPABASE_ANON_KEY', '')

def _sb_headers():
    return {
        'apikey':        SUPABASE_ANON_KEY,
        'Authorization': f'Bearer {SUPABASE_ANON_KEY}',
        'Content-Type':  'application/json',
    }


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    debug_path = os.path.join(BASE_DIR, 'debug_index.txt')
    try:
        with open(debug_path, 'w') as dbg:
            dbg.write('index() called\n')
        html_path = os.path.join(BASE_DIR, 'templates', 'index.html')
        with open(html_path, 'rb') as f:
            content = f.read()
        with open(debug_path, 'a') as dbg:
            dbg.write(f'read {len(content)} bytes\n')
        from flask import Response
        return Response(content, status=200, mimetype='text/html; charset=utf-8')
    except Exception as exc:
        import traceback
        tb = traceback.format_exc()
        with open(debug_path, 'a') as dbg:
            dbg.write('EXCEPTION:\n' + tb)
        return f'<pre>{tb}</pre>', 500


@app.route('/api/upload-prospects', methods=['POST'])
def upload_prospects():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    file = request.files['file']
    if not file or file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    if not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Please upload .xlsx or .xls'}), 400

    filename = f"{uuid.uuid4().hex}_{secure_filename(file.filename)}"
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    try:
        headers, all_rows, total_rows = parse_prospect_file(filepath)
        preview = all_rows[:5]
        # Convert to serialisable format
        preview_list = [dict(row) for row in preview]
        return jsonify({
            'file_id': filename,
            'headers': headers,
            'preview': preview_list,
            'total_rows': total_rows
        })
    except Exception as e:
        return jsonify({'error': f'Failed to parse file: {str(e)}'}), 400


@app.route('/api/validate-templates', methods=['POST'])
def validate_templates_route():
    data = request.json
    headers = data.get('headers', [])
    templates = (
        data.get('subject_templates', []) +
        data.get('body_templates', []) +
        ([data['chaser_subject']] if data.get('chaser_subject') else []) +
        ([data['chaser_body']] if data.get('chaser_body') else [])
    )
    errors = validate_templates(templates, headers)
    return jsonify({'valid': len(errors) == 0, 'errors': errors})


@app.route('/api/schedule-capacity', methods=['GET'])
def schedule_capacity():
    """Return working-day count and max schedulable prospect count for a given
    month, sender count, and daily-limit setting.  Used by the frontend to
    render the live capacity banner before the user hits Generate."""
    year         = request.args.get('year',         type=int)
    month        = request.args.get('month',        type=int)
    sender_count  = request.args.get('sender_count',  type=int, default=1)
    sends_per_day = request.args.get('sends_per_day', type=int, default=10)

    if not year or not month:
        return jsonify({'error': 'year and month are required'}), 400

    sender_count  = max(1, sender_count  or 1)
    sends_per_day = sends_per_day if sends_per_day in (5, 10, 15) else 10

    try:
        working_days = get_working_days(year, month)
        n_days       = len(working_days)
        max_capacity = sends_per_day * sender_count * n_days
        return jsonify({
            'working_days': n_days,
            'max_capacity': max_capacity,
            'daily_sends':  sends_per_day * sender_count,
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/generate-merge', methods=['POST'])
def generate_merge():
    data = request.json
    file_id = data.get('file_id')
    subject_templates = data.get('subject_templates', [])
    body_templates = data.get('body_templates', [])
    chaser_subject = data.get('chaser_subject', '')
    chaser_body = data.get('chaser_body', '')
    sender_emails = data.get('sender_emails', [])
    missing_value = data.get('missing_value', '[MISSING]')
    email_column = data.get('email_column', '')
    year = data.get('year')
    month = data.get('month')
    recipient_tz      = data.get('recipient_tz', 'Europe/London')
    sender_tz         = data.get('sender_tz', 'Europe/London')
    sends_per_day_raw = data.get('sends_per_day', 10)
    sends_per_day     = sends_per_day_raw if sends_per_day_raw in (5, 10, 15) else 10

    if not file_id:
        return jsonify({'error': 'No prospect file specified'}), 400
    if not subject_templates:
        return jsonify({'error': 'At least one subject line template is required'}), 400
    if not body_templates:
        return jsonify({'error': 'At least one email body template is required'}), 400
    if not sender_emails:
        return jsonify({'error': 'At least one sender email address is required'}), 400
    if not year or not month:
        return jsonify({'error': 'Schedule month and year are required'}), 400

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file_id)
    if not os.path.exists(filepath):
        return jsonify({'error': 'Prospect file not found. Please re-upload.'}), 404

    try:
        year = int(year)
        month = int(month)

        headers, all_rows, _ = parse_prospect_file(filepath)

        all_templates = subject_templates + body_templates
        if chaser_subject:
            all_templates.append(chaser_subject)
        if chaser_body:
            all_templates.append(chaser_body)

        errors = validate_templates(all_templates, headers)
        if errors:
            return jsonify({'error': 'Template validation failed', 'details': errors}), 400

        merged_rows = perform_merge(
            all_rows, headers, subject_templates, body_templates,
            chaser_subject, chaser_body, sender_emails, missing_value,
            email_column
        )

        # ── Capacity-aware scheduling ──────────────────────────────────────
        total_prospects = len(merged_rows)

        # How many can fit at the chosen sends-per-day rate?
        working_days_list = get_working_days(year, month)
        max_capacity    = sends_per_day * len(sender_emails) * len(working_days_list)
        scheduled_count = min(total_prospects, max_capacity)
        overflow_count  = total_prospects - scheduled_count

        if scheduled_count < 1:
            return jsonify({
                'error': 'No prospects can be scheduled. '
                         'Check your month selection, sender count, and daily limit.'
            }), 400

        # Generate schedule for the schedulable subset and join onto rows
        schedule = generate_schedule(
            year, month, scheduled_count, sender_emails,
            recipient_tz=recipient_tz, sender_tz=sender_tz,
            max_per_sender_per_day=sends_per_day,
        )
        for entry in schedule:
            row = merged_rows[entry['prospect_id'] - 1]
            row['__send_date__'] = entry['date']
            row['__send_time__'] = entry['send_time']
            row['__sender_number__'] = entry['sender_number']

        # Sort chronologically by date → sender profile order → send time,
        # so the sheet reads: Sender 1 today, Sender 2 today, Sender 1 tomorrow …
        sender_order = {email: idx for idx, email in enumerate(sender_emails)}
        merged_rows.sort(key=lambda r: (
            r.get('__send_date__', ''),
            sender_order.get(r.get('__sender_account__', ''), 999),
            r.get('__send_time__', ''),
        ))
        merged_rows = [r for r in merged_rows if r.get('__send_date__')]

        has_chaser = bool(chaser_subject or chaser_body)
        output_filename = f"outreach_merge_{uuid.uuid4().hex[:8]}.xlsx"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        write_merge_output(output_path, headers, merged_rows, has_chaser, email_column,
                           has_schedule=True)

        return jsonify({
            'download_id': output_filename,
            'total_rows': len(merged_rows),        # rows in the Excel (= scheduled_count)
            'scheduled_count': len(merged_rows),
            'overflow_count': overflow_count,
            'preview': [
                {
                    'sender': r.get('__sender_account__', ''),
                    'recipient': r.get('__recipient_email__', ''),
                    'subject': r.get('__subject_line__', ''),
                    'body_preview': r.get('__email_body__', '')[:120] + '...' if len(r.get('__email_body__', '')) > 120 else r.get('__email_body__', ''),
                    'variant': r.get('__template_variant__', ''),
                    'send_date': r.get('__send_date__', ''),
                    'send_time': r.get('__send_time__', ''),
                }
                for r in merged_rows[:5]
            ]
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/sender-profiles', methods=['GET'])
def get_sender_profiles():
    if not SUPABASE_URL or not SUPABASE_ANON_KEY:
        return jsonify({'error': 'Supabase is not configured on this server.'}), 503
    try:
        r = http_req.get(
            f'{SUPABASE_URL}/rest/v1/sender_profiles',
            headers=_sb_headers(),
            params={'select': 'profile_name,emails', 'order': 'profile_name.asc'},
            timeout=10,
        )
        r.raise_for_status()
        return jsonify(r.json())
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/sender-profiles', methods=['POST'])
def save_sender_profile():
    if not SUPABASE_URL or not SUPABASE_ANON_KEY:
        return jsonify({'error': 'Supabase is not configured on this server.'}), 503
    data = request.json or {}
    profile_name = data.get('profile_name', '').strip()
    emails = data.get('emails', [])
    if not profile_name:
        return jsonify({'error': 'Profile name is required.'}), 400
    if not emails:
        return jsonify({'error': 'At least one email address is required.'}), 400
    try:
        headers = {**_sb_headers(), 'Prefer': 'resolution=merge-duplicates'}
        r = http_req.post(
            f'{SUPABASE_URL}/rest/v1/sender_profiles',
            headers=headers,
            params={'on_conflict': 'profile_name'},
            json={'profile_name': profile_name, 'emails': emails},
            timeout=10,
        )
        r.raise_for_status()
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/sender-profiles/<path:profile_name>', methods=['DELETE'])
def delete_sender_profile(profile_name):
    if not SUPABASE_URL or not SUPABASE_ANON_KEY:
        return jsonify({'error': 'Supabase is not configured on this server.'}), 503
    try:
        r = http_req.delete(
            f'{SUPABASE_URL}/rest/v1/sender_profiles',
            headers=_sb_headers(),
            params={'profile_name': f'eq.{profile_name}'},
            timeout=10,
        )
        r.raise_for_status()
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/download/<filename>')
def download_file(filename):
    if not filename.endswith('.xlsx') or '/' in filename or '\\' in filename or '..' in filename:
        return jsonify({'error': 'Invalid filename'}), 400

    filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    if not os.path.exists(filepath):
        return jsonify({'error': 'File not found'}), 404

    custom_name = request.args.get('name', '').strip()
    if custom_name:
        download_name = secure_filename(custom_name)
        if not download_name.lower().endswith('.xlsx'):
            download_name += '.xlsx'
    else:
        download_name = filename

    return send_file(filepath, as_attachment=True, download_name=download_name)


if __name__ == '__main__':
    print("\n  Mail Merge Tool")
    print("  ---------------------------------")
    print("  Open your browser and go to:")
    print("  http://localhost:5000\n")
    app.run(debug=True, host='0.0.0.0', port=5000, use_reloader=False)
