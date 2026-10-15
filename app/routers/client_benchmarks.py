from fastapi import APIRouter, Request, UploadFile, File, Form, Depends, HTTPException
from starlette.concurrency import run_in_threadpool
from app.utils.benchmark_upload import read_excel, MAX_UPLOAD_BYTES
from fastapi.responses import JSONResponse

from app.deps import templates
from app.utils import client_benchmarks as benchmarks

def require_available():
    """Temporary shutdown, scoped to this tool's page and all API routes."""
    raise HTTPException(status_code=503, detail='Average Client Data is temporarily offline.',
                        headers={'Cache-Control': 'no-store'})


router = APIRouter(dependencies=[Depends(require_available)])


@router.get('/client-benchmarks')
def page(request: Request):
    start, end = benchmarks.default_period()
    return templates.TemplateResponse(request, 'client_benchmarks.html', {
        'start_month': start, 'end_month': end, 'source_url': benchmarks.SOURCE_URL,
    })


@router.get('/api/client-benchmarks')
def data(start: str = '', end: str = ''):
    default_start, default_end = benchmarks.default_period()
    try:
        months = benchmarks.period_months(start or default_start, end or default_end) if start or end else None
    except ValueError as exc:
        return JSONResponse({'error': str(exc)}, status_code=400)
    try:
        source = benchmarks.load_source()
        if months is None:
            today = benchmarks.london_today()
            records = [row for row in source['records'] if row['date'] <= today]
            source = {**source, 'records': records}
            first = min((row['date'] for row in records), default=today)
            months = [benchmarks.month_string(i) for i in range(
                first.year * 12 + first.month - 1, today.year * 12 + today.month)]
            result = benchmarks.summarize(source, months, lead_producing_only=True)
            result.update(history_scope='all_available', through_date=today.isoformat(), current_month_partial=True)
        else:
            result = benchmarks.summarize(source, months, lead_producing_only=True)
    except benchmarks.SourceError as exc:
        return JSONResponse({'error': str(exc)}, status_code=503, headers={'Cache-Control': 'no-store'})
    return JSONResponse(result, headers={'Cache-Control': 'no-store'})


@router.post('/api/client-benchmarks/upload')
async def upload(file: UploadFile = File(...), start: str = Form(''), end: str = Form('')):
    try:
        default_start, default_end = benchmarks.default_period()
        months = benchmarks.period_months(start or default_start, end or default_end)
        if not (file.filename or '').lower().endswith('.xlsx'):
            raise benchmarks.SourceError('Choose an Excel .xlsx file.')
        content = await file.read(MAX_UPLOAD_BYTES + 1)
        source = await run_in_threadpool(read_excel, content)
        return JSONResponse(benchmarks.summarize(source, months), headers={'Cache-Control': 'no-store'})
    except (ValueError, benchmarks.SourceError) as exc:
        return JSONResponse({'error': str(exc)}, status_code=400, headers={'Cache-Control': 'no-store'})
    finally:
        await file.close()


@router.get('/api/client-benchmarks/activity-template')
def activity_template():
    from io import BytesIO
    import openpyxl
    from fastapi.responses import Response
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'Campaign activity'
    sheet.append(['Campaign ID', 'Client', 'Month', 'Company size', 'Active', 'Leads'])
    sheet.freeze_panes = 'A2'
    for column in 'ABCDEF':
        sheet.column_dimensions[column].width = 24
    notes = workbook.create_sheet('Instructions')
    for text in [
        'Supply real campaign activity; do not infer activity from whether a lead was recorded.',
        'One row per campaign, month and targeted size band. Include every active month, even with zero leads.',
        'Campaign ID must uniquely identify a campaign across clients. Client may be Unstuck.',
        'Month: YYYY-MM. Active: Yes or No. Leads: whole number, including explicit 0.',
        'Only include a size band when it was actually targeted that month. Allocate each lead to exactly one band.',
        'Supported bands: 1-10, 11-50, 51-200, 201-500, 501-1000, 1001-5000, 5001-10000, 10001+.',
        'Reconcile the period totals to your lead records. This is a separate activity dataset, not joined automatically to raw lead rows.',
        'An active month means outbound ran during that calendar month. Partial months count once, not as a full-month projection.',
        'Inactive rows must have 0 leads. Reconcile lagged responses before using this same-month activity measure.',
        'The app validates the structure but cannot independently verify completeness of your activity roster.',
        'This template intentionally contains no example results. Upload the completed workbook via Excel snapshot.',
    ]:
        notes.append([text])
    notes.column_dimensions['A'].width = 130
    output = BytesIO()
    workbook.save(output)
    return Response(output.getvalue(), media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    headers={'Content-Disposition': 'attachment; filename="campaign-activity-template.xlsx"', 'Cache-Control': 'no-store'})
