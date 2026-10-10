"""IARE nominal rolls and activity exports. All reconciliation uses Python sets."""
import io
import os
import json
import re
import sys
from collections import Counter
from datetime import date, datetime
from functools import lru_cache
from pathlib import Path

import openpyxl
import pdfplumber

ROLL = r'\d{5}A[A-Z0-9]{4}'
ROOT = Path(__file__).resolve().parents[1]
ROMAN = {3: 'III', 5: 'V'}


def text(value):
    return re.sub(r'\s+', ' ', str(value or '')).strip()


def branch(value):
    value = text(value).upper()
    return {'CSE (AI&ML)': 'AIML', 'CSE (AIML)': 'AIML', 'CSE (DS)': 'DS', 'CSE (CS)': 'CS', 'AERONAUTICAL ENGINEERING': 'AERO', 'INFORMATION TECHNOLOGY': 'IT', 'COMPUTER SCIENCE AND ENGINEERING': 'CSE', 'ELECTRICAL AND ELECTRONICS ENGINEERING': 'EEE', 'ELECTRONICS AND COMMUNICATION ENGINEERING': 'ECE', 'MECHANICAL ENGINEERING': 'MECH'}.get(value, value)


def nominal_index(folder):
    students = {}
    warnings = []
    for file in sorted(Path(folder).glob('*/*STUDENTS LIST*.xlsx')):
        semester = 3 if 'III SEM' in file.name else 5
        book = openpyxl.load_workbook(file, read_only=True, data_only=True)
        for sheet in book:
            for row in sheet.values:
                if len(row) < 5 or not re.fullmatch(ROLL, text(row[1]).upper()):
                    continue
                roll = text(row[1]).upper()
                record = {'roll': roll, 'name': text(row[2]), 'branch': branch(row[3]), 'batch': text(row[4]).upper(), 'semester': semester, 'section': ''}
                if roll in students:
                    if any(students[roll][k] != record[k] for k in ('branch', 'batch', 'semester')):
                        warnings.append(f'{roll}: conflicting master batch/branch. The first total-sheet record is retained; review the nominal rolls.')
                else:
                    students[roll] = record
        book.close()
    if not students:
        raise ValueError('No master student lists found in NOMINAL ROLLS.')
    for file in sorted(Path(folder).rglob('*.xlsx')):
        if 'STUDENTS LIST' in file.name:
            continue
        book = openpyxl.load_workbook(file, read_only=True, data_only=True)
        sheet = book.worksheets[0]  # Other sheets are unlabeled helper lists, not authoritative rosters.
        rows = list(sheet.values)
        header = next((i for i, r in enumerate(rows[:20]) if any(text(c).lower() == 'roll no' for c in r)), None)
        if header is None:
            book.close()
            continue
        columns = {text(c).lower(): i for i, c in enumerate(rows[header]) if c}
        roster_sessions = set(re.findall(r'\b(?:FN|AN)\b', ' '.join(file.relative_to(folder).parts[:-1]).upper()))
        if len(roster_sessions) > 1:
            raise ValueError(f'{file.name}: conflicting FN/AN roster folders.')
        session = next(iter(roster_sessions), '')
        for row in rows[header + 1:]:
            roll = text(row[columns['roll no']]).upper()
            if roll not in students:
                continue
            section = text(row[columns['sec']]).upper() if 'sec' in columns else ''
            if students[roll]['section'] and section and students[roll]['section'] != section:
                raise ValueError(f'{roll}: conflicting sections in nominal rolls.')
            students[roll]['section'] = section or students[roll]['section']
            if session:
                if students[roll].get('session', session) != session:
                    raise ValueError(f'{roll}: conflicting FN/AN sessions in nominal rolls.')
                students[roll]['session'] = session
        book.close()
    return {'students': students, 'warnings': sorted(set(warnings))}


def nominal_folder():
    configured = os.getenv('NOMINAL_ROLLS_DIR')
    if configured:
        folder = Path(configured)
        if not folder.is_dir() or not any(folder.glob('*/*STUDENTS LIST*.xlsx')):
            raise ValueError('NOMINAL_ROLLS_DIR must point to the complete nominal-roll database folder.')
        return folder
    candidates = [Path(sys.executable).parent / 'NOMINAL ROLLS', ROOT.parent / 'NOMINAL ROLLS', ROOT / 'NOMINAL ROLLS']
    return next((p for p in candidates if p.is_dir() and any(p.glob('*/*STUDENTS LIST*.xlsx'))), None)


@lru_cache(maxsize=1)
def load_index():
    folder = nominal_folder()
    if folder:
        return nominal_index(folder)
    bundled = ROOT / 'assets' / 'student-index.json'
    if bundled.exists():
        return json.loads(bundled.read_text(encoding='utf-8'))
    raise ValueError('Student database unavailable. Put NOMINAL ROLLS next to the application.')


def database_info():
    index = load_index()
    return {'students': len(index['students']), 'cdc_students': sum(has_cdc_batch(r) for r in index['students'].values()), 'semesters': dict(Counter(str(r['semester']) for r in index['students'].values())), 'warnings': index['warnings'], 'odd_students': odd_students(index), 'source': 'NOMINAL ROLLS folder' if nominal_folder() else 'Bundled nominal-roll index'}


def has_cdc_batch(student):
    return bool(text(student['batch'])) and text(student['batch']).upper() != 'UNSPECIFIED'


def odd_students(index):
    """Records that need correction before a CDC student can be trusted."""
    problems = {}
    for roll, student in index['students'].items():
        if has_cdc_batch(student) and not student['section']:
            problems[roll] = {'roll': roll, 'name': student['name'], 'semester': student['semester'], 'branch': student['branch'], 'batch': student['batch'], 'issue': 'Missing class section', 'needed': 'Confirm the actual section and add the roll to its section roster.'}
    for warning in index['warnings']:
        roll = warning.split(':', 1)[0]
        student = index['students'].get(roll)
        if student and has_cdc_batch(student):
            problems[roll] = {'roll': roll, 'name': student['name'], 'semester': student['semester'], 'branch': student['branch'], 'batch': student['batch'], 'issue': 'Conflicting master batch or branch', 'needed': 'Check the master sheets and correct the conflicting values.'}
    return [problems[roll] for roll in sorted(problems)]


def activity_tables(data, filename):
    from .engine import tables, key
    if Path(filename).suffix.lower() != '.pdf':
        return tables(data, filename)
    import pandas as pd
    frames = {}
    try:
        with pdfplumber.open(io.BytesIO(data)) as document:
            if not 1 <= len(document.pages) <= 100:
                raise ValueError('Use a PDF with 1–100 pages.')
            for number, page in enumerate(document.pages, 1):
                if not (page.extract_text() or '').strip():
                    raise ValueError(f'Page {number} has no readable text. Scanned PDFs need OCR first.')
                found = False
                for index, rows in enumerate(page.extract_tables()):
                    headers = next((i for i, row in enumerate(rows[:20]) if {'date','period','section','absentees','staff'} <= {key(c) for c in row}), None)
                    if headers is None:
                        continue
                    normalized = [[text(c) for c in row] for row in rows[headers:]]
                    # Page headers repeat in print exports; they are not activity records.
                    normalized = [normalized[0]] + [r for r in normalized[1:] if r != normalized[0]]
                    frames[f'Page {number}, table {index+1}'] = pd.DataFrame(normalized).fillna('')
                    found = True
                if not found:
                    raise ValueError(f'PDF page {number}: no readable activity table. Export a table with Date, Period, Section, Absentees and Staff headers on every page, with visible cell borders.')
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError('Cannot read this activity PDF. Use an unlocked PDF with selectable text and a clear table.') from exc
    return frames


def activities(data, filename):
    from .engine import tables, key
    result = []
    for sheet_name, frame in activity_tables(data, filename).items():
        for index, row in frame.head(20).iterrows():
            cols = {key(v): i for i, v in enumerate(row) if text(v)}
            if not {'date', 'period', 'section', 'absentees', 'staff'} <= set(cols):
                continue
            for n, values in frame.iloc[index + 1:].iterrows():
                if not any(text(v) for v in values):
                    continue
                if {'date','period','section','absentees','staff'} <= {key(v) for v in values}:
                    continue
                raw_date = values.iloc[cols['date']]
                try:
                    day = raw_date.date() if isinstance(raw_date, datetime) else date.fromisoformat(text(raw_date)[:10])
                    hour = int(float(text(values.iloc[cols['period']])))
                except (ValueError, TypeError):
                    raise ValueError(f'{filename}, row {n + 1}: invalid Date or Period.')
                if not 1 <= hour <= 12:
                    raise ValueError(f'{filename}, row {n + 1}: invalid period.')
                section = text(values.iloc[cols['section']]).upper()
                match = re.fullmatch(r'(.+?)-(III|V)-SEM-([A-Z])', section)
                if not match:
                    raise ValueError(f'{filename}, row {n + 1}: unrecognized section “{section}”. Expected AIML-V-SEM-A.')
                absent_text = text(values.iloc[cols['absentees']]).upper()
                absent = set(re.findall(r'\b' + ROLL + r'\b', absent_text))
                residue = re.sub(r'\b' + ROLL + r'\b', '', absent_text)
                if absent_text not in {'', 'NIL', 'NONE', '0', '-', 'NO ABSENTEES'} and re.sub(r'[\s,;|]+', '', residue):
                    raise ValueError(f'{filename}, row {n + 1}: an absentee ID could not be read. Correct that cell.')
                staff = text(values.iloc[cols['staff']])
                faculty = re.match(r'(IARE\d+)\s*[-–]?\s*(.*)', staff, re.I)
                if not faculty:
                    raise ValueError(f'{filename}, row {n + 1}: staff must include faculty ID and name.')
                result.append({'date': day.isoformat(), 'hour': hour, 'branch': branch(match[1]), 'semester': 3 if match[2] == 'III' else 5, 'section': match[3], 'absentees': sorted(absent), 'faculty_id': faculty[1].upper(), 'faculty_name': faculty[2], 'source': filename})
            break
    if not result:
        raise ValueError('No activity records found. Expected Date, Period, Section, Absentees and Staff columns.')
    return result


def cdc_input(data, filename):
    """Read a CDC summary PDF, an activity table, or a complete student register."""
    from .engine import tables, key, ALIASES, PRESENT, ABSENT
    if Path(filename).suffix.lower() == '.pdf':
        try:
            with pdfplumber.open(io.BytesIO(data)) as pdf:
                first = pdf.pages[0].extract_text() if pdf.pages else ''
        except Exception as exc:
            raise ValueError('Cannot open this CDC PDF. Use a valid, unlocked PDF with selectable text.') from exc
        if re.search(r'\b(FN|AN) Attendance Summary', first or '', re.I):
            return cdc_pdf(data, filename)
        frames = activity_tables(data, filename)
    else:
        frames = tables(data, filename)
    index = load_index()['students']
    student_rows = {}
    scopes = set()
    activity_format = any({'date','period','section','absentees','staff'} <= {key(c) for c in row} for frame in frames.values() for _, row in frame.head(20).iterrows())
    if activity_format:
        records = activities(data, filename)
        seen = {}
        for row in records:
            session = 'FN' if row['hour'] <= 3 else 'AN' if row['hour'] <= 6 else ''
            if not session:
                raise ValueError('CDC activity periods must be 1–3 (FN) or 4–6 (AN).')
            group = (row['branch'],row['section'])
            absent = set(row['absentees'])
            section_students = {roll:r for roll,r in index.items() if r['semester']==row['semester'] and r['branch']==row['branch'] and r['section']==row['section']}
            if not section_students or absent - set(section_students):
                raise ValueError('CDC activity IDs or sections do not match the nominal rolls. Update the database.')
            cohort = {roll:r for roll,r in section_students.items() if has_cdc_batch(r)}
            if not cohort:
                continue
            scopes.add((row['date'],session,row['semester']))
            cdc_absent = absent & set(cohort)
            if group in seen and seen[group] != cdc_absent:
                raise ValueError('CDC activity hours have different CDC student absentee lists. Supply one CDC session attendance record per section.')
            seen[group] = cdc_absent
            student_rows.update({roll:(r, roll in absent) for roll,r in cohort.items()})
    else:
        for frame in frames.values():
            for header, values in frame.head(20).iterrows():
                cols = {key(v):i for i,v in enumerate(values)}
                roll_col = next((cols[k] for k in ALIASES['roll'] if k in cols),None)
                status_col = next((cols[k] for k in ('status','attendance','attendancestatus','cdcattendance') if k in cols),None)
                if roll_col is None or status_col is None or not {'date','session','semester'} <= set(cols):
                    continue
                for _, values in frame.iloc[header+1:].iterrows():
                    if not any(text(v) for v in values):
                        continue
                    roll = text(values.iloc[roll_col]).upper()
                    if not re.fullmatch(ROLL,roll) or roll not in index:
                        raise ValueError(f'CDC register: unknown or invalid roll number {roll}.')
                    if not has_cdc_batch(index[roll]):
                        continue
                    day = date.fromisoformat(text(values.iloc[cols['date']])[:10]).isoformat()
                    session = text(values.iloc[cols['session']]).upper()
                    semester = {'III':3,'3':3,'3.0':3,'V':5,'5':5,'5.0':5}.get(text(values.iloc[cols['semester']]).upper())
                    status = text(values.iloc[status_col]).lower()
                    if session not in {'FN','AN'} or semester != index[roll]['semester'] or status not in PRESENT | ABSENT:
                        raise ValueError(f'{roll}: check Session (FN/AN), Semester (III/V) and Status (Present/Absent).')
                    item = (index[roll], status in ABSENT)
                    if roll in student_rows and student_rows[roll] != item:
                        raise ValueError(f'{roll}: conflicting CDC attendance entries.')
                    scopes.add((day,session,semester))
                    student_rows[roll] = item
                break
    if not student_rows:
        raise ValueError('No CDC students found. CDC Excel/CSV needs Date, Session, Semester, Roll Number and Status columns, or the class-style activity table.')
    if len(scopes) != 1:
        raise ValueError('Use one date, semester and FN/AN session per CDC file.')
    day, session, semester = scopes.pop()
    coverage = {}
    absent = {}
    for roll,(student,is_absent) in student_rows.items():
        group = (student['batch'],student['branch'])
        counts = coverage.setdefault(group, {'total':0,'absent':0})
        counts['total'] += 1; counts['absent'] += int(is_absent)
        if is_absent:
            absent[roll] = {'batch':group[0],'branch':group[1]}
    return {'date':day,'session':session,'semester':semester,'absentees':absent,'coverage':[{'batch':b,'branch':br,**v} for (b,br),v in coverage.items()], 'cohort_rolls':list(student_rows),'pages':0,'filename':filename}


def cdc_pdf(data, filename):
    if Path(filename).suffix.lower() != '.pdf':
        raise ValueError('For the college report, upload the CDC attendance-summary PDF like 12.pdf, 13.pdf or 14.pdf.')
    absent = {}
    coverage = {}
    declared_branches = {}
    with pdfplumber.open(io.BytesIO(data)) as document:
        if not 1 <= len(document.pages) <= 100:
            raise ValueError('Use a CDC PDF with at most 100 pages.')
        pages = [p.extract_text() or '' for p in document.pages]
        if any(not p.strip() for p in pages):
            raise ValueError('CDC PDF contains a page without readable text. Scanned PDFs require OCR first.')
        meta = re.search(r'\b(FN|AN) Attendance Summary\s*-\s*(\d{2}-\d{2}-\d{4})', pages[0], re.I)
        semester = re.search(r'B\.Tech\s+(III|V)\s+Semester', pages[0], re.I)
        if not meta or not semester:
            raise ValueError('CDC PDF must identify the session, date and semester, as in the supplied 12/13/14 PDFs.')
        day = datetime.strptime(meta[2], '%d-%m-%Y').date().isoformat()
        for line in pages[0].splitlines():
            row = re.fullmatch(r'(\S+)\s+(.+?)\s+(\d+)\s+(\d+)\s+(\d+)', line)
            if row and row[1] != 'TOTAL':
                group = (row[1].upper(), branch(row[2]))
                values = [int(row[i]) for i in (3, 4, 5)]
                if values[0] != values[1] + values[2] or group in coverage:
                    raise ValueError('CDC PDF summary has inconsistent or duplicate totals.')
                coverage[group] = {'total': values[0], 'absent': values[2]}
        for page_no, page in enumerate(pages, 1):
            heading = re.search(r'(?:III|V) SEM (.+?)\s*-\s*Absentees List', page)
            if not heading:
                if re.search(r'\b' + ROLL + r'\b', page):
                    raise ValueError(f'CDC PDF page {page_no}: student IDs are outside a labeled absentee list.')
                continue
            department = branch(heading[1])
            for line in page.splitlines():
                if not re.search(r'\b' + ROLL + r'\b', line):
                    continue
                match = re.match(r'^\d+\s+(' + ROLL + r')\s+(.+?)\s+(S[A-Z]\d+)\s*$', line, re.I)
                if not match:
                    raise ValueError(f'CDC PDF page {page_no}: cannot reliably read an absentee row.')
                roll, batch = match[1].upper(), match[3].upper()
                record = {'batch': batch, 'branch': department}
                if roll in absent:
                    raise ValueError(f'CDC PDF contains a repeated absentee ID: {roll}.')
                absent[roll] = record
            total = re.search(r'Total Absent:\s*(\d+)', page)
            if total:
                declared_branches[department] = int(total[1])
    if not coverage:
        raise ValueError('No batch/branch attendance summary found on the first PDF page.')
    observed = Counter((r['batch'], r['branch']) for r in absent.values())
    for group, counts in coverage.items():
        if observed[group] != counts['absent']:
            raise ValueError(f'CDC {group[0]} {group[1]}: extracted {observed[group]} absentees but summary declares {counts["absent"]}.')
    if set(observed) - set(coverage):
        raise ValueError('CDC absentee list includes a batch/branch missing from its summary.')
    for department, count in declared_branches.items():
        if sum(r['branch'] == department for r in absent.values()) != count:
            raise ValueError(f'CDC {department}: absentee list total does not match extracted IDs.')
    return {'date': day, 'session': meta[1].upper(), 'semester': 3 if semester[1].upper() == 'III' else 5, 'absentees': absent, 'coverage': [{'batch': b, 'branch': br, **v} for (b, br), v in coverage.items()], 'pages': len(pages), 'filename': filename}


def compare_college(cdc, activity_rows, selected_date):
    if selected_date != cdc['date']:
        raise ValueError(f'CDC file is dated {cdc["date"]}; select that report date and matching class files.')
    matching = [r for r in activity_rows if r['date'] == selected_date and r['semester'] == cdc['semester']]
    if not matching:
        dates = ', '.join(sorted({r['date'] for r in activity_rows}))
        semesters = ', '.join(ROMAN[s] for s in sorted({r['semester'] for r in activity_rows}))
        raise ValueError(f'No matching class records. CDC: {ROMAN[cdc["semester"]]} semester, {cdc["date"]}. Class file(s): {semesters} semester, date(s) {dates}. Upload files for the same date and semester.')
    index = load_index()
    students = index['students']
    excluded = {}
    def omit(roll, reason, source):
        item = excluded.setdefault(roll, {'Roll Number': roll, 'Student Name': students.get(roll, {}).get('name', ''), 'Sources': set(), 'Reasons': set()})
        item['Sources'].add(source)
        item['Reasons'].add(reason)
    coverage = {(r['batch'], r['branch']) for r in cdc['coverage']}
    cohort = {roll: r for roll, r in students.items() if has_cdc_batch(r) and r['semester'] == cdc['semester'] and (r['batch'], r['branch']) in coverage}
    # A batch match alone cannot establish presence in the other CDC session.
    cohort = {roll: r for roll, r in cohort.items() if r.get('session', cdc['session']) == cdc['session']}
    if 'cohort_rolls' in cdc:
        included = set(cdc['cohort_rolls'])
        cohort = {roll:r for roll,r in cohort.items() if roll in included}
    # An explicit CDC absentee roll is attendance evidence; batch/name labels
    # must not suppress it. Keep section/semester/session checks for class scope.
    for roll in cdc['absentees']:
        student = students.get(roll)
        if student and student['semester'] == cdc['semester'] and student.get('session', cdc['session']) == cdc['session']:
            cohort[roll] = student
    if not cohort:
        raise ValueError('No students in the nominal rolls match the CDC attendance rolls or batch/branch summary.')
    warnings = []
    for roll, student in cohort.items():
        if not student['section']:
            omit(roll, 'Class section is missing from nominal rolls.', 'Database')
    for warning in index['warnings']:
        roll = warning.split(':', 1)[0]
        if roll in cohort and roll not in cdc['absentees']:
            omit(roll, 'Conflicting master batch or branch; correct the master sheets.', 'Database')
    for roll, record in cdc['absentees'].items():
        student = students.get(roll)
        if student is None:
            continue  # A matching class row can establish the section below.
        if roll not in cohort:
            omit(roll, 'CDC semester or session does not match the nominal rolls.', 'CDC')
        elif student['branch'] != record['branch']:
            omit(roll, 'CDC branch does not match the class roster; class membership is unresolved.', 'CDC')
    sessions = {'FN': range(1, 4), 'AN': range(4, 7)}
    unique = {}
    for activity in matching:
        if activity['hour'] not in sessions[cdc['session']]:
            continue
        group = (activity['branch'], activity['hour'], activity['section'])
        if group in unique and any(activity[k] != unique[group][k] for k in ('absentees', 'faculty_id', 'faculty_name')):
            raise ValueError(f'Conflicting activity records for {group}. Upload one final record per class/hour.')
        unique[group] = activity
    for (department, hour, section), activity in unique.items():
        if not any(r['branch'] == department and r['section'] == section for r in cohort.values()):
            continue
        for roll in activity['absentees']:
            if roll in cohort and cohort[roll]['section'] and (cohort[roll]['branch'] != department or cohort[roll]['section'] != section):
                omit(roll, f'Class file section differs from nominal rolls ({department}-{section}, hour {hour}).', 'Class')
    for roll in cdc['absentees']:
        if roll not in students and roll not in cohort and roll not in excluded:
            omit(roll, 'CDC absentee ID has no nominal-roll record or matching class section.', 'CDC')
    # Remove questionable records from both attendance sets before any counts.
    cohort = {roll: student for roll, student in cohort.items() if roll not in excluded}
    for group in cdc['coverage']:
        count = sum((r['batch'], r['branch']) == (group['batch'], group['branch']) and roll in students for roll, r in cohort.items())
        if count != group['total']:
            warnings.append(f'{group["batch"]} {group["branch"]}: CDC source strength {group["total"]}; verified nominal-roll students available for comparison {count}.')
    rows, details, skipped = [], [], []
    order = {'CSE': 0, 'EEE': 1, 'IT': 2, 'AIML': 3, 'DS': 4, 'ECE': 5}
    for (department, hour, section), activity in sorted(unique.items(), key=lambda p: (order.get(p[0][0], 99), p[0][0], p[0][1], p[0][2])):
        section_cohort = {roll for roll, r in cohort.items() if r['branch'] == department and r['section'] == section}
        if not section_cohort:
            if any(r['semester'] == cdc['semester'] and r['branch'] == department and r['section'] == section and has_cdc_batch(r) for r in students.values()):
                skipped.append(f'{department}-{ROMAN[cdc["semester"]]}-{section}, hour {hour}')
            continue
        cdc_absent = section_cohort & set(cdc['absentees'])
        dept_absent = section_cohort & set(activity['absentees'])
        # Count each CDC-absent roll missing from this class absentee list once.
        discrepancy_rolls = cdc_absent - dept_absent
        rows.append({'department': department, 'hour': hour, 'section': section, 'faculty_id': activity['faculty_id'], 'faculty_name': activity['faculty_name'], 'cdc_absent': len(cdc_absent), 'department_absent': len(dept_absent), 'department_present': len(section_cohort - dept_absent), 'discrepancy': len(discrepancy_rolls), 'cdc_only': len(cdc_absent - dept_absent), 'class_only': len(dept_absent - cdc_absent), 'both': len(cdc_absent & dept_absent), 'cohort': len(section_cohort)})
        for roll in sorted(section_cohort):
            a, b = roll in cdc_absent, roll in dept_absent
            student = cohort[roll]
            details.append({'Batch Code': student['batch'], 'Branch': department, 'Roll Number': roll, 'Student Name': student.get('name', ''), 'Section': section, 'Hour': hour, 'CDC Attendance': 'Absent' if a else 'Present', 'Class Attendance': 'Absent' if b else 'Present', 'Missing Attendance': 'Both' if a and b else 'CDC' if a else 'Class' if b else 'None', 'Discrepancy': 'Yes' if roll in discrepancy_rolls else 'No'})
    if not rows:
        raise ValueError(f'No valid CDC students remain for the uploaded {cdc["session"]} class records. Check matching sections and the nominal rolls. FN uses periods 1–3; AN uses periods 4–6.')
    if skipped:
        warnings.append('Excluded class groups outside the supplied CDC cohort: ' + '; '.join(skipped))
    covered = {(r['department'], r['section']) for r in rows}
    missing_groups = {(r['branch'], r['section']) for r in cohort.values() if r['section']} - covered
    if missing_groups:
        warnings.append('Class files do not cover all CDC sections. This report includes only uploaded matching activity records.')
    missing_periods = [f'{br}-{sec}: ' + ','.join(str(h) for h in sessions[cdc['session']] if not any(r['department'] == br and r['section'] == sec and r['hour'] == h for r in rows)) for br, sec in covered if any(not any(r['department'] == br and r['section'] == sec and r['hour'] == h for r in rows) for h in sessions[cdc['session']])]
    if missing_periods:
        warnings.append('Periods not supplied: ' + '; '.join(sorted(missing_periods)))
    omissions = [{'Roll Number': r['Roll Number'], 'Student Name': r['Student Name'], 'Source': ', '.join(sorted(r['Sources'])), 'Reason': ' '.join(sorted(r['Reasons']))} for _, r in sorted(excluded.items())]
    if omissions:
        warnings.insert(0, f'Partial report: {len(omissions)} student IDs were excluded from both attendance counts. See Excluded Records in Excel.')
    represented = {r['Roll Number'] for r in details}
    return {'date': selected_date, 'semester': cdc['semester'], 'session': cdc['session'], 'rows': rows, 'details': details, 'warnings': list(dict.fromkeys(warnings)), 'excluded': omissions, 'partial': bool(omissions), 'sources': [cdc['filename']] + sorted({r['source'] for r in matching}), 'cdc_students': len(represented), 'cdc_absentees': len(represented & set(cdc['absentees']))}
