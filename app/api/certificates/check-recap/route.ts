import { NextRequest, NextResponse } from 'next/server';
import { getGoogleClients, getOrCreateFolder, getServiceAccountSheetsClient } from '@/lib/google-api';

const PARENT_FOLDER_NAME = 'SkillLab Certificate';

export async function POST(req: NextRequest) {
  try {
    const token = req.headers.get('Authorization')?.split('Bearer ')[1];
    if (!token) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { drive, sheets } = getGoogleClients(token);

    let { sheetUrl } = await req.json().catch(() => ({ sheetUrl: '' }));

    let recapSpreadsheetId = '';
    let multiple = false;

    if (sheetUrl) {
      const match = sheetUrl.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
      if (!match) return NextResponse.json({ error: 'Invalid Google Sheets URL format.' }, { status: 400 });
      recapSpreadsheetId = match[1];
    } else {
      const templateId = process.env.TEMPLATE_SHEET_ID;
      if (!templateId) return NextResponse.json({ error: 'TEMPLATE_SHEET_ID not set.' }, { status: 500 });

      const rootFolderId = await getOrCreateFolder(drive, PARENT_FOLDER_NAME);

      const q = `name contains 'Recap SkillLab @' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false and '${rootFolderId}' in parents`;
      const res = await drive.files.list({ q, spaces: 'drive', fields: 'files(id, name, modifiedTime)', orderBy: 'modifiedTime desc' });

      const matches = res.data.files || [];

      if (matches.length > 0) {
        recapSpreadsheetId = matches[0].id!;
        multiple = matches.length > 1;

        // if (multiple) {
        //   // Best-effort audit log — never blocks the main flow if it fails.
        //   try {
        //     const authRes = await drive.about.get({ fields: 'user' });
        //     const branchEmail = authRes.data.user?.emailAddress || 'unknown';
        //     await logDuplicate(branchEmail, matches.length, matches.map((f: any) => f.name).join(', '));
        //   } catch (logErr: any) {
        //     console.error('Duplicate log failed (non-fatal):', logErr.message);
        //   }
        // }
      } else {
        const copyRes = await drive.files.copy({
          fileId: templateId,
          requestBody: {
            name: 'Recap SkillLab @branch',
            parents: [rootFolderId]
          }
        });
        recapSpreadsheetId = copyRes.data.id!;

        await drive.permissions.create({
          fileId: recapSpreadsheetId,
          requestBody: { role: 'reader', type: 'anyone' }
        });

        const adminEmail = process.env.ADMIN_EMAIL;
        if (adminEmail) {
          await drive.permissions.create({
            fileId: recapSpreadsheetId,
            requestBody: { role: 'writer', type: 'user', emailAddress: adminEmail }
          });
        }

        // Log every copy event (first-time or repeat) so patterns of
        // repeated re-copying by the same branch can be spotted via
        // conditional formatting / pivot on the Duplicates tab.
        try {
          const authRes = await drive.about.get({ fields: 'user' });
          const branchEmail = authRes.data.user?.emailAddress || 'unknown';
          await logDuplicate(branchEmail, 1, 'Recap SkillLab @branch (new copy)');
        } catch (logErr: any) {
          console.error('Copy-event log failed (non-fatal):', logErr.message);
        }
      }
    }

    // Resolve the actual tab name case-insensitively (avoids hard-fail on casing)
    const meta = await sheets.spreadsheets.get({
      spreadsheetId: recapSpreadsheetId,
      fields: 'sheets.properties.title'
    });
    const actualTab = (meta.data?.sheets || []).find(
      (s: any) => s.properties?.title?.trim().toLowerCase() === '01. recap'
    )?.properties?.title;

    if (!actualTab) {
      return NextResponse.json({ error: "No tab named '01. Recap' found in this spreadsheet." }, { status: 400 });
    }

    // UNFORMATTED_VALUE + SERIAL_NUMBER: dates come back as Sheets serial numbers,
    // not locale-formatted strings — avoids ambiguous day/month string parsing,
    // and avoids timezone drift since conversion is done with Date.UTC below.
    const sheetData = await sheets.spreadsheets.values.get({
      spreadsheetId: recapSpreadsheetId,
      range: `'${actualTab}'`,
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'SERIAL_NUMBER'
    });

    const rows = sheetData.data.values;
    if (!rows || rows.length < 2) {
      return NextResponse.json({
        url: `https://docs.google.com/spreadsheets/d/${recapSpreadsheetId}`,
        rows: [],
        total: 0,
        multiple
      });
    }

    const headers = rows[0].map((h: string) => h.toString().trim().toLowerCase());
    const nameIdx   = headers.findIndex((h: string) => ['student name', 'name'].includes(h));
    const dateIdx   = headers.findIndex((h: string) => h === 'date');
    const courseIdx = headers.findIndex((h: string) => h === 'course');
    const linkIdx   = headers.findIndex((h: string) => h === 'link');
    const statusIdx = headers.findIndex((h: string) => h === 'status');

    if (nameIdx === -1 || dateIdx === -1 || courseIdx === -1 || linkIdx === -1) {
      return NextResponse.json({ error: 'Required columns not found in recap.' }, { status: 400 });
    }
    if (statusIdx === -1) {
      return NextResponse.json({ error: "Column 'Status' not found in recap." }, { status: 400 });
    }

    const pendingRows = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const status = row[statusIdx] ? row[statusIdx].toString().trim() : '';
      const link   = row[linkIdx]   ? row[linkIdx].toString().trim()   : '';

      // Gate: only rows explicitly marked ready by the sheet's own formula,
      // and not already generated. The formula upstream already guarantees
      // name/date/course are non-empty whenever status = 'Siap cetak'.
      if (status !== 'Siap cetak') continue;
      if (link !== '') continue;

      const name        = row[nameIdx]   ? row[nameIdx].toString().trim()   : '';
      const rawDateCell = row[dateIdx];
      const course      = row[courseIdx] ? row[courseIdx].toString().trim() : '';

      let dateInfo: { y: number; m: number; d: number } | null = null;
      let displayDate = '';

      if (typeof rawDateCell === 'number' && rawDateCell > 1000) {
        const ms = Date.UTC(1899, 11, 30) + rawDateCell * 86400000; // Sheets epoch
        const dt = new Date(ms);
        dateInfo = { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
        displayDate = `${String(dateInfo.d).padStart(2, '0')}/${String(dateInfo.m).padStart(2, '0')}/${dateInfo.y}`;
      } else if (rawDateCell) {
        displayDate = rawDateCell.toString().trim();
      }

      pendingRows.push({
        rowIndex: i + 1,
        linkColLetter: colLetter(linkIdx + 1),
        statusColLetter: colLetter(statusIdx + 1),
        name,
        rawDate: displayDate,
        dateY: dateInfo?.y ?? null,
        dateM: dateInfo?.m ?? null,
        dateD: dateInfo?.d ?? null,
        course
      });
    }

    return NextResponse.json({
      url: `https://docs.google.com/spreadsheets/d/${recapSpreadsheetId}`,
      rows: pendingRows,
      total: pendingRows.length,
      multiple
    });

  } catch (error: any) {
    console.error('Check recap error:', error);
    return NextResponse.json({ error: error.message || 'Internal Server Error' }, { status: 500 });
  }
}

// Multi-letter safe column index -> letter (A, ..., Z, AA, AB, ...)
function colLetter(col: number): string {
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// Records to the Master sheet's Duplicates tab whenever a branch has more
// than one "Recap SkillLab @" file in their folder — lets the admin follow
// up and consolidate. Creates the tab on first use.
// Records to the dedicated Log sheet's Duplicates tab whenever the app
// creates a new copy of the template for a branch (first-time or repeat) —
// lets the admin spot patterns of repeated re-copying via conditional
// formatting / pivot on the sheet. Creates the tab on first use.
async function logDuplicate(branchEmail: string, count: number, fileNames: string) {
  const logSheetId = process.env.LOG_SHEET_ID;
  if (!logSheetId) return;

  const logSheets = getServiceAccountSheetsClient();

  const meta = await logSheets.spreadsheets.get({
    spreadsheetId: logSheetId,
    fields: 'sheets.properties.title'
  });
  const existingTab = (meta.data?.sheets || []).find(
    (s: any) => s.properties?.title?.trim().toLowerCase() === 'duplicates'
  )?.properties?.title;

  if (!existingTab) {
    await logSheets.spreadsheets.batchUpdate({
      spreadsheetId: logSheetId,
      requestBody: { requests: [{ addSheet: { properties: { title: 'Duplicates' } } }] }
    });
    await logSheets.spreadsheets.values.append({
      spreadsheetId: logSheetId,
      range: 'Duplicates',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [['Timestamp', 'Branch Email', 'File Count', 'File Names']] }
    });
  }

  await logSheets.spreadsheets.values.append({
    spreadsheetId: logSheetId,
    range: 'Duplicates',
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[new Date().toISOString(), branchEmail, count, fileNames]] }
  });
}
