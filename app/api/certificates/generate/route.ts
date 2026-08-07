import { NextRequest, NextResponse } from 'next/server';
import { getGoogleClients, getOrCreateFolder, getServiceAccountSheetsClient } from '@/lib/google-api';

const PARENT_FOLDER_NAME = 'SkillLab Certificate';
const DAYS   = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];

export async function POST(req: NextRequest) {
  try {
    const token = req.headers.get('Authorization')?.split('Bearer ')[1];
    if (!token) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { drive, sheets, slides } = getGoogleClients(token);
    const { sheetUrl, rows, userEmail } = await req.json();

    const masterId = process.env.MASTER_SHEET_ID;
    if (!masterId) throw new Error('MASTER_SHEET_ID not set.');

    const recapMatch = sheetUrl.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!recapMatch) throw new Error('Invalid recap sheet URL.');
    const recapSpreadsheetId = recapMatch[1];

    // Resolve actual recap tab name case-insensitively
    const recapMeta = await sheets.spreadsheets.get({
      spreadsheetId: recapSpreadsheetId,
      fields: 'sheets.properties.title'
    });
    const recapTab = recapMeta.data.sheets?.find(
      (s: any) => s.properties.title.trim().toLowerCase() === '01. recap'
    )?.properties.title;
    if (!recapTab) throw new Error("Could not find '01. Recap' tab to write back to.");

    // Resolve actual Master tab name case-insensitively
    const masterMeta = await sheets.spreadsheets.get({
      spreadsheetId: masterId,
      fields: 'sheets.properties.title'
    });
    const masterTab = masterMeta.data.sheets?.find(
      (s: any) => s.properties.title.trim().toLowerCase() === 'master'
    )?.properties.title;
    if (!masterTab) throw new Error("No tab named 'Master' found in the Master spreadsheet.");

    const masterData = await sheets.spreadsheets.values.get({
      spreadsheetId: masterId,
      range: `'${masterTab}'`,
    });

    const masterRows = masterData.data.values || [];
    const masterHeaders = masterRows[0]?.map((h: string) => h.toString().trim().toLowerCase()) || [];
    const courseIdx = masterHeaders.indexOf('course');
    const certifIdx = masterHeaders.findIndex((h: string) => h === 'certif id' || h === 'certif_id' || h === 'certificate id');

    if (courseIdx === -1 || certifIdx === -1) {
      throw new Error('Required columns not found in Master sheet.');
    }

    const templateMap = new Map<string, string>();
    for (let i = 1; i < masterRows.length; i++) {
      const course = masterRows[i][courseIdx]?.toString().trim().toLowerCase();
      const id = masterRows[i][certifIdx]?.toString().trim();
      if (course && id) templateMap.set(course, id);
    }

    const results = [];
    const parentFolderId = await getOrCreateFolder(drive, PARENT_FOLDER_NAME);

    // Each entry updates BOTH the link cell and the status cell for its row.
    const recapUpdates: { range: string; values: any[][] }[] = [];
    const logEntries: any[][] = [];

    for (const row of rows) {
      try {
        // Status column already gates completeness upstream (formula-driven on
        // the sheet side) — 'Siap cetak' only appears when name/date/course are
        // all non-empty. No redundant completeness check needed here.

        const templateId = templateMap.get(row.course.toLowerCase());
        if (!templateId) throw new Error(`No template for course "${row.course}"`);

        // Date arrives pre-resolved from check-recap as {dateY, dateM, dateD}
        // (from Sheets' serial number) — deterministic, no string re-parsing.
        const y = row.dateY, m = row.dateM, d = row.dateD;
        if (y == null || m == null || d == null) throw new Error('Invalid date.');

        const dateObj = new Date(Date.UTC(y, m - 1, d));
        const dayName = DAYS[dateObj.getUTCDay()];
        const dd      = String(d).padStart(2, '0');
        const mmmm    = MONTHS[m - 1];
        const yymmdd  = `${String(y).slice(-2)}${String(m).padStart(2, '0')}${dd}`;
        const formattedDate = `${dayName}, ${dd} ${mmmm} ${y}`;

        const subFolderName = `${yymmdd} - ${row.course}`;
        const subFolderId   = await getOrCreateFolder(drive, subFolderName, parentFolderId);

        const fileName = `${row.name} - ${subFolderName}`;

        const copyRes = await drive.files.copy({
          fileId: templateId,
          requestBody: { name: fileName, parents: [subFolderId] },
        });
        const newSlideId = copyRes.data.id!;

        await slides.presentations.batchUpdate({
          presentationId: newSlideId,
          requestBody: {
            requests: [
              { replaceAllText: { containsText: { text: '{{name}}',   matchCase: false }, replaceText: row.name } },
              { replaceAllText: { containsText: { text: '{{date}}',   matchCase: false }, replaceText: formattedDate } },
              { replaceAllText: { containsText: { text: '{{course}}', matchCase: false }, replaceText: row.course } },
            ]
          }
        });

        const pdfExport = await drive.files.export({
          fileId: newSlideId,
          mimeType: 'application/pdf'
        }, { responseType: 'stream' });

        const pdfRes = await drive.files.create({
          requestBody: { name: `${fileName}.pdf`, parents: [subFolderId] },
          media: { mimeType: 'application/pdf', body: pdfExport.data },
          fields: 'id, webViewLink'
        });

        const pdfFileId = pdfRes.data.id!;
        const pdfLink   = pdfRes.data.webViewLink!;

        await drive.permissions.create({
          fileId: pdfFileId,
          requestBody: { role: 'reader', type: 'anyone' }
        });

        await drive.files.update({ fileId: newSlideId, requestBody: { trashed: true } });

        recapUpdates.push({
          range: `'${recapTab}'!${row.linkColLetter}${row.rowIndex}`,
          values: [[pdfLink]]
        });
        recapUpdates.push({
          range: `'${recapTab}'!${row.statusColLetter}${row.rowIndex}`,
          values: [['Selesai']]
        });

        logEntries.push([new Date().toISOString(), userEmail, sheetUrl, row.name, row.course, formattedDate, pdfLink]);
        results.push({ rowIndex: row.rowIndex, success: true, fileName: `${fileName}.pdf`, pdfLink });

      } catch (err: any) {
        console.error(`Row ${row.rowIndex} error:`, err);

        // Mark the row as failed directly in the sheet, so the branch sees it
        // without needing to check this app's ephemeral log panel.
        recapUpdates.push({
          range: `'${recapTab}'!${row.statusColLetter}${row.rowIndex}`,
          values: [['Gagal']]
        });

        results.push({ rowIndex: row.rowIndex, error: err.message });
      }
    }

    if (recapUpdates.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: recapSpreadsheetId,
        requestBody: { valueInputOption: 'USER_ENTERED', data: recapUpdates }
      });
    }

    // Log write is best-effort: certificates and sheet updates above are already
    // committed, so a logging failure must never fail the request or flip
    // already-successful rows to an error state.
    if (logEntries.length > 0) {
      try {
        const logSheets = getServiceAccountSheetsClient();
        await logSheets.spreadsheets.values.append({
          spreadsheetId: masterId,
          range: 'CertLog',
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: logEntries }
        });
      } catch (logErr: any) {
        console.error('Log write failed (non-fatal):', logErr.message);
      }
    }

    return NextResponse.json({ results });
  } catch (error: any) {
    console.error('Generate error:', error);
    return NextResponse.json({ error: error.message || 'Internal Server Error' }, { status: 500 });
  }
}
