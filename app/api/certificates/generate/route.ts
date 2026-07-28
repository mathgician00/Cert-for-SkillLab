import { NextRequest, NextResponse } from 'next/server';
import { getGoogleClients, getOrCreateFolder } from '@/lib/google-api';

const PARENT_FOLDER_NAME = 'SkillLab Certificate';

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

    // Get master sheet templates
    const masterData = await sheets.spreadsheets.values.get({
      spreadsheetId: masterId,
      range: 'Master',
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

    // Prepare update data for recap sheet
    const recapUpdates = [];
    // Prepare log entries
    const logEntries = [];

    for (const row of rows) {
      try {
        const templateId = templateMap.get(row.course.toLowerCase());
        if (!templateId) throw new Error(`No template for course ${row.course}`);

        // Parse date to YYMMDD and display format
        // Basic parsing for this example
        const dateObj = new Date(row.rawDate);
        let yymmdd = "000000";
        let formattedDate = row.rawDate;
        
        if (!isNaN(dateObj.getTime())) {
          const y = dateObj.getFullYear().toString().slice(-2);
          const m = (dateObj.getMonth() + 1).toString().padStart(2, '0');
          const d = dateObj.getDate().toString().padStart(2, '0');
          yymmdd = `${y}${m}${d}`;
          formattedDate = dateObj.toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
        }

        if (yymmdd === "000000") throw new Error('Invalid date');

        const subFolderName = `${yymmdd} - ${row.course}`;
        const subFolderId = await getOrCreateFolder(drive, subFolderName, parentFolderId);

        const fileName = `${row.name} - ${subFolderName}`;

        // 1. Copy slide
        const copyRes = await drive.files.copy({
          fileId: templateId,
          requestBody: { name: fileName, parents: [subFolderId] },
        });
        const newSlideId = copyRes.data.id!;

        // 2. Replace text
        await slides.presentations.batchUpdate({
          presentationId: newSlideId,
          requestBody: {
            requests: [
              { replaceAllText: { containsText: { text: '{{name}}', matchCase: false }, replaceText: row.name } },
              { replaceAllText: { containsText: { text: '{{date}}', matchCase: false }, replaceText: formattedDate } },
              { replaceAllText: { containsText: { text: '{{course}}', matchCase: false }, replaceText: row.course } },
            ]
          }
        });

        // 3. Export PDF
        const pdfExport = await drive.files.export({
          fileId: newSlideId,
          mimeType: 'application/pdf'
        }, { responseType: 'stream' });

        // 4. Create new PDF file in Drive
        const pdfRes = await drive.files.create({
          requestBody: { name: `${fileName}.pdf`, parents: [subFolderId] },
          media: { mimeType: 'application/pdf', body: pdfExport.data },
          fields: 'id, webViewLink'
        });

        const pdfFileId = pdfRes.data.id!;
        const pdfLink = pdfRes.data.webViewLink!;

        // 5. Share PDF
        await drive.permissions.create({
          fileId: pdfFileId,
          requestBody: { role: 'reader', type: 'anyone' }
        });

        // 6. Trash the slide copy
        await drive.files.update({ fileId: newSlideId, requestBody: { trashed: true } });

        recapUpdates.push({
          range: `01. Recap!${row.linkColLetter}${row.rowIndex}`,
          values: [[pdfLink]]
        });

        logEntries.push([new Date().toISOString(), userEmail, sheetUrl, row.name, row.course, formattedDate, pdfLink]);
        results.push({ rowIndex: row.rowIndex, success: true, fileName: `${fileName}.pdf`, pdfLink });

      } catch (err: any) {
        console.error(`Row ${row.rowIndex} error:`, err);
        results.push({ rowIndex: row.rowIndex, error: err.message });
      }
    }

    // Write back to Recap Sheet
    if (recapUpdates.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: recapSpreadsheetId,
        requestBody: { valueInputOption: 'USER_ENTERED', data: recapUpdates }
      });
    }

    // Log to Master Sheet
    if (logEntries.length > 0) {
      try {
        await sheets.spreadsheets.values.append({
          spreadsheetId: masterId,
          range: 'CertLog',
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: logEntries }
        });
      } catch (e) {
        // Ignore if CertLog doesn't exist
        console.error('Failed to log to CertLog', e);
      }
    }

    return NextResponse.json({ results });
  } catch (error: any) {
    console.error('Generate error:', error);
    return NextResponse.json({ error: error.message || 'Internal Server Error' }, { status: 500 });
  }
}
