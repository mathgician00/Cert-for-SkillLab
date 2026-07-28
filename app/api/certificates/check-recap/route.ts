import { NextRequest, NextResponse } from 'next/server';
import { getGoogleClients, getOrCreateFolder } from '@/lib/google-api';

const PARENT_FOLDER_NAME = 'SkillLab Certificate';

export async function POST(req: NextRequest) {
  try {
    const token = req.headers.get('Authorization')?.split('Bearer ')[1];
    if (!token) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { drive, sheets } = getGoogleClients(token);
    
    let { sheetUrl } = await req.json().catch(() => ({ sheetUrl: '' }));

    let recapSpreadsheetId = '';

    if (sheetUrl) {
      const match = sheetUrl.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
      if (!match) return NextResponse.json({ error: 'Invalid Google Sheets URL format.' }, { status: 400 });
      recapSpreadsheetId = match[1];
    } else {
      const templateId = process.env.TEMPLATE_SHEET_ID;
      if (!templateId) return NextResponse.json({ error: 'TEMPLATE_SHEET_ID not set.' }, { status: 500 });

      // Find or create Recap Sheet
      const q = `name contains 'Recap SkillLab @' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`;
      const res = await drive.files.list({ q, spaces: 'drive', fields: 'files(id, name, modifiedTime)', orderBy: 'modifiedTime desc' });
      
      if (res.data.files && res.data.files.length > 0) {
        recapSpreadsheetId = res.data.files[0].id!;
      } else {
        const rootFolderId = await getOrCreateFolder(drive, PARENT_FOLDER_NAME);
        const copyRes = await drive.files.copy({
          fileId: templateId,
          requestBody: {
            name: 'Recap SkillLab @branch',
            parents: [rootFolderId]
          }
        });
        recapSpreadsheetId = copyRes.data.id!;

        // Share the copy
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
      }
    }

    // Read the recap sheet
    const sheetData = await sheets.spreadsheets.values.get({
      spreadsheetId: recapSpreadsheetId,
      range: '01. Recap', // Ensure this tab exists
    });

    const rows = sheetData.data.values;
    if (!rows || rows.length < 2) {
      return NextResponse.json({ url: `https://docs.google.com/spreadsheets/d/${recapSpreadsheetId}`, rows: [], total: 0 });
    }

    const headers = rows[0].map((h: string) => h.toString().trim().toLowerCase());
    const nameIdx = headers.findIndex((h: string) => ['student name', 'name'].includes(h));
    const dateIdx = headers.findIndex((h: string) => h === 'date');
    const courseIdx = headers.findIndex((h: string) => h === 'course');
    const linkIdx = headers.findIndex((h: string) => h === 'link');

    if (nameIdx === -1 || dateIdx === -1 || courseIdx === -1 || linkIdx === -1) {
      return NextResponse.json({ error: 'Required columns not found in recap.' }, { status: 400 });
    }

    const pendingRows = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const link = row[linkIdx] ? row[linkIdx].toString().trim() : '';
      if (link !== '') continue;

      const name = row[nameIdx] ? row[nameIdx].toString().trim() : '';
      const rawDate = row[dateIdx] ? row[dateIdx].toString().trim() : '';
      const course = row[courseIdx] ? row[courseIdx].toString().trim() : '';

      if (!name && !rawDate && !course) continue;

      pendingRows.push({
        rowIndex: i + 1,
        linkColLetter: String.fromCharCode(65 + linkIdx),
        name,
        rawDate,
        course
      });
    }

    return NextResponse.json({ 
      url: `https://docs.google.com/spreadsheets/d/${recapSpreadsheetId}`,
      rows: pendingRows, 
      total: pendingRows.length 
    });

  } catch (error: any) {
    console.error('Check recap error:', error);
    return NextResponse.json({ error: error.message || 'Internal Server Error' }, { status: 500 });
  }
}
