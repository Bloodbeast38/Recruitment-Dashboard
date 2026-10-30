const fs = require('fs');
const XLSX = require('xlsx');

const filePath = 'Weekly Report Master File- 2026.xlsx';

if (!fs.existsSync(filePath)) {
    console.error(`Excel file not found at: ${filePath}`);
    process.exit(1);
}

const workbook = XLSX.readFile(filePath, { cellDates: true });

let targetSheetName = workbook.SheetNames.find(name => 
    name.toLowerCase().includes('master') || name.toLowerCase().includes('2026')
);
if (!targetSheetName) {
    targetSheetName = workbook.SheetNames[0];
}

const worksheet = workbook.Sheets[targetSheetName];
const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

let headerRowIdx = 0;
for (let i = 0; i < Math.min(10, rawData.length); i++) {
    const rowStr = rawData[i].map(cell => String(cell).toLowerCase()).join(" ");
    if (rowStr.includes("positions") || rowStr.includes("geo") || rowStr.includes("status") || rowStr.includes("sr no")) {
        headerRowIdx = i;
        break;
    }
}

const parsedData = XLSX.utils.sheet_to_json(worksheet, { range: headerRowIdx, defval: "" });

const masterRecords = parsedData
    .filter(row => row['Positions'] || row['Geo'] || row['Status'] || row['Sr No'])
    .map((row, index) => ({
        id: Number(row['Sr No'] || row['id'] || (index + 1)),
        geo: String(row['Geo'] || 'India').trim(),
        dept: String(row['Department'] || row['Practice'] || 'General').trim(),
        squad: String(row['Squads'] || row['Project'] || '-').trim(),
        pos: String(row['Positions'] || 'Unnamed Position').trim(),
        head: String(row['Dept Head'] || '-').trim(),
        type: String(row['New/Replacement'] || 'New').trim(),
        replaced: String(row['Replaced Employee Name'] || row['Replaced Employee'] || '').trim(),
        empType: String(row['Full-Time/Consultant'] || 'Full-time').trim(),
        loc: String(row['Hiring Location'] || '-').trim(),
        open: parseDate(row['Position Open Date']),
        offer: parseDate(row['Date Offered']),
        join: parseDate(row['Joining Date']),
        q: String(row['Quarter Joining'] || 'Q1').trim(),
        aging: Number(row['Aging (Days)'] || 0),
        timeToOffer: Number(row['Position Offered'] || 0),
        timeToClose: Number(row['Position Closed'] || 0),
        status: String(row['Status'] || 'Open').trim(),
        cand: String(row['Offered Candidates'] || '').trim(),
        rec: String(row['Recruiter'] || row['Recruiter Name'] || 'Unassigned').trim(),
        sourceChannel: String(row['Applied/ Sourced/ Reffered (Internal/External)'] || 'External / Sourced').trim(),
        sourceDetail: '',
        rag: String(row['RAG Status'] || 'Green').trim(),
        buddy: String(row['Buddy - Onboarding'] || '').trim(),
        backup: String(row['Back-up Plans'] || '').trim(),
        comm: String(row['Comments'] || row['Remarks '] || '').trim()
    }));

fs.writeFileSync('data.json', JSON.stringify(masterRecords, null, 2));
console.log(`Successfully processed ${masterRecords.length} records into data.json from sheet: "${targetSheetName}"`);

function parseDate(value) {
    if (!value) return '';
    if (value instanceof Date) return value.toISOString().split('T')[0];
    if (typeof value === 'number') {
        const utcDays = Math.floor(value - 25569);
        const dateObj = new Date(utcDays * 86400 * 1000);
        return !isNaN(dateObj.getTime()) ? dateObj.toISOString().split('T')[0] : String(value);
    }
    return String(value);
}

