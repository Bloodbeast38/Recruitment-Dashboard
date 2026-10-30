const fs = require('fs');
const XLSX = require('xlsx');

const filePath = 'Weekly Report Master File- 2026.xlsx';

if (!fs.existsSync(filePath)) {
    console.error(`Excel file not found at: ${filePath}`);
    process.exit(1);
}

const workbook = XLSX.readFile(filePath, { cellDates: true });

// Target the Master 2026 sheet directly
let targetSheetName = workbook.SheetNames.find(name => 
    name.toLowerCase().includes('master 2026') || name.toLowerCase().includes('master')
);
if (!targetSheetName) {
    targetSheetName = workbook.SheetNames[0];
}

const worksheet = workbook.Sheets[targetSheetName];
const parsedData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

// Intelligent getter for column names
function findVal(row, keys, defaultVal = '') {
    const rowKeys = Object.keys(row);
    for (const k of keys) {
        const found = rowKeys.find(rk => rk.toLowerCase().trim() === k.toLowerCase().trim() || rk.toLowerCase().includes(k.toLowerCase()));
        if (found !== undefined && row[found] !== '' && row[found] !== null) {
            return row[found];
        }
    }
    return defaultVal;
}

const masterRecords = parsedData
    .filter(row => {
        const srNo = findVal(row, ['Sr No', 'id', 'No'], '');
        const pos = findVal(row, ['Positions', 'Position', 'Role Title'], '');
        return srNo !== '' || pos !== '';
    })
    .map((row, index) => ({
        id: Number(findVal(row, ['Sr No', 'id', 'No'], index + 1)),
        geo: String(findVal(row, ['Geo', 'Geography', 'Region'], 'India')).trim(),
        dept: String(findVal(row, ['Department', 'Practice', 'Function'], 'General')).trim(),
        squad: String(findVal(row, ['Squads', 'Project', 'Team'], '-')).trim(),
        pos: String(findVal(row, ['Positions', 'Position', 'Role Title'], 'Unnamed Position')).trim(),
        head: String(findVal(row, ['Dept Head', 'Head', 'Manager'], '-')).trim(),
        type: String(findVal(row, ['New/Replacement', 'Type'], 'New')).trim(),
        replaced: String(findVal(row, ['Replaced Employee Name', 'Replaced Employee'], '')).trim(),
        empType: String(findVal(row, ['Full-Time/Consultant', 'Employment Type'], 'Full-time')).trim(),
        loc: String(row['Hiring Location'] || findVal(row, ['Location', 'City'], '-')).trim(),
        open: parseDate(findVal(row, ['Position Open Date', 'Open Date'])),
        offer: parseDate(findVal(row, ['Date Offered', 'Offer Date'])),
        join: parseDate(findVal(row, ['Joining Date', 'Joined Date'])),
        q: String(findVal(row, ['Quarter Joining', 'Quarter'], 'Q1')).trim(),
        aging: Number(findVal(row, ['Aging (Days)', 'Aging'], 0)),
        timeToOffer: Number(findVal(row, ['Position Offered', 'Time to Offer'], 0)),
        timeToClose: Number(findVal(row, ['Position Closed', 'Time to Close'], 0)),
        status: String(findVal(row, ['Status', 'Hiring Status'], 'Open')).trim(),
        cand: String(findVal(row, ['Offered Candidates', 'Candidate Name', 'Candidate'], '')).trim(),
        rec: String(findVal(row, ['Recruiter', 'Recruiter Name'], 'Unassigned')).trim(),
        sourceChannel: String(findVal(row, ['Applied/ Sourced/ Reffered (Internal/External)', 'Channel', 'Source'], 'External / Sourced')).trim(),
        sourceDetail: '',
        rag: String(findVal(row, ['RAG Status', 'RAG'], 'Green')).trim(),
        buddy: String(findVal(row, ['Buddy - Onboarding', 'Buddy'], '')).trim(),
        backup: String(findVal(row, ['Back-up Plans', 'Backup'], '')).trim(),
        comm: String(findVal(row, ['Comments', 'Remarks'], '')).trim()
    }));

fs.writeFileSync('data.json', JSON.stringify(masterRecords, null, 2));
console.log(`Successfully compiled ${masterRecords.length} master records into data.json.`);

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
