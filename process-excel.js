const fs = require('fs');
const XLSX = require('xlsx');

const filePath = 'Weekly Report Master File- 2026.xlsx';

if (!fs.existsSync(filePath)) {
    console.error(`Excel file not found at: ${filePath}`);
    process.exit(1);
}

const workbook = XLSX.readFile(filePath, { cellDates: true });
const targetSheetName = workbook.SheetNames.includes('Master 2026') ? 'Master 2026' : workbook.SheetNames[0];
const worksheet = workbook.Sheets[targetSheetName];
const parsedData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

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

function getQuarter(dateStr) {
    if (!dateStr) return 'Q1';
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return 'Q1';
    const m = d.getMonth() + 1;
    if (m >= 1 && m <= 3) return 'Q1';
    if (m >= 4 && m <= 6) return 'Q2';
    if (m >= 7 && m <= 9) return 'Q3';
    return 'Q4';
}

const masterRecords = parsedData
    .filter(row => {
        const srNo = findVal(row, ['Sr No', 'id', 'No'], '');
        const pos = findVal(row, ['Positions', 'Position', 'Role Title'], '');
        return srNo !== '' || pos !== '';
    })
    .map((row, index) => {
        const joinDate = parseDate(findVal(row, ['Joining Date', 'Joined Date']));
        const explicitQ = String(findVal(row, ['Quarter Joining', 'Quarter'], '')).trim();
        const finalQ = (explicitQ && explicitQ !== '' && explicitQ !== 'undefined') ? explicitQ : getQuarter(joinDate);
        
        let statusVal = String(findVal(row, ['Status', 'Hiring Status'], 'Open')).trim();
        if (!statusVal || statusVal === '') statusVal = 'Open';

        return {
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
            join: joinDate,
            q: finalQ,
            aging: Number(findVal(row, ['Aging (Days)', 'Aging'], 0)),
            timeToOffer: Number(findVal(row, ['Position Offered', 'Time to Offer'], 0)),
            timeToClose: Number(findVal(row, ['Position Closed', 'Time to Close'], 0)),
            status: statusVal,
            cand: String(findVal(row, ['Offered Candidates', 'Candidate Name', 'Candidate'], '')).trim(),
            rec: String(findVal(row, ['Recruiter', 'Recruiter Name'], 'Unassigned')).trim(),
            sourceChannel: String(findVal(row, ['Applied/ Sourced/ Reffered (Internal/External)', 'Channel', 'Source'], 'External / Sourced')).trim(),
            sourceDetail: '',
            rag: String(findVal(row, ['RAG Status', 'RAG'], 'Green')).trim(),
            buddy: String(findVal(row, ['Buddy - Onboarding', 'Buddy'], '')).trim(),
            backup: String(findVal(row, ['Back-up Plans', 'Backup'], '')).trim(),
            comm: String(findVal(row, ['Comments', 'Remarks'], '')).trim()
        };
    });

fs.writeFileSync('data.json', JSON.stringify(masterRecords, null, 2));
console.log(`Successfully compiled ${masterRecords.length} records into data.json.`);

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
