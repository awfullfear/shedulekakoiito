// Vercel Serverless Function: /api/schedule
const XLSX = require('xlsx');

const PUBLIC_KEY = 'https://disk.yandex.ru/d/7XafcYMVe1ec7g';
const API_URL = `https://cloud-api.yandex.net/v1/disk/public/resources?public_key=${encodeURIComponent(PUBLIC_KEY)}&limit=100`;

const TIME_REGEX = /(\d{1,2}[:.]\d{2})\s*[-–—]\s*(\d{1,2}[:.]\d{2})/;
const ROOM_REGEX = /(?:,\s*)?(?:(?:Ауд\.?|АУд\.?)\s*[\w\d/]*\s*(?:\([^)]+\))?|(?:Актовый\s*зал|Конференц[- ]*зал|Спорт\.?\s*зал|Он[- ]*лайн|ДО)(?:\s*\([^)]+\))?|\((?:Конференц[- ]*зал|Актовый\s*зал)\))\s*$/i;
const TEACHER_REGEX = /([А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ]\.+\s*[А-ЯЁ]\.*|\s+[А-ЯЁ][а-яё]+))\s*$/;
const LESSON_TYPE_REGEX = /\(\s*(лек|лекц|лекция|пр|практ|практика|семинар|лаб|зачет|зачёт|экзамен|конс)\.?\s*\)/i;

function parseDayOfWeek(text) {
    if (!text) return null;
    const clean = String(text).trim().toLowerCase();
    if (clean === 'пн' || clean.startsWith('понед')) return 1;
    if (clean === 'вт' || clean.startsWith('втор')) return 2;
    if (clean === 'ср' || clean.startsWith('сред')) return 3;
    if (clean === 'чт' || clean.startsWith('четв')) return 4;
    if (clean === 'пт' || clean.startsWith('пятн')) return 5;
    if (clean === 'сб' || clean.startsWith('субб')) return 6;
    if (clean === 'вс' || clean.startsWith('воскр')) return 7;
    return null;
}

function parseDateFromCell(cell) {
    if (!cell || cell.v === undefined || cell.v === null) return null;
    if (cell.t === 'd' || cell.v instanceof Date) {
        const d = cell.v instanceof Date ? cell.v : new Date(cell.w || cell.v);
        return d.toISOString().split('T')[0];
    }
    const text = String(cell.w || cell.v || '').trim();
    if (!text) return null;

    // YYYY-MM-DD
    const m1 = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m1) {
        return `${m1[1]}-${m1[2].padStart(2, '0')}-${m1[3].padStart(2, '0')}`;
    }
    // DD.MM.YYYY
    const m2 = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})/);
    if (m2) {
        let y = parseInt(m2[3], 10);
        if (y < 100) y += 2000;
        return `${y}-${m2[2].padStart(2, '0')}-${m2[1].padStart(2, '0')}`;
    }
    return null;
}

function parseFileNameDate(name) {
    if (!name) return null;
    const m = name.match(/(?<!\d)(\d{1,2})[._-](\d{1,2})(?:[._-](\d{2,4}))?(?!\d)/);
    if (m) {
        let y = m[3] ? parseInt(m[3], 10) : null;
        if (y && y < 100) y += 2000;
        return { day: parseInt(m[1], 10), month: parseInt(m[2], 10), year: y };
    }
    return null;
}

function normalizeContent(s) {
    if (!s) return '';
    return String(s).toLowerCase().replace(/[^a-zа-яё0-9]/g, '');
}

function getPairNumber(timeStr) {
    const clean = String(timeStr).replace(/:/g, '.').replace(/\s+/g, '');
    if (clean.startsWith('9.00') || clean.startsWith('09.00')) return '1 пара';
    if (clean.startsWith('10.30')) return '2 пара';
    if (clean.startsWith('12.05')) return '3 пара';
    if (clean.startsWith('13.45')) return '4 пара';
    if (clean.startsWith('15.20')) return '5 пара';
    if (clean.startsWith('16.55')) return '6 пара';
    if (clean.startsWith('18.30')) return '7 пара';
    if (clean.startsWith('20.00')) return '8 пара';
    return '';
}

function parseWorkbook(buffer, fileName = '') {
    const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:Z100');

    // Groups in row 2 (index 1)
    const groupsByCol = {};
    const allGroups = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = sheet[XLSX.utils.encode_cell({ r: 1, c })];
        if (cell && cell.v && String(cell.v).includes('26')) {
            const g = String(cell.v).trim();
            groupsByCol[c] = g;
            if (!allGroups.includes(g)) allGroups.push(g);
        }
    }

    // Detected dates
    const detectedDates = [];
    for (let r = range.s.r; r <= range.e.r; r++) {
        let foundDate = null;
        let foundDow = null;
        let foundDowText = '';

        for (let c = 0; c <= Math.min(7, range.e.c); c++) {
            const cell = sheet[XLSX.utils.encode_cell({ r, c })];
            if (!cell) continue;
            if (!foundDate) {
                const d = parseDateFromCell(cell);
                if (d) foundDate = d;
            }
            const txt = String(cell.w || cell.v || '').trim();
            if (!foundDow && txt) {
                const dow = parseDayOfWeek(txt);
                if (dow) {
                    foundDow = dow;
                    foundDowText = txt;
                }
            }
        }

        if (foundDate) {
            detectedDates.push({ r, rawDate: foundDate, dow: foundDow, dowText: foundDowText });
        }
    }

    const fnDate = parseFileNameDate(fileName);
    const finalDateRows = [];

    for (let i = 0; i < detectedDates.length; i++) {
        const item = detectedDates[i];
        const dateObj = new Date(item.rawDate + 'T00:00:00Z');
        const calDow = dateObj.getUTCDay() === 0 ? 7 : dateObj.getUTCDay();
        let resolved = item.rawDate;

        if (item.dow && item.dow !== calDow) {
            let diff = item.dow - calDow;
            if (diff < -3) diff += 7;
            if (diff > 3) diff -= 7;
            dateObj.setUTCDate(dateObj.getUTCDate() + diff);
            resolved = dateObj.toISOString().split('T')[0];
        } else if (detectedDates.length === 1 && fnDate) {
            const cellD = dateObj.getUTCDate();
            const cellM = dateObj.getUTCMonth() + 1;
            if (cellD !== fnDate.day || cellM !== fnDate.month) {
                dateObj.setUTCDate(fnDate.day);
                dateObj.setUTCMonth(fnDate.month - 1);
                if (fnDate.year) dateObj.setUTCFullYear(fnDate.year);
                resolved = dateObj.toISOString().split('T')[0];
            }
        }

        finalDateRows.push({ r: item.r, date: resolved });
    }

    const schedulesByGroup = {};
    for (const g of allGroups) {
        schedulesByGroup[g] = { days: {}, availableDates: [] };
    }

    for (let i = 0; i < finalDateRows.length; i++) {
        const startR = finalDateRows[i].r;
        const dateStr = finalDateRows[i].date;
        const endR = i + 1 < finalDateRows.length ? finalDateRows[i + 1].r : Math.min(startR + 25, range.e.r + 1);

        const rowTimes = {};
        const rowContents = {};

        for (let r = startR; r < endR; r++) {
            let timeStr = '';
            for (let c = 0; c <= Math.min(6, range.e.c); c++) {
                const cell = sheet[XLSX.utils.encode_cell({ r, c })];
                if (!cell) continue;
                const val = String(cell.w || cell.v || '');
                const m = val.match(TIME_REGEX);
                if (m) {
                    timeStr = `${m[1]} – ${m[2]}`;
                    break;
                }
            }
            rowTimes[r] = timeStr;

            for (const [col, gname] of Object.entries(groupsByCol)) {
                const cell = sheet[XLSX.utils.encode_cell({ r, c: parseInt(col, 10) })];
                if (!cell) continue;
                const txt = String(cell.w || cell.v || '').trim();
                if (txt && txt.toLowerCase() !== 'none' && !txt.toLowerCase().includes('отмена')) {
                    rowContents[`${r},${col}`] = txt;
                }
            }
        }

        for (const [col, gname] of Object.entries(groupsByCol)) {
            const lessons = [];
            for (let r = startR; r < endR; r++) {
                const rawText = rowContents[`${r},${col}`];
                if (!rawText) continue;

                const text = rawText.replace(/\r/g, '').replace(/tc=\{[^}]+\}[^\n]*/g, '').replace(/(МСА|Светлана -):.*/g, '').replace(/\n+/g, '\n').trim();
                if (!text) continue;

                const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
                let rawSubject = lines[0] || '';
                let teacher = '';
                let room = '';

                if (lines.length > 1) {
                    for (let j = 1; j < lines.length; j++) {
                        let line = lines[j];
                        const mRoom = line.match(ROOM_REGEX);
                        if (mRoom) {
                            room = mRoom[0].trim().replace(/^,\s*/, '');
                            line = line.replace(ROOM_REGEX, '').trim();
                        }
                        if (line && !teacher) teacher = line;
                    }
                } else {
                    let line = lines[0];
                    const mRoom = line.match(ROOM_REGEX);
                    if (mRoom) {
                        room = mRoom[0].trim().replace(/^,\s*/, '');
                        line = line.replace(ROOM_REGEX, '').trim();
                    }
                    const mTeach = line.match(TEACHER_REGEX);
                    if (mTeach) {
                        teacher = mTeach[1].trim();
                        rawSubject = line.replace(TEACHER_REGEX, '').trim().replace(/,\s*$/, '');
                    } else {
                        rawSubject = line;
                    }
                }

                let lessonType = '';
                let cleanSubject = rawSubject;
                const mType = rawSubject.match(LESSON_TYPE_REGEX);
                if (mType) {
                    const rawLt = mType[1].toLowerCase();
                    if (rawLt.includes('лекц') || rawLt.includes('лек')) lessonType = 'Лекция';
                    else if (rawLt.includes('практ') || rawLt.includes('пр')) lessonType = 'Практическое занятие';
                    else if (rawLt.includes('семинар')) lessonType = 'Семинар';
                    else if (rawLt.includes('лаб')) lessonType = 'Лабораторная';
                    else if (rawLt.includes('зачет') || rawLt.includes('зачёт')) lessonType = 'Зачет';
                    else if (rawLt.includes('экзамен')) lessonType = 'Экзамен';
                    else if (rawLt.includes('конс')) lessonType = 'Консультация';
                    cleanSubject = rawSubject.replace(LESSON_TYPE_REGEX, '').trim().replace(/^,\s*|\s*,\s*$/g, '');
                }

                const normCell = normalizeContent(rawText);
                const jointGroups = [];
                if (normCell.length >= 4 && !/^\d+$/.test(normCell)) {
                    for (const [otherCol, otherGname] of Object.entries(groupsByCol)) {
                        if (otherCol === col || otherGname === gname || jointGroups.includes(otherGname)) continue;
                        const otherVal = rowContents[`${r},${otherCol}`];
                        if (otherVal && otherVal.length >= 4 && normalizeContent(otherVal) === normCell) {
                            jointGroups.push(otherGname);
                        }
                    }
                }

                const timeStr = rowTimes[r] || '';
                lessons.push({
                    pairNumber: getPairNumber(timeStr),
                    time: timeStr,
                    subject: cleanSubject || rawSubject,
                    cleanSubject: cleanSubject || rawSubject,
                    lessonType,
                    room,
                    teacher,
                    jointGroups: jointGroups.sort()
                });
            }

            if (lessons.length || !(dateStr in schedulesByGroup[gname].days)) {
                schedulesByGroup[gname].days[dateStr] = lessons;
                if (!schedulesByGroup[gname].availableDates.includes(dateStr)) {
                    schedulesByGroup[gname].availableDates.push(dateStr);
                }
            }
        }
    }

    for (const g of allGroups) {
        schedulesByGroup[g].availableDates.sort();
    }

    return { groups: allGroups, schedules: schedulesByGroup };
}

module.exports = async function handler(req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        return res.status(200).end();
    }

    try {
        const fetch = global.fetch || require('node-fetch');
        const apiResp = await fetch(API_URL, { headers: { 'User-Agent': 'Mozilla/5.0' } });
        if (!apiResp.ok) {
            return res.status(apiResp.status).json({ error: `Yandex Disk API returned status ${apiResp.status}` });
        }

        const diskData = await apiResp.json();
        const items = (diskData._embedded && diskData._embedded.items) || [];
        const xlsxItems = items.filter(it => (it.name || '').toLowerCase().endsWith('.xlsx'));

        if (!xlsxItems.length) {
            return res.status(404).json({ error: 'No .xlsx files found on Yandex Disk' });
        }

        // Sort latest first
        xlsxItems.sort((a, b) => new Date(b.modified || b.created).getTime() - new Date(a.modified || a.created).getTime());
        const latest = xlsxItems[0];

        const fileResp = await fetch(latest.file, { headers: { 'User-Agent': 'Mozilla/5.0' } });
        if (!fileResp.ok) {
            return res.status(500).json({ error: `Failed to download ${latest.name}` });
        }
        const buffer = Buffer.from(await fileResp.arrayBuffer());
        const parsed = parseWorkbook(buffer, latest.name);

        // Merge with weekly if single day
        const availableDatesCount = Math.max(...Object.values(parsed.schedules).map(s => s.availableDates.length));
        if (availableDatesCount < 4 && xlsxItems.length > 1) {
            const weeklyCandidate = xlsxItems.find(it => it !== latest && (it.name.includes('-') || it.name.includes('–') || it.name.toLowerCase().includes('недел'))) || xlsxItems[1];
            if (weeklyCandidate && weeklyCandidate.file) {
                try {
                    const secResp = await fetch(weeklyCandidate.file, { headers: { 'User-Agent': 'Mozilla/5.0' } });
                    if (secResp.ok) {
                        const secBuf = Buffer.from(await secResp.arrayBuffer());
                        const secParsed = parseWorkbook(secBuf, weeklyCandidate.name);
                        for (const g of parsed.groups) {
                            if (secParsed.schedules[g]) {
                                const merged = { ...secParsed.schedules[g].days, ...parsed.schedules[g].days };
                                parsed.schedules[g].days = merged;
                                parsed.schedules[g].availableDates = Object.keys(merged).sort();
                            }
                        }
                    }
                } catch (_) {}
            }
        }

        res.setHeader('Cache-Control', 'public, s-maxage=300, stale-while-revalidate=600');
        return res.status(200).json({
            lastUpdated: Date.now(),
            fileName: latest.name,
            yandexDiskUrl: PUBLIC_KEY,
            groups: parsed.groups,
            schedules: parsed.schedules
        });
    } catch (e) {
        return res.status(500).json({ error: e.message || String(e) });
    }
};
