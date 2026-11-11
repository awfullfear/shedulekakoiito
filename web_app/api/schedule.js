// Vercel Serverless Function: /api/schedule
const XLSX = require('xlsx');

const PUBLIC_KEY = 'https://disk.yandex.ru/d/7XafcYMVe1ec7g';
const API_URL = `https://cloud-api.yandex.net/v1/disk/public/resources?public_key=${encodeURIComponent(PUBLIC_KEY)}&limit=100`;

const TIME_REGEX = /(\d{1,2}[:.]\d{2})\s*[-–—]\s*(\d{1,2}[:.]\d{2})/;
const ROOM_REGEX = /(?:,\s*)?(?:(?:Ауд\.?|АУд\.?)\s*[\w\d/]*\s*(?:\([^)]+\))?|(?:Актовый\s*зал|Конференц[- ]*зал|Спорт\.?\s*зал|Он[- ]*лайн|ДО)(?:\s*\([^)]+\))?|\((?:Конференц[- ]*зал|Актовый\s*зал)\))\s*$/i;
const TEACHER_REGEX = /([А-ЯЁа-яё][а-яё]{1,}\s+[А-ЯЁ]\.+\s*[А-ЯЁ]\.*)\s*$/;
const LESSON_TYPE_REGEX = /\(\s*(практическ(?:ое|ие)\s+заняти[ея]|практ\.?\s*заняти[ея]|практик[аи]|практ|пр\.?\s*з|пр|лекци[яи]|лекц|лек|семинарск(?:ое|ие)\s+заняти[ея]|семинары|семинар|сем|лабораторн(?:ая|ые)\s+работ[аы]|лабораторн(?:ая|ые)|лаб|дифференцированн(?:ый)?\s*зач[её]т|диф\.?\s*зач[её]т|зач[её]т|экзамен|экз|консультаци[яи]|конс)\.?\s*\)/gi;

function normalizeLessonTypeLabel(rawType) {
    if (!rawType) return '';
    const t = String(rawType).trim().toLowerCase();
    if (t.startsWith('лек')) return 'Лекция';
    if (t.startsWith('пр')) return 'Практическое занятие';
    if (t.startsWith('сем')) return 'Семинар';
    if (t.startsWith('лаб')) return 'Лабораторная';
    if (t.includes('зач')) return 'Зачёт';
    if (t.startsWith('экз')) return 'Экзамен';
    if (t.startsWith('конс')) return 'Консультация';
    return String(rawType).trim();
}

function cleanSubjectText(text) {
    return String(text || '')
        .replace(/\s+/g, ' ')
        .replace(/\s+([,.;:!?])/g, '$1')
        .replace(/^[,.;:\s-]+|[,;\s-]+$/g, '')
        .trim();
}

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
    const coveredDatesSet = new Set();
    for (const g of allGroups) {
        schedulesByGroup[g] = { days: {}, availableDates: [] };
    }

    for (let i = 0; i < finalDateRows.length; i++) {
        const startR = finalDateRows[i].r;
        const dateStr = finalDateRows[i].date;
        if (dateStr) coveredDatesSet.add(dateStr);
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

                let workingText = text;
                let teacher = '';
                let room = '';
                let lessonType = '';

                const mRoom = workingText.match(ROOM_REGEX);
                if (mRoom) {
                    room = mRoom[0].trim().replace(/^,\s*/, '');
                    workingText = workingText.substring(0, workingText.length - mRoom[0].length).trim();
                }

                workingText = workingText.replace(LESSON_TYPE_REGEX, (_, rawType) => {
                    if (!lessonType) lessonType = normalizeLessonTypeLabel(rawType);
                    return ' ';
                });

                workingText = cleanSubjectText(workingText);

                const mTeach = workingText.match(TEACHER_REGEX);
                if (mTeach) {
                    teacher = mTeach[1].trim();
                    workingText = workingText.substring(0, workingText.length - mTeach[0].length).trim();
                }

                const cleanSubject = cleanSubjectText(workingText) || cleanSubjectText(text);
                const rawSubject = cleanSubjectText(text);

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

    return {
        groups: allGroups,
        coveredDates: Array.from(coveredDatesSet).sort(),
        schedules: schedulesByGroup
    };
}

const HISTORY_RETENTION_DAYS = 7;

function getMskTodayStr() {
    const mskNow = new Date(Date.now() + 3 * 60 * 60 * 1000);
    return mskNow.toISOString().slice(0, 10);
}

function shiftDateStr(dateStr, deltaDays) {
    const [y, m, d] = String(dateStr).split('-').map(Number);
    const dt = new Date(Date.UTC(y, m - 1, d + deltaDays));
    return dt.toISOString().slice(0, 10);
}

function collectCoveredDates(data) {
    const dates = new Set();
    if (!data || typeof data !== 'object') return dates;
    if (Array.isArray(data.coveredDates)) {
        for (const d of data.coveredDates) {
            if (d) dates.add(d);
        }
    }
    if (data.schedules && typeof data.schedules === 'object') {
        for (const gdata of Object.values(data.schedules)) {
            if (gdata && typeof gdata === 'object' && gdata.days && typeof gdata.days === 'object') {
                for (const d of Object.keys(gdata.days)) {
                    if (d) dates.add(d);
                }
            }
        }
    }
    return dates;
}

function pruneScheduleData(data, retainDays = HISTORY_RETENTION_DAYS, refTodayStr = getMskTodayStr()) {
    if (!data || typeof data !== 'object') return data;
    const cutoffStr = shiftDateStr(refTodayStr, -retainDays);
    const prunedSchedules = {};
    for (const [g, gdata] of Object.entries(data.schedules || {})) {
        if (!gdata || typeof gdata !== 'object') continue;
        const keptDays = {};
        for (const d of Object.keys(gdata.days || {}).sort()) {
            if (d >= cutoffStr) keptDays[d] = gdata.days[d];
        }
        prunedSchedules[g] = {
            days: keptDays,
            availableDates: Object.keys(keptDays).sort()
        };
    }
    const covered = Array.from(collectCoveredDates(data)).filter(d => d >= cutoffStr).sort();
    return {
        ...data,
        coveredDates: covered,
        schedules: prunedSchedules
    };
}

function mergeScheduleData(baseData, incomingData, retainDays = HISTORY_RETENTION_DAYS, refTodayStr = getMskTodayStr()) {
    const cutoffStr = shiftDateStr(refTodayStr, -retainDays);
    if (!incomingData || !incomingData.schedules || Object.keys(incomingData.schedules).length === 0) {
        return baseData ? pruneScheduleData(baseData, retainDays, refTodayStr) : incomingData;
    }
    if (!baseData || !baseData.schedules || Object.keys(baseData.schedules).length === 0) {
        return pruneScheduleData(incomingData, retainDays, refTodayStr);
    }

    const incomingCovered = collectCoveredDates(incomingData);
    const baseCovered = collectCoveredDates(baseData);
    const baseScheds = baseData.schedules || {};
    const incScheds = incomingData.schedules || {};

    const incomingGroups = Array.isArray(incomingData.groups) ? [...incomingData.groups] : [];
    const mergedGroups = [...incomingGroups];
    for (const g of baseData.groups || []) {
        if (!mergedGroups.includes(g)) {
            const hasRetained = Object.keys(baseScheds[g]?.days || {}).some(d => d >= cutoffStr && !incomingCovered.has(d));
            if (hasRetained) mergedGroups.push(g);
        }
    }

    const mergedSchedules = {};
    const allGroupKeys = new Set([...mergedGroups, ...Object.keys(baseScheds), ...Object.keys(incScheds)]);
    for (const g of allGroupKeys) {
        const mergedDays = {};
        for (const [d, lessons] of Object.entries(baseScheds[g]?.days || {})) {
            if (d >= cutoffStr && !incomingCovered.has(d)) mergedDays[d] = lessons;
        }
        for (const [d, lessons] of Object.entries(incScheds[g]?.days || {})) {
            if (d >= cutoffStr) mergedDays[d] = lessons;
        }
        const sortedDays = {};
        for (const d of Object.keys(mergedDays).sort()) {
            sortedDays[d] = mergedDays[d];
        }
        if (incomingGroups.includes(g) || Object.keys(sortedDays).length > 0) {
            mergedSchedules[g] = {
                days: sortedDays,
                availableDates: Object.keys(sortedDays)
            };
        }
    }

    const mergedCovered = Array.from(new Set([...baseCovered, ...incomingCovered])).filter(d => d >= cutoffStr).sort();
    return {
        ...baseData,
        ...incomingData,
        groups: mergedGroups,
        coveredDates: mergedCovered,
        schedules: mergedSchedules
    };
}

function countRecentWorkdays(data, refTodayStr = getMskTodayStr()) {
    const covered = collectCoveredDates(data);
    const [y, m, d] = refTodayStr.split('-').map(Number);
    const todayUtc = new Date(Date.UTC(y, m - 1, d));
    let recentWorkdays = 0;
    for (let offset = 0; offset < 7; offset++) {
        const dt = new Date(todayUtc);
        dt.setUTCDate(dt.getUTCDate() - offset);
        if (dt.getUTCDay() !== 0 && covered.has(dt.toISOString().slice(0, 10))) {
            recentWorkdays += 1;
        }
    }
    return recentWorkdays;
}

function needsRecentCoverage(data, refTodayStr = getMskTodayStr()) {
    const covered = collectCoveredDates(data);
    const [y, m, d] = refTodayStr.split('-').map(Number);
    const todayUtc = new Date(Date.UTC(y, m - 1, d));
    if (todayUtc.getUTCDay() !== 0 && !covered.has(refTodayStr)) {
        return true;
    }
    return countRecentWorkdays(data, refTodayStr) < 3;
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
        const xlsxItems = items.filter(it => (it.name || '').toLowerCase().endsWith('.xlsx') && it.file);

        if (!xlsxItems.length) {
            return res.status(404).json({ error: 'No .xlsx files found on Yandex Disk' });
        }

        // Sort latest first
        xlsxItems.sort((a, b) => new Date(b.modified || b.created).getTime() - new Date(a.modified || a.created).getTime());
        const latest = xlsxItems[0];
        const latestSig = `${latest.name}|${latest.modified || latest.created || ''}`;
        const todayStr = getMskTodayStr();

        const fileResp = await fetch(latest.file, { headers: { 'User-Agent': 'Mozilla/5.0' } });
        if (!fileResp.ok) {
            return res.status(500).json({ error: `Failed to download ${latest.name}` });
        }
        const buffer = Buffer.from(await fileResp.arrayBuffer());
        const latestParsed = parseWorkbook(buffer, latest.name);
        let merged = pruneScheduleData(latestParsed, HISTORY_RETENTION_DAYS, todayStr);

        const [ty, tm, td] = todayStr.split('-').map(Number);
        const isTodaySunday = new Date(Date.UTC(ty, tm - 1, td)).getUTCDay() === 0;
        const latestMissesToday = !isTodaySunday && !collectCoveredDates(latestParsed).has(todayStr);

        // Backfill past 7 days + today from recent Yandex.Disk files when latest starts from tomorrow or only covers partial days
        if ((needsRecentCoverage(merged, todayStr) || latestMissesToday) && xlsxItems.length > 1) {
            const olderParsedList = [];
            for (const candidate of xlsxItems.slice(1, 4)) {
                try {
                    const secResp = await fetch(candidate.file, { headers: { 'User-Agent': 'Mozilla/5.0' } });
                    if (secResp.ok) {
                        const secBuf = Buffer.from(await secResp.arrayBuffer());
                        const secParsed = parseWorkbook(secBuf, candidate.name);
                        if (secParsed && secParsed.groups && secParsed.groups.length > 0) {
                            olderParsedList.push(secParsed);
                            let rebuilt = null;
                            for (let idx = olderParsedList.length - 1; idx >= 0; idx--) {
                                rebuilt = mergeScheduleData(rebuilt, olderParsedList[idx], HISTORY_RETENTION_DAYS, todayStr);
                            }
                            merged = mergeScheduleData(rebuilt, latestParsed, HISTORY_RETENTION_DAYS, todayStr);
                            const curMissesToday = !isTodaySunday && !collectCoveredDates(merged).has(todayStr);
                            if (!curMissesToday && countRecentWorkdays(merged, todayStr) >= 5) {
                                break;
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
            sourceSignature: latestSig,
            yandexDiskUrl: PUBLIC_KEY,
            groups: merged.groups,
            coveredDates: merged.coveredDates,
            schedules: merged.schedules
        });
    } catch (e) {
        return res.status(500).json({ error: e.message || String(e) });
    }
};
