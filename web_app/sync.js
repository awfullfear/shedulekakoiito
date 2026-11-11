/**
 * Standalone schedule syncer for Mabiu Schedule.
 * Preserves a rolling 7-day history when newer files only cover future days.
 * Can be run manually or by GitHub Actions on a schedule.
 */
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const PUBLIC_KEY = 'https://disk.yandex.ru/d/7XafcYMVe1ec7g';
const API_URL = `https://cloud-api.yandex.net/v1/disk/public/resources?public_key=${encodeURIComponent(PUBLIC_KEY)}&limit=100`;
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
    prunedSchedules[g] = { days: keptDays };
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
      mergedSchedules[g] = { days: sortedDays };
    }
  }

  const mergedCovered = Array.from(new Set([...baseCovered, ...incomingCovered])).filter(d => d >= cutoffStr).sort();
  return {
    ...baseData,
    ...incomingData,
    lastUpdated: Math.max(baseData.lastUpdated || 0, incomingData.lastUpdated || 0),
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

function parseWorkbookBuffer(buffer, fileName) {
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

  // Detect dates
  const detectedDates = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    let foundDate = null;
    let foundDow = null;
    let foundDowText = '';

    for (let c = 0; c <= Math.min(7, range.e.c); c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v === undefined || cell.v === null) continue;

      if (!foundDate) {
        if (cell.t === 'd' || cell.v instanceof Date) {
          const d = cell.v instanceof Date ? cell.v : new Date(cell.w || cell.v);
          foundDate = d.toISOString().split('T')[0];
        } else {
          const text = String(cell.w || cell.v || '').trim();
          const m1 = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
          if (m1) foundDate = `${m1[1]}-${m1[2].padStart(2, '0')}-${m1[3].padStart(2, '0')}`;
          const m2 = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})/);
          if (m2) {
            let y = parseInt(m2[3], 10);
            if (y < 100) y += 2000;
            foundDate = `${y}-${m2[2].padStart(2, '0')}-${m2[1].padStart(2, '0')}`;
          }
        }
      }

      if (!foundDow) {
        const text = String(cell.w || cell.v || '').trim().toLowerCase();
        if (text === 'пн' || text.startsWith('понед')) foundDow = 1;
        else if (text === 'вт' || text.startsWith('втор')) foundDow = 2;
        else if (text === 'ср' || text.startsWith('сред')) foundDow = 3;
        else if (text === 'чт' || text.startsWith('четв')) foundDow = 4;
        else if (text === 'пт' || text.startsWith('пятн')) foundDow = 5;
        else if (text === 'сб' || text.startsWith('субб')) foundDow = 6;
        else if (text === 'вс' || text.startsWith('воскр')) foundDow = 7;
        if (foundDow) foundDowText = text;
      }
    }

    if (foundDate || foundDow) {
      detectedDates.push({ r, date: foundDate, dow: foundDow, dowText: foundDowText });
    }
  }

  // Row to date mapping with dow correction
  const rowDateMap = {};
  let lastGoodDateObj = null;

  for (let i = 0; i < detectedDates.length; i++) {
    const cur = detectedDates[i];
    let resolvedDateStr = cur.date;

    if (resolvedDateStr && cur.dow) {
      const [y, m, d] = resolvedDateStr.split('-').map(Number);
      const testD = new Date(Date.UTC(y, m - 1, d));
      const actualDow = testD.getUTCDay() === 0 ? 7 : testD.getUTCDay();

      if (actualDow !== cur.dow) {
        if (lastGoodDateObj) {
          const expectedD = new Date(lastGoodDateObj);
          let diffDays = cur.dow - (expectedD.getUTCDay() === 0 ? 7 : expectedD.getUTCDay());
          if (diffDays <= 0) diffDays += 7;
          expectedD.setUTCDate(expectedD.getUTCDate() + diffDays);
          resolvedDateStr = expectedD.toISOString().split('T')[0];
        } else {
          let diff = cur.dow - actualDow;
          testD.setUTCDate(testD.getUTCDate() + diff);
          resolvedDateStr = testD.toISOString().split('T')[0];
        }
      }
    } else if (!resolvedDateStr && cur.dow && lastGoodDateObj) {
      const expectedD = new Date(lastGoodDateObj);
      let diffDays = cur.dow - (expectedD.getUTCDay() === 0 ? 7 : expectedD.getUTCDay());
      if (diffDays <= 0) diffDays += 7;
      expectedD.setUTCDate(expectedD.getUTCDate() + diffDays);
      resolvedDateStr = expectedD.toISOString().split('T')[0];
    }

    if (resolvedDateStr) {
      const [y, m, d] = resolvedDateStr.split('-').map(Number);
      lastGoodDateObj = new Date(Date.UTC(y, m - 1, d));
    }

    const nextR = (i + 1 < detectedDates.length) ? detectedDates[i + 1].r : (range.e.r + 1);
    for (let row = cur.r; row < nextR; row++) {
      if (resolvedDateStr) rowDateMap[row] = resolvedDateStr;
    }
  }

  // Parse schedules
  const schedules = {};
  allGroups.forEach(g => { schedules[g] = { days: {} }; });

  const rawRowsByDate = {};
  for (let r = 2; r <= range.e.r; r++) {
    const dateStr = rowDateMap[r];
    if (!dateStr) continue;

    let timeStr = '';
    for (let c = 0; c <= Math.min(7, range.e.c); c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.v) {
        const text = String(cell.w || cell.v).trim();
        const m = text.match(/(\d{1,2}[:.]\d{2})\s*[-–—]\s*(\d{1,2}[:.]\d{2})/);
        if (m) {
          timeStr = `${m[1]} – ${m[2]}`;
          break;
        }
      }
    }

    if (!timeStr) continue;

    const cleanT = timeStr.replace(/:/g, '.').replace(/\s+/g, '');
    let pairNum = '';
    if (cleanT.startsWith('9.00') || cleanT.startsWith('09.00')) pairNum = '1 пара';
    else if (cleanT.startsWith('10.30')) pairNum = '2 пара';
    else if (cleanT.startsWith('12.05')) pairNum = '3 пара';
    else if (cleanT.startsWith('13.45')) pairNum = '4 пара';
    else if (cleanT.startsWith('15.20')) pairNum = '5 пара';
    else if (cleanT.startsWith('16.55')) pairNum = '6 пара';
    else if (cleanT.startsWith('18.30')) pairNum = '7 пара';
    else if (cleanT.startsWith('20.00')) pairNum = '8 пара';

    if (!rawRowsByDate[dateStr]) rawRowsByDate[dateStr] = [];

    const cellsInRow = {};
    for (const [col, group] of Object.entries(groupsByCol)) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c: parseInt(col, 10) })];
      const val = cell && cell.v ? String(cell.w || cell.v).trim() : '';
      if (val) cellsInRow[group] = val;
    }

    rawRowsByDate[dateStr].push({ pairNum, timeStr, cellsInRow });
  }

  // Populate lessons & joints
  const typeRegex = /\(\s*(практическ(?:ое|ие)\s+заняти[ея]|практ\.?\s*заняти[ея]|практик[аи]|практ|пр\.?\s*з|пр|лекци[яи]|лекц|лек|семинарск(?:ое|ие)\s+заняти[ея]|семинары|семинар|сем|лабораторн(?:ая|ые)\s+работ[аы]|лабораторн(?:ая|ые)|лаб|дифференцированн(?:ый)?\s*зач[её]т|диф\.?\s*зач[её]т|зач[её]т|экзамен|экз|консультаци[яи]|конс)\.?\s*\)/gi;

  for (const [dateStr, entries] of Object.entries(rawRowsByDate)) {
    for (const entry of entries) {
      const contentToGroups = {};
      for (const [group, content] of Object.entries(entry.cellsInRow)) {
        const norm = content.toLowerCase().replace(/[^a-zа-яё0-9]/g, '');
        if (!norm) continue;
        if (!contentToGroups[norm]) contentToGroups[norm] = [];
        contentToGroups[norm].push(group);
      }

      for (const [group, content] of Object.entries(entry.cellsInRow)) {
        if (!content) continue;
        const norm = content.toLowerCase().replace(/[^a-zа-яё0-9]/g, '');
        const joint = (contentToGroups[norm] || []).filter(g => g !== group);

        let text = String(content).trim();
        let room = '';
        let teacher = '';
        let lessonType = '';

        const roomMatch = text.match(/(?:,\s*)?(?:(?:Ауд\.?|АУд\.?)\s*[\w\d/]*\s*(?:\([^)]+\))?|(?:Актовый\s*зал|Конференц[- ]*зал|Спорт\.?\s*зал|Он[- ]*лайн|ДО)(?:\s*\([^)]+\))?|\((?:Конференц[- ]*зал|Актовый\s*зал)\))\s*$/i);
        if (roomMatch) {
          room = roomMatch[0].replace(/^,\s*/, '').trim();
          text = text.substring(0, text.length - roomMatch[0].length).trim();
        }

        text = text.replace(typeRegex, (_, rawType) => {
          if (!lessonType) {
            const t = String(rawType).trim().toLowerCase();
            if (t.startsWith('лек')) lessonType = 'Лекция';
            else if (t.startsWith('пр')) lessonType = 'Практическое занятие';
            else if (t.startsWith('сем')) lessonType = 'Семинар';
            else if (t.startsWith('лаб')) lessonType = 'Лабораторная';
            else if (t.includes('зач')) lessonType = 'Зачёт';
            else if (t.startsWith('экз')) lessonType = 'Экзамен';
            else if (t.startsWith('конс')) lessonType = 'Консультация';
            else lessonType = String(rawType).trim();
          }
          return ' ';
        });

        text = text
          .replace(/\s+/g, ' ')
          .replace(/\s+([,.;:!?])/g, '$1')
          .replace(/^[,.;:\s-]+|[,;\s-]+$/g, '')
          .trim();

        const teacherMatch = text.match(/([А-ЯЁа-яё][а-яё]{1,}\s+[А-ЯЁ]\.+\s*[А-ЯЁ]\.*)\s*$/);
        if (teacherMatch) {
          teacher = teacherMatch[1].trim();
          text = text.substring(0, text.length - teacherMatch[0].length).trim();
        }

        const cleanSubject = text
          .replace(/\s+/g, ' ')
          .replace(/\s+([,.;:!?])/g, '$1')
          .replace(/^[,.;:\s-]+|[,;\s-]+$/g, '')
          .trim() || content;

        if (!schedules[group].days[dateStr]) schedules[group].days[dateStr] = [];
        schedules[group].days[dateStr].push({
          pairNumber: entry.pairNum,
          time: entry.timeStr,
          subject: content,
          cleanSubject,
          lessonType,
          room,
          teacher,
          jointGroups: joint
        });
      }
    }
  }

  return {
    lastUpdated: Date.now(),
    fileName,
    yandexDiskUrl: PUBLIC_KEY,
    groups: allGroups,
    coveredDates: Object.keys(rawRowsByDate).sort(),
    schedules
  };
}

async function sync() {
  const fetch = global.fetch || require('node-fetch');
  console.log('Connecting to Yandex.Disk API...');

  const resp = await fetch(API_URL, { headers: { 'User-Agent': 'Mozilla/5.0' } });
  if (!resp.ok) {
    throw new Error(`Yandex API status ${resp.status}`);
  }

  const data = await resp.json();
  const items = (data._embedded && data._embedded.items) || [];
  const xlsxFiles = items.filter(it => (it.name || '').toLowerCase().endsWith('.xlsx') && it.file);

  if (!xlsxFiles.length) {
    console.log('No .xlsx files found on disk.');
    return;
  }

  xlsxFiles.sort((a, b) => new Date(b.modified || b.created).getTime() - new Date(a.modified || a.created).getTime());
  const latest = xlsxFiles[0];
  const latestSig = `${latest.name}|${latest.modified || latest.created || ''}`;
  console.log(`Latest file on disk: ${latest.name} (modified: ${latest.modified})`);

  const todayStr = getMskTodayStr();
  const cutoffStr = shiftDateStr(todayStr, -HISTORY_RETENTION_DAYS);
  const outPath = path.join(__dirname, 'data', 'schedule.json');

  let existingData = null;
  let existingRawPruned = false;
  if (fs.existsSync(outPath)) {
    try {
      const parsed = JSON.parse(fs.readFileSync(outPath, 'utf8'));
      if (parsed && parsed.schedules) {
        const rawDates = collectCoveredDates(parsed);
        existingRawPruned = Array.from(rawDates).some(ds => ds < cutoffStr);
        existingData = pruneScheduleData(parsed, HISTORY_RETENTION_DAYS, todayStr);
      }
    } catch (_) {}
  }

  if (
    existingData &&
    existingData.sourceSignature === latestSig &&
    !needsRecentCoverage(existingData, todayStr)
  ) {
    if (!existingRawPruned) {
      console.log('Schedule is already up to date with full 7-day rolling history.');
      return;
    }
    fs.mkdirSync(path.dirname(outPath), { recursive: true });
    fs.writeFileSync(outPath, JSON.stringify(existingData, null, 2), 'utf8');
    console.log(`Pruned expired history (< ${cutoffStr}) in ${outPath}.`);
    return;
  }

  const fileResp = await fetch(latest.file, { headers: { 'User-Agent': 'Mozilla/5.0' } });
  if (!fileResp.ok) {
    throw new Error(`Failed to download ${latest.name}`);
  }

  const buffer = Buffer.from(await fileResp.arrayBuffer());
  const latestParsed = parseWorkbookBuffer(buffer, latest.name);
  let merged = mergeScheduleData(existingData, latestParsed, HISTORY_RETENTION_DAYS, todayStr);

  const [ty, tm, td] = todayStr.split('-').map(Number);
  const isTodaySunday = new Date(Date.UTC(ty, tm - 1, td)).getUTCDay() === 0;
  const latestMissesToday = !isTodaySunday && !collectCoveredDates(latestParsed).has(todayStr);

  if ((needsRecentCoverage(merged, todayStr) || latestMissesToday) && xlsxFiles.length > 1) {
    console.log('Latest file lacks today or recent week; backfilling from recent Yandex.Disk files...');
    const olderParsedList = [];
    for (const candidate of xlsxFiles.slice(1, 4)) {
      try {
        const fbResp = await fetch(candidate.file, { headers: { 'User-Agent': 'Mozilla/5.0' } });
        if (fbResp.ok) {
          const fbBuf = Buffer.from(await fbResp.arrayBuffer());
          const fbParsed = parseWorkbookBuffer(fbBuf, candidate.name);
          if (fbParsed && fbParsed.groups && fbParsed.groups.length > 0) {
            olderParsedList.push(fbParsed);
            let rebuilt = existingData;
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
      } catch (err) {
        console.warn(`Skipping fallback ${candidate.name}:`, err.message || err);
      }
    }
  }

  const finalResult = {
    lastUpdated: Date.now(),
    fileName: latest.name,
    sourceSignature: latestSig,
    yandexDiskUrl: PUBLIC_KEY,
    groups: merged.groups,
    coveredDates: merged.coveredDates,
    schedules: merged.schedules
  };

  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, JSON.stringify(finalResult, null, 2), 'utf8');
  console.log(`Successfully updated ${outPath} with ${latest.name} (retained ${merged.coveredDates.length} dates from ${cutoffStr})`);
}

sync().catch(err => {
  console.error('Sync failed:', err);
  process.exit(1);
});
