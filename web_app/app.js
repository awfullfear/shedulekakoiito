/**
 * Mabiu Schedule Web App
 * Features:
 * - 26-cohort filter
 * - Live status countdown ticker (ongoing lesson, break, tomorrow start time, weekend)
 * - Swipeable day carousel with "Сегодня" jump
 * - Joint groups ("Совместно с: ...")
 * - Client-side .xlsx manual upload & parsing with auto-date correction
 * - Auto-sync with Yandex.Disk (/api/schedule or CORS fallback)
 * - PWA offline support
 */

// Configuration & State
const CONFIG = {
  yandexPublicKey: 'https://disk.yandex.ru/d/7XafcYMVe1ec7g',
  yandexKeyId: '7XafcYMVe1ec7g',
  corsProxy: 'https://api.allorigins.win/raw?url='
};

const LESSON_TYPE_REGEX = /\(\s*(практическ(?:ое|ие)\s+заняти[ея]|практ\.?\s*заняти[ея]|практик[аи]|практ|пр\.?\s*з|пр|лекци[яи]|лекц|лек|семинарск(?:ое|ие)\s+заняти[ея]|семинары|семинар|сем|лабораторн(?:ая|ые)\s+работ[аы]|лабораторн(?:ая|ые)|лаб|дифференцированн(?:ый)?\s*зач[её]т|диф\.?\s*зач[её]т|зач[её]т|экзамен|экз|консультаци[яи]|конс)\.?\s*\)/gi;
const TEACHER_END_REGEX = /([А-ЯЁа-яё][а-яё]{1,}\s+[А-ЯЁ]\.+\s*[А-ЯЁ]\.*)\s*$/;

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

function cleanSubjectTextClient(text) {
  return String(text || '')
    .replace(/\s+/g, ' ')
    .replace(/\s+([,.;:!?])/g, '$1')
    .replace(/^[,.;:\s-]+|[,;\s-]+$/g, '')
    .trim();
}

function normalizeLessonRecord(lesson) {
  if (!lesson || typeof lesson !== 'object') return lesson;
  let lessonType = normalizeLessonTypeLabel(lesson.lessonType || '');
  let clean = String(lesson.cleanSubject || lesson.subject || '').trim();

  clean = clean.replace(LESSON_TYPE_REGEX, (_, rawType) => {
    if (!lessonType) lessonType = normalizeLessonTypeLabel(rawType);
    return ' ';
  });

  if (!lessonType && lesson.subject) {
    String(lesson.subject).replace(LESSON_TYPE_REGEX, (_, rawType) => {
      if (!lessonType) lessonType = normalizeLessonTypeLabel(rawType);
      return ' ';
    });
  }

  clean = cleanSubjectTextClient(clean);

  let teacher = lesson.teacher || '';
  if (!teacher) {
    const tm = clean.match(TEACHER_END_REGEX);
    if (tm) {
      teacher = tm[1].trim();
      clean = cleanSubjectTextClient(clean.substring(0, clean.length - tm[0].length));
    }
  }

  return {
    ...lesson,
    cleanSubject: clean || String(lesson.subject || '').trim(),
    lessonType,
    teacher
  };
}

function normalizeScheduleData(data) {
  if (!data || typeof data !== 'object' || !data.schedules) return data;
  for (const grp of Object.keys(data.schedules)) {
    const days = data.schedules[grp]?.days;
    if (!days || typeof days !== 'object') continue;
    for (const dateStr of Object.keys(days)) {
      if (Array.isArray(days[dateStr])) {
        days[dateStr] = days[dateStr].map(normalizeLessonRecord);
      }
    }
  }
  return data;
}

const SCHEDULE_HISTORY_DAYS = 7;

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

function pruneScheduleData(data, retainDays = SCHEDULE_HISTORY_DAYS, refTodayStr = getTodayDateStr()) {
  if (!data || typeof data !== 'object') return data;
  const cutoffStr = shiftDateStr(refTodayStr, -retainDays);
  const prunedSchedules = {};
  const srcSchedules = data.schedules || {};

  for (const grp of Object.keys(srcSchedules)) {
    const gdata = srcSchedules[grp];
    if (!gdata || typeof gdata !== 'object') continue;
    const days = gdata.days || {};
    const keptDays = {};
    for (const d of Object.keys(days).sort()) {
      if (d >= cutoffStr) {
        keptDays[d] = days[d];
      }
    }
    const nextGroupObj = { ...gdata, days: keptDays };
    if (Array.isArray(gdata.availableDates)) {
      nextGroupObj.availableDates = Object.keys(keptDays).sort();
    }
    prunedSchedules[grp] = nextGroupObj;
  }

  const covered = Array.from(collectCoveredDates(data))
    .filter(d => d >= cutoffStr)
    .sort();

  return {
    ...data,
    coveredDates: covered,
    schedules: prunedSchedules
  };
}

function mergeScheduleData(baseData, incomingData, options = {}) {
  const retainDays = options.retainDays !== undefined ? options.retainDays : SCHEDULE_HISTORY_DAYS;
  const refTodayStr = options.todayStr || getTodayDateStr();
  const cutoffStr = shiftDateStr(refTodayStr, -retainDays);

  if (!incomingData || !incomingData.schedules || Object.keys(incomingData.schedules).length === 0) {
    return baseData ? pruneScheduleData(baseData, retainDays, refTodayStr) : incomingData;
  }
  if (!baseData || !baseData.schedules || Object.keys(baseData.schedules).length === 0) {
    return pruneScheduleData(incomingData, retainDays, refTodayStr);
  }

  const incomingCovered = collectCoveredDates(incomingData);
  const baseCovered = collectCoveredDates(baseData);

  const mergedSchedules = {};
  const baseScheds = baseData.schedules || {};
  const incScheds = incomingData.schedules || {};
  const incomingGroups = Array.isArray(incomingData.groups) ? [...incomingData.groups] : [];
  const mergedGroups = [...incomingGroups];

  if (Array.isArray(baseData.groups)) {
    for (const g of baseData.groups) {
      if (!mergedGroups.includes(g)) {
        const hasRetained = Object.keys(baseScheds[g]?.days || {}).some(d => d >= cutoffStr && !incomingCovered.has(d));
        if (hasRetained) mergedGroups.push(g);
      }
    }
  }

  const allGroupKeys = new Set([
    ...mergedGroups,
    ...Object.keys(baseScheds),
    ...Object.keys(incScheds)
  ]);

  for (const g of allGroupKeys) {
    const mergedDays = {};
    const baseDays = baseScheds[g]?.days || {};
    const incDays = incScheds[g]?.days || {};

    for (const d of Object.keys(baseDays)) {
      if (d >= cutoffStr && !incomingCovered.has(d)) {
        mergedDays[d] = baseDays[d];
      }
    }
    for (const d of Object.keys(incDays)) {
      if (d >= cutoffStr) {
        mergedDays[d] = incDays[d];
      }
    }

    const sortedDays = {};
    for (const d of Object.keys(mergedDays).sort()) {
      sortedDays[d] = mergedDays[d];
    }
    if (incomingGroups.includes(g) || Object.keys(sortedDays).length > 0) {
      mergedSchedules[g] = { days: sortedDays };
    }
  }

  const mergedCovered = Array.from(new Set([...baseCovered, ...incomingCovered]))
    .filter(d => d >= cutoffStr)
    .sort();

  const mergedSourceSignature = incomingData.sourceSignature !== undefined
    ? incomingData.sourceSignature
    : baseData.sourceSignature;

  return {
    ...baseData,
    ...incomingData,
    lastUpdated: Math.max(baseData.lastUpdated || 0, incomingData.lastUpdated || 0),
    groups: mergedGroups,
    coveredDates: mergedCovered,
    sourceSignature: mergedSourceSignature,
    schedules: mergedSchedules
  };
}

function countRecentWorkdays(data, refTodayStr = getTodayDateStr()) {
  const covered = collectCoveredDates(data);
  const [y, m, d] = String(refTodayStr).split('-').map(Number);
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

function needsRecentCoverage(data, refTodayStr = getTodayDateStr()) {
  const covered = collectCoveredDates(data);
  const [y, m, d] = String(refTodayStr).split('-').map(Number);
  const todayUtc = new Date(Date.UTC(y, m - 1, d));
  const todayDow = todayUtc.getUTCDay();
  if (todayDow !== 0 && !covered.has(refTodayStr)) {
    return true;
  }
  return countRecentWorkdays(data, refTodayStr) < 3;
}

async function fetchWithTimeout(url, options = {}, timeoutMs = 10000) {
  if (typeof AbortController === 'undefined') {
    return fetch(url, options);
  }
  const controller = new AbortController();
  const timerId = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } finally {
    clearTimeout(timerId);
  }
}

function saveScheduleDataSafe(data) {
  const trySave = (obj) => {
    localStorage.setItem('mabiu_schedule_data', JSON.stringify(obj));
    return obj;
  };
  try {
    return trySave(data);
  } catch (err) {
    for (const days of [3, 1, 0]) {
      try {
        const trimmed = pruneScheduleData(data, days);
        return trySave(trimmed);
      } catch (_) {}
    }
    console.warn('Unable to cache schedule in localStorage:', err);
    return data;
  }
}

const THEMES = [
  {
    id: 'blue',
    name: 'Океан',
    emoji: '🌊',
    color: '#2563eb',
    desc: 'Глубокий океан и морской бриз',
    floatingEmojis: ['🌊', '🫧', '🐚', '🐬', '🐳', '💧'],
    burstColors: ['#3b82f6', '#38bdf8', '#93c5fd', '#2563eb'],
    sparkEmoji: '🫧'
  },
  {
    id: 'emerald',
    name: 'Изумруд',
    emoji: '🌿',
    color: '#059669',
    desc: 'Свежая лесная листва и природа',
    floatingEmojis: ['🌿', '🍃', '🍀', '🌱', '🪴', '🪷'],
    burstColors: ['#10b981', '#34d399', '#6ee7b7', '#059669'],
    sparkEmoji: '🍃'
  },
  {
    id: 'violet',
    name: 'Аметист',
    emoji: '🔮',
    color: '#7c3aed',
    desc: 'Ночной космос, звёзды и кристаллы',
    floatingEmojis: ['🔮', '✨', '🪐', '🌙', '⭐', '💎'],
    burstColors: ['#8b5cf6', '#c084fc', '#a78bfa', '#7c3aed'],
    sparkEmoji: '✨'
  },
  {
    id: 'rose',
    name: 'Сакура',
    emoji: '🌸',
    color: '#e11d48',
    desc: 'Весенние лепестки и бабочки',
    floatingEmojis: ['🌸', '🌺', '🌷', '🦋', '💮', '🎀'],
    burstColors: ['#f43f5e', '#fb7185', '#fda4af', '#e11d48'],
    sparkEmoji: '🌸'
  },
  {
    id: 'amber',
    name: 'Осень',
    emoji: '🍁',
    color: '#d97706',
    desc: 'Тёплый янтарь, листопад и уют',
    floatingEmojis: ['🍁', '🍂', '☕', '🧣', '🌰', '✨'],
    burstColors: ['#f59e0b', '#fbbf24', '#ea580c', '#d97706'],
    sparkEmoji: '🍂'
  },
  {
    id: 'arctic',
    name: 'Арктика',
    emoji: '❄️',
    color: '#0284c7',
    desc: 'Снежинки, лёд и зимняя свежесть',
    floatingEmojis: ['❄️', '🧊', '⛄', '🌬️', '💎', '✨'],
    burstColors: ['#0ea5e9', '#38bdf8', '#bae6fd', '#0284c7'],
    sparkEmoji: '❄️'
  }
];

function getValidThemeId(themeId) {
  return THEMES.some(t => t.id === themeId) ? themeId : 'blue';
}

const storedGroupRaw = (localStorage.getItem('mabiu_group') || '').trim();
let scheduleData = null;
let currentGroup = storedGroupRaw || null;
let currentTheme = getValidThemeId(localStorage.getItem('mabiu_theme'));
let emojisEnabled = localStorage.getItem('mabiu_emojis_enabled') !== '0';
let selectedDateStr = getTodayDateStr();
let tickerTimerId = null;

document.documentElement.setAttribute('data-emojis', emojisEnabled ? 'on' : 'off');

// Floating emojis runtime state
let activeFloatingEmojis = [];
const emojiSpawnTimers = new Set();
let emojiGeneration = 0;
let appliedDecorationTheme = null;

// DOM Elements
const currentGroupLabel = document.getElementById('currentGroupLabel');
const groupSelectorBtn = document.getElementById('groupSelectorBtn');
const refreshBtn = document.getElementById('refreshBtn');
const settingsBtn = document.getElementById('settingsBtn');
const themeBtn = document.getElementById('themeBtn');
const headerThemeEmoji = document.getElementById('headerThemeEmoji');
const statusThemeEmoji = document.getElementById('statusThemeEmoji');
const themeColorMeta = document.getElementById('themeColorMeta');
const themeDecorations = document.getElementById('themeDecorations');
const themeGrid = document.getElementById('themeGrid');
const emojisToggle = document.getElementById('emojisToggle');

const themeModal = document.getElementById('themeModal');
const closeThemeModalBtn = document.getElementById('closeThemeModalBtn');
const themeSelectList = document.getElementById('themeSelectList');

const statusBanner = document.getElementById('statusBanner');
const statusIconBox = document.getElementById('statusIconBox');
const statusBadge = document.getElementById('statusBadge');
const statusTitle = document.getElementById('statusTitle');
const statusSubtitle = document.getElementById('statusSubtitle');

const monthLabel = document.getElementById('monthLabel');
const todayJumpBtn = document.getElementById('todayJumpBtn');
const daysCarousel = document.getElementById('daysCarousel');

const dayHeadingText = document.getElementById('dayHeadingText');
const dayBadgeToday = document.getElementById('dayBadgeToday');
const lessonsList = document.getElementById('lessonsList');
const scheduleContainer = document.getElementById('scheduleContainer');

const groupModal = document.getElementById('groupModal');
const closeGroupModalBtn = document.getElementById('closeGroupModalBtn');
const groupSearchInput = document.getElementById('groupSearchInput');
const groupList = document.getElementById('groupList');

const settingsModal = document.getElementById('settingsModal');
const closeSettingsModalBtn = document.getElementById('closeSettingsModalBtn');
const dropzone = document.getElementById('dropzone');
const xlsxFileInput = document.getElementById('xlsxFileInput');
const syncNowBtn = document.getElementById('syncNowBtn');
const resetScheduleBtn = document.getElementById('resetScheduleBtn');
const currentFileNameLabel = document.getElementById('currentFileName');
const lastUpdatedLabel = document.getElementById('lastUpdatedLabel');
const diskSourceLabel = document.getElementById('diskSourceLabel');
const toastEl = document.getElementById('toast');

// --- Helper Functions ---

function getTodayDateStr() {
  const d = new Date();
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function formatDateToRu(dateStr) {
  const [y, m, d] = dateStr.split('-').map(Number);
  const date = new Date(y, m - 1, d);
  const daysOfWeek = ['Воскресенье', 'Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота'];
  const months = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
  return `${daysOfWeek[date.getDay()]}, ${d} ${months[m - 1]}`;
}

function getMonthNameRu(monthIndex, year) {
  const months = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
  return `${months[monthIndex]} ${year}`;
}

function showToast(message, duration = 3000) {
  if (!toastEl) return;
  toastEl.textContent = message;
  toastEl.classList.add('show');
  setTimeout(() => {
    toastEl.classList.remove('show');
  }, duration);
}

function timeStringToMinutes(timeStr) {
  if (!timeStr) return 0;
  const parts = timeStr.replace('.', ':').split(':');
  return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
}

function formatMinutesToTime(totalMin) {
  const h = Math.floor(totalMin / 60);
  const m = totalMin % 60;
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}

function formatDuration(diffMs) {
  if (diffMs <= 0) return '00:00';
  const totalSeconds = Math.floor(diffMs / 1000);
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;

  if (hours > 0) {
    return `${hours} ч ${minutes} мин`;
  }
  return `${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
}

// --- Themes Management & Interactive Floating Emojis ---

function getThemeById(themeId) {
  return THEMES.find(t => t.id === themeId) || THEMES[0];
}

function getEmojiCandidateZones() {
  const w = window.innerWidth || 390;
  let leftMax = 12;
  let rightMin = 88;
  if (w >= 760) {
    const gutterPct = Math.floor((((w - 520) / 2) / w) * 100);
    leftMax = Math.min(26, Math.max(12, gutterPct - 2));
    rightMin = 100 - leftMax;
  }

  return [
    { xMin: rightMin, xMax: 95, yMin: 13, yMax: 21 },
    { xMin: 4, xMax: leftMax, yMin: 15, yMax: 24 },
    { xMin: rightMin, xMax: 96, yMin: 27, yMax: 38 },
    { xMin: 4, xMax: leftMax, yMin: 33, yMax: 45 },
    { xMin: rightMin, xMax: 95, yMin: 44, yMax: 56 },
    { xMin: 4, xMax: leftMax, yMin: 52, yMax: 64 },
    { xMin: rightMin, xMax: 96, yMin: 63, yMax: 74 },
    { xMin: 5, xMax: leftMax, yMin: 70, yMax: 81 },
    { xMin: rightMin - 4, xMax: 94, yMin: 80, yMax: 90 },
    { xMin: 5, xMax: leftMax + 4, yMin: 82, yMax: 91 }
  ];
}

function pickFloatingSlot(avoidSlotIndex = -1, avoidPos = null) {
  const zones = getEmojiCandidateZones();
  const occupiedIndices = new Set(activeFloatingEmojis.map(item => item.slotIndex));

  let candidates = [];
  for (let i = 0; i < zones.length; i++) {
    if (i === avoidSlotIndex || occupiedIndices.has(i)) continue;
    const z = zones[i];
    const x = z.xMin + Math.random() * Math.max(2, z.xMax - z.xMin);
    const y = z.yMin + Math.random() * Math.max(2, z.yMax - z.yMin);

    let tooClose = false;
    for (const active of activeFloatingEmojis) {
      const dx = x - active.x;
      const dy = y - active.y;
      if (Math.hypot(dx, dy) < 18) {
        tooClose = true;
        break;
      }
    }
    if (!tooClose && avoidPos) {
      if (Math.hypot(x - avoidPos.x, y - avoidPos.y) < 20) {
        tooClose = true;
      }
    }
    if (!tooClose) {
      candidates.push({ slotIndex: i, x, y });
    }
  }

  if (candidates.length === 0) {
    for (let i = 0; i < zones.length; i++) {
      if (i !== avoidSlotIndex && !occupiedIndices.has(i)) {
        const z = zones[i];
        candidates.push({
          slotIndex: i,
          x: (z.xMin + z.xMax) / 2,
          y: (z.yMin + z.yMax) / 2
        });
      }
    }
  }

  if (candidates.length === 0) {
    const fallbackIdx = (avoidSlotIndex + 3) % zones.length;
    const z = zones[fallbackIdx];
    return {
      slotIndex: fallbackIdx,
      x: (z.xMin + z.xMax) / 2,
      y: (z.yMin + z.yMax) / 2
    };
  }

  return candidates[Math.floor(Math.random() * candidates.length)];
}

function pickThemeEmojiGlyph(theme, avoidEmoji = null) {
  const pool = theme.floatingEmojis && theme.floatingEmojis.length > 0 ? theme.floatingEmojis : [theme.emoji];
  const activeGlyphs = new Set(activeFloatingEmojis.map(item => item.emoji));

  let fresh = pool.filter(e => e !== avoidEmoji && !activeGlyphs.has(e));
  if (fresh.length === 0) {
    fresh = pool.filter(e => e !== avoidEmoji);
  }
  if (fresh.length === 0) {
    fresh = pool;
  }
  return fresh[Math.floor(Math.random() * fresh.length)];
}

function spawnPopParticles(cx, cy, theme) {
  if (!themeDecorations) return;
  const colors = theme.burstColors || [theme.color];
  const count = 8;

  // Expanding ring
  const ring = document.createElement('div');
  ring.className = 'emoji-pop-ring';
  ring.style.left = `${cx}px`;
  ring.style.top = `${cy}px`;
  ring.style.setProperty('--ring-color', theme.color);
  themeDecorations.appendChild(ring);
  setTimeout(() => ring.remove(), 460);

  for (let i = 0; i < count; i++) {
    const angle = (Math.PI * 2 * i) / count + (Math.random() - 0.5) * 0.4;
    const distance = 26 + Math.random() * 28;
    const tx = Math.cos(angle) * distance;
    const ty = Math.sin(angle) * distance;
    const p = document.createElement('div');
    const isSpark = i % 3 === 0;

    p.className = `emoji-burst-particle${isSpark ? ' spark' : ''}`;
    p.style.left = `${cx}px`;
    p.style.top = `${cy}px`;
    p.style.setProperty('--tx', `${tx.toFixed(1)}px`);
    p.style.setProperty('--ty', `${ty.toFixed(1)}px`);
    p.style.setProperty('--trot', `${Math.round((Math.random() - 0.5) * 180)}deg`);

    if (isSpark) {
      p.textContent = theme.sparkEmoji || '✨';
    } else {
      const size = 6 + Math.random() * 5;
      p.style.width = `${size.toFixed(1)}px`;
      p.style.height = `${size.toFixed(1)}px`;
      p.style.backgroundColor = colors[i % colors.length];
    }

    themeDecorations.appendChild(p);
    setTimeout(() => p.remove(), 520);
  }
}

function popFloatingEmoji(item, theme, gen) {
  if (!item || !item.el || item.el.dataset.popping === '1') return;
  item.el.dataset.popping = '1';

  const rect = item.el.getBoundingClientRect();
  const cx = rect.left + rect.width / 2;
  const cy = rect.top + rect.height / 2;

  item.el.classList.remove('spawning');
  item.el.classList.add('popping');
  spawnPopParticles(cx, cy, theme);

  activeFloatingEmojis = activeFloatingEmojis.filter(a => a !== item);

  setTimeout(() => {
    if (item.el && item.el.parentNode) {
      item.el.remove();
    }
  }, 270);

  // Schedule respawn in a different spot with another theme emoji
  const respawnDelay = 1100 + Math.random() * 1400;
  const timerId = setTimeout(() => {
    emojiSpawnTimers.delete(timerId);
    if (gen !== emojiGeneration) return;
    spawnSingleFloatingEmoji(theme, gen, {
      avoidSlotIndex: item.slotIndex,
      avoidPos: { x: item.x, y: item.y },
      avoidEmoji: item.emoji
    });
  }, respawnDelay);
  emojiSpawnTimers.add(timerId);
}

function spawnSingleFloatingEmoji(theme, gen, opts = {}) {
  if (!themeDecorations || !emojisEnabled || gen !== emojiGeneration) return;

  const slot = pickFloatingSlot(
    opts.avoidSlotIndex !== undefined ? opts.avoidSlotIndex : -1,
    opts.avoidPos || null
  );
  const emojiChar = pickThemeEmojiGlyph(theme, opts.avoidEmoji || null);

  const el = document.createElement('div');
  el.className = 'floating-emoji spawning';
  el.style.left = `${slot.x.toFixed(2)}%`;
  el.style.top = `${slot.y.toFixed(2)}%`;

  const glyph = document.createElement('span');
  glyph.className = 'floating-emoji-glyph';
  glyph.textContent = emojiChar;

  const fontSize = Math.round(23 + Math.random() * 7);
  const floatDur = (4.2 + Math.random() * 2.4).toFixed(2);
  const floatDelay = (-Math.random() * 3).toFixed(2);
  const floatX = `${Math.round((Math.random() - 0.5) * 14)}px`;
  const floatY = `${Math.round(-8 - Math.random() * 8)}px`;
  const floatRot = `${Math.round((Math.random() - 0.5) * 16)}deg`;
  const opacity = (0.64 + Math.random() * 0.22).toFixed(2);

  glyph.style.setProperty('--emoji-size', `${fontSize}px`);
  glyph.style.setProperty('--float-dur', `${floatDur}s`);
  glyph.style.setProperty('--float-delay', `${floatDelay}s`);
  glyph.style.setProperty('--float-x', floatX);
  glyph.style.setProperty('--float-y', floatY);
  glyph.style.setProperty('--float-rot', floatRot);
  glyph.style.setProperty('--emoji-opacity', opacity);

  el.appendChild(glyph);

  const item = {
    el,
    slotIndex: slot.slotIndex,
    x: slot.x,
    y: slot.y,
    emoji: emojiChar
  };

  const cleanSpawningClass = () => el.classList.remove('spawning');
  el.addEventListener('animationend', cleanSpawningClass, { once: true });
  setTimeout(cleanSpawningClass, 480);

  el.addEventListener('click', (e) => {
    e.stopPropagation();
    popFloatingEmoji(item, theme, gen);
  });

  activeFloatingEmojis.push(item);
  themeDecorations.appendChild(el);
}

function clearFloatingEmojis() {
  emojiGeneration += 1;
  appliedDecorationTheme = null;
  emojiSpawnTimers.forEach((id) => clearTimeout(id));
  emojiSpawnTimers.clear();
  activeFloatingEmojis = [];
  if (themeDecorations) {
    themeDecorations.innerHTML = '';
  }
}

function renderFloatingEmojisForTheme(theme) {
  clearFloatingEmojis();
  if (!themeDecorations || !emojisEnabled) return;
  appliedDecorationTheme = theme.id;
  const gen = emojiGeneration;
  const targetCount = (window.innerWidth || 390) >= 600 ? 5 : 4;

  for (let i = 0; i < targetCount; i++) {
    const tId = setTimeout(() => {
      emojiSpawnTimers.delete(tId);
      if (gen !== emojiGeneration || !emojisEnabled) return;
      spawnSingleFloatingEmoji(theme, gen);
    }, i * 90);
    emojiSpawnTimers.add(tId);
  }
}

function setEmojisEnabled(enabled, notify = true) {
  emojisEnabled = Boolean(enabled);
  localStorage.setItem('mabiu_emojis_enabled', emojisEnabled ? '1' : '0');
  document.documentElement.setAttribute('data-emojis', emojisEnabled ? 'on' : 'off');
  if (emojisToggle) {
    emojisToggle.checked = emojisEnabled;
  }

  if (emojisEnabled) {
    renderFloatingEmojisForTheme(getThemeById(currentTheme));
  } else {
    clearFloatingEmojis();
  }

  if (typeof renderLessons === 'function' && lessonsList) {
    renderLessons();
  }
  if (typeof updateTickerOnce === 'function') {
    updateTickerOnce();
  }

  if (notify) {
    showToast(emojisEnabled ? 'Смайлики включены' : 'Смайлики отключены');
  }
}

function renderThemeGrid() {
  if (!themeGrid) return;
  themeGrid.innerHTML = '';

  THEMES.forEach((theme) => {
    const btn = document.createElement('button');
    btn.className = `theme-btn ${theme.id === currentTheme ? 'active' : ''}`;
    btn.dataset.theme = theme.id;
    btn.type = 'button';
    btn.title = theme.name;

    btn.innerHTML = `
      <span class="theme-swatch" style="background: ${theme.color}22; border: 1.5px solid ${theme.color}66;">
        ${theme.emoji}
      </span>
      <span class="theme-name">${theme.name}</span>
    `;

    btn.addEventListener('click', () => {
      applyTheme(theme.id, true);
    });

    themeGrid.appendChild(btn);
  });
}

function applyTheme(themeId, notify = false) {
  const safeId = getValidThemeId(themeId);
  const theme = getThemeById(safeId);
  currentTheme = safeId;
  localStorage.setItem('mabiu_theme', safeId);
  document.documentElement.setAttribute('data-theme', safeId);

  if (themeColorMeta) {
    themeColorMeta.setAttribute('content', theme.color);
  }

  // Update theme buttons in settings
  if (themeGrid) {
    themeGrid.querySelectorAll('.theme-btn').forEach((btn) => {
      btn.classList.toggle('active', btn.dataset.theme === safeId);
    });
  }

  // Update theme choice items in theme modal
  if (themeSelectList) {
    themeSelectList.querySelectorAll('.theme-choice-item').forEach((item) => {
      item.classList.toggle('active', item.dataset.theme === safeId);
    });
  }

  // Header theme badge emoji
  if (headerThemeEmoji) {
    headerThemeEmoji.textContent = theme.emoji;
  }

  // Status banner icon emoji matching theme
  if (statusThemeEmoji) {
    statusThemeEmoji.textContent = theme.emoji;
  }

  if (emojisEnabled) {
    if (appliedDecorationTheme !== safeId) {
      renderFloatingEmojisForTheme(theme);
    }
  } else {
    clearFloatingEmojis();
  }

  if (notify) {
    showToast(emojisEnabled ? `Тема: ${theme.emoji} ${theme.name}` : `Тема: ${theme.name}`);
  }
}

function openThemeModal() {
  renderThemeSelectList();
  if (themeModal) themeModal.classList.add('open');
}

function closeThemeModal() {
  if (themeModal) themeModal.classList.remove('open');
}

function renderThemeSelectList() {
  if (!themeSelectList) return;
  themeSelectList.innerHTML = '';

  THEMES.forEach((theme) => {
    const isActive = (theme.id === currentTheme);
    const item = document.createElement('div');
    item.className = `theme-choice-item ${isActive ? 'active' : ''}`;
    item.dataset.theme = theme.id;

    item.innerHTML = `
      <div class="theme-choice-left">
        <span class="theme-choice-swatch" style="background: ${theme.color}22; border: 1px solid ${theme.color}44;">
          ${theme.emoji}
        </span>
        <div class="theme-choice-info">
          <div class="theme-choice-title">${theme.name}</div>
          <div class="theme-choice-desc">${theme.desc}</div>
        </div>
      </div>
      <svg class="theme-choice-check" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
        <polyline points="20 6 9 17 4 12"></polyline>
      </svg>
    `;

    item.addEventListener('click', () => {
      applyTheme(theme.id, true);
      closeThemeModal();
    });

    themeSelectList.appendChild(item);
  });
}

function cycleTheme() {
  const currentIndex = THEMES.findIndex(t => t.id === currentTheme);
  const nextIndex = (currentIndex + 1) % THEMES.length;
  applyTheme(THEMES[nextIndex].id, true);
}

// --- Data Loading & Persistence ---

function updateGroupLabelUI() {
  if (currentGroupLabel) {
    currentGroupLabel.textContent = currentGroup || 'Выбрать направление';
  }
}

async function initApp() {
  renderThemeGrid();
  setEmojisEnabled(emojisEnabled, false);
  applyTheme(currentTheme, false);
  updateGroupLabelUI();
  if (diskSourceLabel) diskSourceLabel.textContent = CONFIG.yandexKeyId;

  // Clear stale service worker caches from window
  if ('caches' in window) {
    caches.keys().then((keys) => {
      keys.forEach((k) => {
        if (k !== 'mabiu-schedule-v11') {
          caches.delete(k);
        }
      });
    }).catch(() => {});
  }

  // 1. Try loading cached schedule from localStorage (preserving 7-day rolling history)
  const cachedJson = localStorage.getItem('mabiu_schedule_data');
  if (cachedJson) {
    try {
      const parsed = JSON.parse(cachedJson);
      if (parsed && parsed.fileName && parsed.schedules && Object.keys(parsed.schedules).length > 0) {
        scheduleData = pruneScheduleData(normalizeScheduleData(parsed));
        saveScheduleDataSafe(scheduleData);
      } else {
        localStorage.removeItem('mabiu_schedule_data');
      }
    } catch (e) {
      console.warn('Failed to parse cached schedule data, resetting:', e);
      localStorage.removeItem('mabiu_schedule_data');
    }
  }

  // 2. If no valid cache, load bundled schedule.json (checks both ./schedule.json and ./data/schedule.json)
  if (!scheduleData) {
    try {
      let resp = await fetch(`./schedule.json?_t=${Date.now()}`);
      if (!resp.ok) resp = await fetch(`./data/schedule.json?_t=${Date.now()}`);
      if (resp.ok) {
        scheduleData = pruneScheduleData(normalizeScheduleData(await resp.json()));
        saveScheduleDataSafe(scheduleData);
      }
    } catch (err) {
      console.error('Failed to load bundled schedule.json:', err);
    }
  }

  // 3. Fallback structure if both failed
  if (!scheduleData) {
    scheduleData = {
      groups: currentGroup ? [currentGroup] : [],
      schedules: {},
      fileName: 'default',
      lastUpdated: Date.now()
    };
  }

  // Validate previously saved group against loaded groups (never auto-select if unset)
  if (currentGroup && scheduleData.groups && scheduleData.groups.length > 0 && !scheduleData.groups.includes(currentGroup)) {
    currentGroup = null;
    localStorage.removeItem('mabiu_group');
  }
  updateGroupLabelUI();

  updateSettingsLabels();
  setupDaysCarousel();
  renderLessons();
  startLiveTicker();

  // Prompt user to pick direction immediately if not saved yet
  if (!currentGroup) {
    openGroupModal({ autoFocus: false });
  }

  // 4. Background check for fresh schedule.json from repository (checks both ./schedule.json and ./data/schedule.json)
  (async () => {
    try {
      const candidates = ['./schedule.json', './data/schedule.json'];
      let bestFresh = null;
      for (const c of candidates) {
        try {
          const r = await fetch(`${c}?_t=${Date.now()}`);
          if (r.ok) {
            const j = await r.json();
            if (j && j.groups && j.lastUpdated) {
              if (!bestFresh || j.lastUpdated > bestFresh.lastUpdated) {
                bestFresh = j;
              }
            }
          }
        } catch (_) {}
      }

      const hasExistingSchedules = Boolean(
        scheduleData && scheduleData.schedules && Object.keys(scheduleData.schedules).length > 0
      );
      const freshHasSchedules = Boolean(
        bestFresh && bestFresh.schedules && Object.keys(bestFresh.schedules).length > 0
      );
      const shouldUpdate = freshHasSchedules && (
        !hasExistingSchedules ||
        (bestFresh.lastUpdated && bestFresh.lastUpdated > (scheduleData.lastUpdated || 0))
      );

      if (shouldUpdate) {
        const prevSignature = scheduleData?.sourceSignature;
        bestFresh = pruneScheduleData(normalizeScheduleData(bestFresh));
        scheduleData = mergeScheduleData(scheduleData, bestFresh);
        if (!bestFresh.sourceSignature && prevSignature && bestFresh.fileName === scheduleData.fileName) {
          scheduleData.sourceSignature = prevSignature;
        }
        saveScheduleDataSafe(scheduleData);
        if (currentGroup && scheduleData.groups && !scheduleData.groups.includes(currentGroup)) {
          currentGroup = null;
          localStorage.removeItem('mabiu_group');
          updateGroupLabelUI();
        }
        updateSettingsLabels();
        setupDaysCarousel();
        renderLessons();
        updateTickerOnce();
        if (groupModal && groupModal.classList.contains('open')) {
          renderGroupList(groupSearchInput ? groupSearchInput.value : '');
        }
        showToast(`Расписание обновлено: ${bestFresh.fileName}`);
      }
    } catch (_) {}
  })();

  // 5. Background direct check for new .xlsx on Yandex.Disk
  checkYandexDiskUpdates(false);
}

function updateSettingsLabels() {
  if (currentFileNameLabel && scheduleData) {
    currentFileNameLabel.textContent = scheduleData.fileName || 'Не загружено';
  }
  if (lastUpdatedLabel && scheduleData && scheduleData.lastUpdated) {
    const d = new Date(scheduleData.lastUpdated);
    lastUpdatedLabel.textContent = d.toLocaleString('ru-RU', {
      day: '2-digit',
      month: '2-digit',
      hour: '2-digit',
      minute: '2-digit'
    });
  }
}

// --- Days Navigation & Carousel ---

function getSortedAvailableDates() {
  if (!scheduleData) return [getTodayDateStr()];
  const groupSchedule = currentGroup ? scheduleData.schedules?.[currentGroup] : null;
  const dateSet = new Set();

  // Include dates that have lessons or are in schedules
  if (groupSchedule && groupSchedule.days) {
    for (const d of Object.keys(groupSchedule.days)) {
      dateSet.add(d);
    }
  }

  // Also include the surrounding days of current week (Monday to Saturday)
  const today = new Date();
  const currentDow = today.getDay() === 0 ? 7 : today.getDay();
  const monday = new Date(today);
  monday.setDate(today.getDate() - (currentDow - 1));

  for (let i = 0; i < 6; i++) {
    const d = new Date(monday);
    d.setDate(monday.getDate() + i);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    dateSet.add(`${y}-${m}-${day}`);
  }

  // Convert to sorted array
  return Array.from(dateSet).sort();
}

function setupDaysCarousel() {
  const dates = getSortedAvailableDates();
  const todayStr = getTodayDateStr();

  // If selectedDateStr is not in list, find closest or select today
  if (!dates.includes(selectedDateStr)) {
    if (dates.includes(todayStr)) {
      selectedDateStr = todayStr;
    } else if (dates.length > 0) {
      selectedDateStr = dates[0];
    }
  }

  daysCarousel.innerHTML = '';
  const dayNamesShort = ['ВС', 'ПН', 'ВТ', 'СР', 'ЧТ', 'ПТ', 'СБ'];

  // Track month for header
  const [sy, sm] = selectedDateStr.split('-').map(Number);
  monthLabel.textContent = getMonthNameRu(sm - 1, sy);

  dates.forEach((dStr) => {
    const [y, m, d] = dStr.split('-').map(Number);
    const dateObj = new Date(y, m - 1, d);
    const dow = dateObj.getDay();
    const isWeekend = (dow === 0 || dow === 6);
    const isToday = (dStr === todayStr);
    const isActive = (dStr === selectedDateStr);

    const pill = document.createElement('div');
    pill.className = `day-pill ${isWeekend ? 'weekend' : ''} ${isActive ? 'active' : ''}`;
    pill.dataset.date = dStr;

    pill.innerHTML = `
      <span class="day-name">${dayNamesShort[dow]}</span>
      <span class="day-number">${d}</span>
      ${isToday ? '<span class="today-dot"></span>' : ''}
    `;

    pill.addEventListener('click', () => {
      selectDay(dStr);
    });

    daysCarousel.appendChild(pill);
  });

  // Scroll active pill into view smoothly
  requestAnimationFrame(() => {
    const activePill = daysCarousel.querySelector('.day-pill.active');
    if (activePill) {
      activePill.scrollIntoView({ behavior: 'smooth', inline: 'center', block: 'nearest' });
    }
  });
}

function selectDay(dateStr) {
  selectedDateStr = dateStr;

  // Update pills UI
  const pills = daysCarousel.querySelectorAll('.day-pill');
  pills.forEach((p) => {
    p.classList.toggle('active', p.dataset.date === dateStr);
  });

  const activePill = daysCarousel.querySelector('.day-pill.active');
  if (activePill) {
    activePill.scrollIntoView({ behavior: 'smooth', inline: 'center', block: 'nearest' });
  }

  // Update month label
  const [y, m] = dateStr.split('-').map(Number);
  monthLabel.textContent = getMonthNameRu(m - 1, y);

  renderLessons();
}

// --- Schedule Lessons Rendering ---

function renderLessons() {
  const isToday = (selectedDateStr === getTodayDateStr());
  dayHeadingText.textContent = formatDateToRu(selectedDateStr);
  dayBadgeToday.style.display = isToday ? 'inline-block' : 'none';

  lessonsList.innerHTML = '';

  if (!currentGroup) {
    lessonsList.innerHTML = `
      <div class="empty-card">
        <svg class="empty-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
          <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
          <circle cx="9" cy="7" r="4"></circle>
          <path d="M23 21v-2a4 4 0 0 0-3-3.87"></path>
          <path d="M16 3.13a4 4 0 0 1 0 7.75"></path>
        </svg>
        <div class="empty-title">Направление не выбрано</div>
        <div class="empty-desc">Укажите вашу группу из списка, чтобы открыть расписание пар и таймер.</div>
        <button id="selectGroupPromptBtn" class="select-group-cta" type="button">Выбрать направление</button>
      </div>
    `;
    const promptBtn = document.getElementById('selectGroupPromptBtn');
    if (promptBtn) {
      promptBtn.addEventListener('click', () => openGroupModal());
    }
    return;
  }

  const groupSchedule = scheduleData?.schedules?.[currentGroup];
  const lessons = groupSchedule?.days?.[selectedDateStr] || [];

  if (!lessons || lessons.length === 0) {
    // Empty state
    const [y, m, d] = selectedDateStr.split('-').map(Number);
    const dateObj = new Date(y, m - 1, d);
    const isWeekend = (dateObj.getDay() === 0 || dateObj.getDay() === 6);

    const nextLessonsInfo = findNextDayWithLessons(selectedDateStr);

    lessonsList.innerHTML = `
      <div class="empty-card">
        <svg class="empty-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
          ${isWeekend ?
            '<circle cx="12" cy="12" r="10"></circle><path d="M8 14s1.5 2 4 2 4-2 4-2"></path><line x1="9" y1="9" x2="9.01" y2="9"></line><line x1="15" y1="9" x2="15.01" y2="9"></line>' :
            '<rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect><line x1="16" y1="2" x2="16" y2="6"></line><line x1="8" y1="2" x2="8" y2="6"></line><line x1="3" y1="10" x2="21" y2="10"></line>'
          }
        </svg>
        <div class="empty-title">${isWeekend ? (emojisEnabled ? 'Выходной день 🌴' : 'Выходной день') : (emojisEnabled ? 'Пар нет 🎉' : 'Пар нет')}</div>
        <div class="empty-desc">${isWeekend ? 'В этот день занятий нет, можно отдохнуть!' : 'На этот день занятий не запланировано.'}</div>
        ${nextLessonsInfo ? `<div class="next-lessons-pill">${nextLessonsInfo}</div>` : ''}
      </div>
    `;
    return;
  }

  // Render Lesson Cards
  lessons.forEach((lesson) => {
    const card = document.createElement('div');
    card.className = 'lesson-card';

    const interval = parseLessonTimeInterval(lesson.time);
    if (interval) {
      card.dataset.startSec = String(interval.startSec);
      card.dataset.endSec = String(interval.endSec);
    }

    // Lesson type badge styling
    let typeClass = 'type-practice';
    const typeLower = (lesson.lessonType || '').toLowerCase();
    if (typeLower.includes('лекц') || typeLower.includes('лек')) {
      typeClass = 'type-lecture';
    } else if (typeLower.includes('лаб')) {
      typeClass = 'type-lab';
    } else if (typeLower.includes('семинар')) {
      typeClass = 'type-seminar';
    }

    // Joint groups text
    const hasJoint = lesson.jointGroups && lesson.jointGroups.length > 0;
    const jointText = hasJoint ? lesson.jointGroups.join(', ') : '';

    card.innerHTML = `
      <div class="lesson-top-row">
        <div class="pair-time-group">
          <span class="pair-badge">${escapeHtml(lesson.pairNumber || 'Пара')}</span>
          <span class="pair-time">${escapeHtml(lesson.time || '')}</span>
        </div>
        ${lesson.room ? `<span class="room-badge">${escapeHtml(lesson.room)}</span>` : ''}
      </div>
      <div class="lesson-subject-row">
        <div class="lesson-subject">${escapeHtml(lesson.cleanSubject || lesson.subject || 'Занятие')}</div>
      </div>
      <div class="lesson-badges-row">
        ${lesson.lessonType ? `<span class="type-chip ${typeClass}">${escapeHtml(lesson.lessonType)}</span>` : ''}
      </div>
      ${lesson.teacher ? `
        <div class="lesson-teacher">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path>
            <circle cx="12" cy="7" r="4"></circle>
          </svg>
          <span>${escapeHtml(lesson.teacher)}</span>
        </div>
      ` : ''}
      ${hasJoint ? `
        <div class="joint-groups-badge" title="Совместное занятие">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
            <circle cx="9" cy="7" r="4"></circle>
            <path d="M23 21v-2a4 4 0 0 0-3-3.87"></path>
            <path d="M16 3.13a4 4 0 0 1 0 7.75"></path>
          </svg>
          <span>Совместно с: ${escapeHtml(jointText)}</span>
        </div>
      ` : ''}
      ${interval ? `
        <div class="lesson-progress">
          <div class="lesson-progress-meta">
            <span class="lesson-progress-status">
              <span class="lesson-live-dot" aria-hidden="true"></span>
              <span class="lesson-progress-status-text">Идёт сейчас · 0%</span>
            </span>
            <span class="lesson-progress-remaining">Осталось 00:00</span>
          </div>
          <div class="lesson-progress-track" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0">
            <div class="lesson-progress-fill"></div>
          </div>
        </div>
      ` : ''}
    `;

    lessonsList.appendChild(card);
  });

  updateLessonCardsProgress();
}

function findNextDayWithLessons(fromDateStr) {
  if (!currentGroup) return null;
  const groupSchedule = scheduleData?.schedules?.[currentGroup];
  if (!groupSchedule?.days) return null;

  const dates = Object.keys(groupSchedule.days).sort();
  for (const d of dates) {
    if (d > fromDateStr && groupSchedule.days[d].length > 0) {
      const firstPair = groupSchedule.days[d][0];
      const [y, m, day] = d.split('-').map(Number);
      const dateObj = new Date(y, m - 1, day);
      const daysOfWeek = ['воскресенье', 'понедельник', 'вторник', 'среду', 'четверг', 'пятницу', 'субботу'];
      const dowName = daysOfWeek[dateObj.getDay()];
      const timeStart = (firstPair.time || '').split(/[-–—]/)[0].trim().replace('.', ':');
      return `Следующие пары в ${dowName} (${day}.${String(m).padStart(2, '0')}) к ${timeStart}`;
    }
  }
  return null;
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// --- Live Countdown Ticker & Active Lesson Progress (Real-Time 1s Tick) ---

function parseLessonTimeInterval(timeStr) {
  const rawTime = String(timeStr || '').replace(/\s+/g, '');
  const m = rawTime.match(/(\d{1,2})[.:](\d{2})[-–—](\d{1,2})[.:](\d{2})/);
  if (!m) return null;

  const sh = parseInt(m[1], 10);
  const sm = parseInt(m[2], 10);
  const eh = parseInt(m[3], 10);
  const em = parseInt(m[4], 10);
  const startSec = (sh * 60 + sm) * 60;
  const endSec = (eh * 60 + em) * 60;
  if (endSec <= startSec) return null;

  return {
    startSec,
    endSec,
    startTimeStr: `${String(sh).padStart(2, '0')}:${String(sm).padStart(2, '0')}`,
    endTimeStr: `${String(eh).padStart(2, '0')}:${String(em).padStart(2, '0')}`
  };
}

function updateLessonCardsProgress(now = new Date()) {
  if (!lessonsList) return;
  const cards = lessonsList.querySelectorAll('.lesson-card');
  if (cards.length === 0) return;

  const todayStr = getTodayDateStr();
  const isPastDay = selectedDateStr < todayStr;
  const isViewingToday = (selectedDateStr === todayStr);
  const currentTimeInSec = now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds();

  cards.forEach((card) => {
    if (isPastDay) {
      card.classList.remove('lesson-active');
      card.classList.add('lesson-finished');
      return;
    }

    const startSec = parseInt(card.dataset.startSec || '', 10);
    const endSec = parseInt(card.dataset.endSec || '', 10);

    if (!isViewingToday || isNaN(startSec) || isNaN(endSec) || endSec <= startSec) {
      card.classList.remove('lesson-active', 'lesson-finished');
      return;
    }

    if (currentTimeInSec >= startSec && currentTimeInSec <= endSec) {
      card.classList.add('lesson-active');
      card.classList.remove('lesson-finished');

      const totalSec = endSec - startSec;
      const elapsedSec = currentTimeInSec - startSec;
      const remainingSec = endSec - currentTimeInSec;
      const percent = Math.min(100, Math.max(0, (elapsedSec / totalSec) * 100));
      const percentRounded = Math.round(percent);

      const statusTextEl = card.querySelector('.lesson-progress-status-text');
      const remainingEl = card.querySelector('.lesson-progress-remaining');
      const trackEl = card.querySelector('.lesson-progress-track');
      const fillEl = card.querySelector('.lesson-progress-fill');

      if (statusTextEl) statusTextEl.textContent = `Идёт сейчас · ${percentRounded}%`;
      if (remainingEl) remainingEl.textContent = `Осталось ${formatDuration(remainingSec * 1000)}`;
      if (trackEl) trackEl.setAttribute('aria-valuenow', String(percentRounded));
      if (fillEl) fillEl.style.width = `${percent.toFixed(2)}%`;
    } else if (currentTimeInSec > endSec) {
      card.classList.remove('lesson-active');
      card.classList.add('lesson-finished');
    } else {
      card.classList.remove('lesson-active', 'lesson-finished');
    }
  });
}

function startLiveTicker() {
  if (tickerTimerId) clearInterval(tickerTimerId);
  updateTickerOnce();
  tickerTimerId = setInterval(updateTickerOnce, 1000);
}

function updateTickerOnce() {
  if (!currentGroup) {
    if (statusBanner) statusBanner.classList.add('clickable');
    statusBadge.textContent = 'Начало работы';
    statusBadge.style.color = 'var(--primary)';
    statusBadge.style.backgroundColor = 'var(--primary-light)';
    statusTitle.textContent = 'Выберите направление';
    statusSubtitle.textContent = 'Нажмите сюда, чтобы указать вашу группу';
    return;
  }

  if (statusBanner) statusBanner.classList.remove('clickable');
  const now = new Date();
  updateLessonCardsProgress(now);

  const todayStr = getTodayDateStr();
  const groupSchedule = scheduleData?.schedules?.[currentGroup];
  const todayLessons = groupSchedule?.days?.[todayStr] || [];

  if (!todayLessons || todayLessons.length === 0) {
    // No lessons today. Find next lesson day!
    const nextInfo = findNextDayWithLessons(todayStr);
    statusBadge.textContent = 'Выходной';
    statusBadge.style.color = 'var(--text-muted)';
    statusBadge.style.backgroundColor = 'var(--surface-variant)';
    statusTitle.textContent = emojisEnabled ? 'Сегодня занятий нет 🎉' : 'Сегодня занятий нет';
    statusSubtitle.textContent = nextInfo || 'Расписание на следующие дни пока не загружено';
    return;
  }

  // Parse lesson intervals for today
  const currentTimeInSec = now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds();

  const parsedLessons = todayLessons.map((l) => {
    const interval = parseLessonTimeInterval(l.time);
    if (!interval) return null;
    return {
      lesson: l,
      startSec: interval.startSec,
      endSec: interval.endSec,
      startTimeStr: interval.startTimeStr,
      endTimeStr: interval.endTimeStr
    };
  }).filter(Boolean);

  if (parsedLessons.length === 0) {
    statusBadge.textContent = 'Расписание';
    statusTitle.textContent = `${todayLessons.length} пар(ы) сегодня`;
    statusSubtitle.textContent = 'Проверьте список занятий ниже';
    return;
  }

  const firstLesson = parsedLessons[0];
  const lastLesson = parsedLessons[parsedLessons.length - 1];

  // 1. Before first lesson
  if (currentTimeInSec < firstLesson.startSec) {
    const diffSec = firstLesson.startSec - currentTimeInSec;
    const diffMs = diffSec * 1000;
    statusBadge.textContent = 'Скоро начнётся';
    statusBadge.style.color = 'var(--primary)';
    statusBadge.style.backgroundColor = 'var(--primary-light)';
    statusTitle.textContent = `До 1-й пары: ${formatDuration(diffMs)}`;
    statusSubtitle.textContent = `В ${firstLesson.startTimeStr} — ${firstLesson.lesson.cleanSubject || firstLesson.lesson.subject}`;
    return;
  }

  // 2. Check if currently inside a lesson
  for (let i = 0; i < parsedLessons.length; i++) {
    const p = parsedLessons[i];
    if (currentTimeInSec >= p.startSec && currentTimeInSec <= p.endSec) {
      const diffSec = p.endSec - currentTimeInSec;
      const diffMs = diffSec * 1000;
      statusBadge.textContent = 'Идёт пара';
      statusBadge.style.color = '#16a34a';
      statusBadge.style.backgroundColor = 'rgba(22, 163, 74, 0.14)';
      statusTitle.textContent = `Осталось ${formatDuration(diffMs)} (до ${p.endTimeStr})`;
      statusSubtitle.textContent = `${p.lesson.pairNumber || 'Пара'}: ${p.lesson.cleanSubject || p.lesson.subject}${p.lesson.room ? ` (${p.lesson.room})` : ''}`;
      return;
    }
  }

  // 3. Check if in break between lessons
  for (let i = 0; i < parsedLessons.length - 1; i++) {
    const currentP = parsedLessons[i];
    const nextP = parsedLessons[i + 1];
    if (currentTimeInSec > currentP.endSec && currentTimeInSec < nextP.startSec) {
      const diffSec = nextP.startSec - currentTimeInSec;
      const diffMs = diffSec * 1000;
      statusBadge.textContent = 'Перерыв';
      statusBadge.style.color = '#d97706';
      statusBadge.style.backgroundColor = 'rgba(217, 119, 6, 0.14)';
      statusTitle.textContent = `Перерыв: до пары ${formatDuration(diffMs)}`;
      statusSubtitle.textContent = `Далее в ${nextP.startTimeStr}: ${nextP.lesson.cleanSubject || nextP.lesson.subject}`;
      return;
    }
  }

  // 4. Lessons for today are completely finished!
  if (currentTimeInSec > lastLesson.endSec) {
    statusBadge.textContent = 'Завершено';
    statusBadge.style.color = 'var(--text-muted)';
    statusBadge.style.backgroundColor = 'var(--surface-variant)';
    statusTitle.textContent = 'Пары на сегодня закончились!';

    // Calculate tomorrow's start time as requested by user
    const tomorrow = new Date(now);
    tomorrow.setDate(now.getDate() + 1);
    const ty = tomorrow.getFullYear();
    const tm = String(tomorrow.getMonth() + 1).padStart(2, '0');
    const td = String(tomorrow.getDate()).padStart(2, '0');
    const tomorrowStr = `${ty}-${tm}-${td}`;

    const tomorrowLessons = groupSchedule?.days?.[tomorrowStr] || [];
    if (tomorrowLessons.length > 0) {
      const firstT = tomorrowLessons[0];
      const rawTTime = (firstT.time || '').split(/[-–—]/)[0].trim().replace('.', ':');
      statusSubtitle.textContent = `Завтра ко стольки: к ${rawTTime} (${firstT.pairNumber || '1 пара'}, ${firstT.cleanSubject || firstT.subject})`;
    } else {
      // Find next day
      const nextInfo = findNextDayWithLessons(todayStr);
      statusSubtitle.textContent = nextInfo || 'Завтра пар нет, приятного отдыха!';
    }
  }
}

// --- Group Selector Modal ---

function openGroupModal(options = {}) {
  const autoFocus = !(options && options.autoFocus === false);
  groupSearchInput.value = '';
  renderGroupList('');
  groupModal.classList.add('open');
  if (autoFocus) {
    setTimeout(() => groupSearchInput.focus(), 150);
  }
}

function closeGroupModal() {
  groupModal.classList.remove('open');
}

function renderGroupList(query) {
  groupList.innerHTML = '';
  const q = query.trim().toLowerCase();
  const allGroups = (scheduleData?.groups && scheduleData.groups.length > 0)
    ? scheduleData.groups
    : (currentGroup ? [currentGroup] : []);

  // Filter 26 cohort groups
  const filtered = allGroups.filter((g) => {
    return g.toLowerCase().includes(q);
  });

  if (filtered.length === 0) {
    groupList.innerHTML = '<div style="text-align: center; padding: 24px; color: var(--text-muted); font-size: 14px;">Направления не найдены</div>';
    return;
  }

  filtered.forEach((grp) => {
    const item = document.createElement('div');
    const isActive = (grp === currentGroup);
    item.className = `group-item ${isActive ? 'active' : ''}`;
    item.textContent = grp;

    item.addEventListener('click', () => {
      setGroup(grp);
      closeGroupModal();
    });

    groupList.appendChild(item);
  });
}

function setGroup(groupName) {
  if (!groupName || currentGroup === groupName) return;
  currentGroup = groupName;
  localStorage.setItem('mabiu_group', currentGroup);
  updateGroupLabelUI();

  setupDaysCarousel();
  renderLessons();
  updateTickerOnce();
  showToast(`Выбрано направление: ${groupName}`);
}

// --- Settings & Manual XLSX Upload ---

function openSettingsModal() {
  updateSettingsLabels();
  settingsModal.classList.add('open');
}

function closeSettingsModal() {
  settingsModal.classList.remove('open');
}

function setupDropzone() {
  dropzone.addEventListener('click', () => {
    xlsxFileInput.click();
  });

  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dragover');
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleManualFile(e.dataTransfer.files[0]);
    }
  });

  xlsxFileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleManualFile(e.target.files[0]);
    }
  });
}

async function handleManualFile(file) {
  if (!file.name.match(/\.(xlsx|xls)$/i)) {
    showToast('Пожалуйста, выберите файл Excel (.xlsx или .xls)');
    return;
  }

  showToast('Чтение и разбор файла...');
  try {
    const arrayBuffer = await file.arrayBuffer();
    const parsedData = parseWorkbookClient(arrayBuffer, file.name);

    if (!parsedData || !parsedData.groups || parsedData.groups.length === 0) {
      throw new Error('В файле не найдено групп набора 26');
    }

    const normalizedIncoming = pruneScheduleData(normalizeScheduleData(parsedData));
    scheduleData = mergeScheduleData(scheduleData, normalizedIncoming);
    scheduleData.fileName = parsedData.fileName;
    scheduleData.lastUpdated = Date.now();
    scheduleData.sourceSignature = `manual:${file.name}:${file.lastModified || Date.now()}`;
    saveScheduleDataSafe(scheduleData);

    // Validate previously selected group (do not auto-pick if unset)
    if (currentGroup && !scheduleData.groups.includes(currentGroup)) {
      currentGroup = null;
      localStorage.removeItem('mabiu_group');
      updateGroupLabelUI();
    }

    updateSettingsLabels();
    setupDaysCarousel();
    renderLessons();
    updateTickerOnce();
    closeSettingsModal();
    showToast(`Файл «${file.name}» успешно загружен!`);
    if (!currentGroup) {
      openGroupModal({ autoFocus: false });
    }
  } catch (err) {
    console.error('Manual parse error:', err);
    showToast(`Ошибка разбора файла: ${err.message || 'Неверный формат'}`);
  }
}

// Client-Side SheetJS Parser with Auto Date Correction
function parseWorkbookClient(arrayBuffer, fileName) {
  if (typeof XLSX === 'undefined') {
    throw new Error('Библиотека XLSX не загружена');
  }

  const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
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

  if (allGroups.length === 0) {
    throw new Error('Не найдено столбцов с группами 26 набора в строке 2');
  }

  // Row scan for dates & days of week
  const detectedDates = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    let foundDate = null;
    let foundDow = null;
    let foundDowText = '';

    for (let c = 0; c <= Math.min(7, range.e.c); c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v === undefined || cell.v === null) continue;

      if (!foundDate) {
        foundDate = parseDateFromCellClient(cell);
      }
      if (!foundDow) {
        const dow = parseDayOfWeekClient(cell.w || cell.v);
        if (dow) {
          foundDow = dow;
          foundDowText = String(cell.w || cell.v).trim();
        }
      }
    }

    if (foundDate || foundDow) {
      detectedDates.push({ r, date: foundDate, dow: foundDow, dowText: foundDowText });
    }
  }

  // Auto-correct dates using day-of-week context
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
        // Date typo detected in file! Correct based on dow
        if (lastGoodDateObj) {
          const expectedD = new Date(lastGoodDateObj);
          let diffDays = cur.dow - (expectedD.getUTCDay() === 0 ? 7 : expectedD.getUTCDay());
          if (diffDays <= 0) diffDays += 7;
          expectedD.setUTCDate(expectedD.getUTCDate() + diffDays);
          resolvedDateStr = expectedD.toISOString().split('T')[0];
        } else {
          // Adjust testD by offset
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

    // Fill rows up to next detected date
    const nextR = (i + 1 < detectedDates.length) ? detectedDates[i + 1].r : (range.e.r + 1);
    for (let row = cur.r; row < nextR; row++) {
      if (resolvedDateStr) rowDateMap[row] = resolvedDateStr;
    }
  }

  // Parse schedules per group
  const schedules = {};
  allGroups.forEach(g => { schedules[g] = { days: {} }; });

  const rawRowsByDate = {};
  for (let r = 2; r <= range.e.r; r++) {
    const dateStr = rowDateMap[r];
    if (!dateStr) continue;

    // Time cell in column 0 or 1
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
    const pairNum = getPairNumberClient(timeStr);

    if (!rawRowsByDate[dateStr]) rawRowsByDate[dateStr] = [];

    const cellsInRow = {};
    for (const [col, group] of Object.entries(groupsByCol)) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c: parseInt(col, 10) })];
      const val = cell && cell.v ? String(cell.w || cell.v).trim() : '';
      if (val) cellsInRow[group] = val;
    }

    rawRowsByDate[dateStr].push({ pairNum, timeStr, cellsInRow });
  }

  // Populate lessons and joint groups
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

        const parsed = parseLessonContentClient(content);
        if (!schedules[group].days[dateStr]) schedules[group].days[dateStr] = [];

        schedules[group].days[dateStr].push({
          pairNumber: entry.pairNum,
          time: entry.timeStr,
          subject: content,
          cleanSubject: parsed.cleanSubject,
          lessonType: parsed.lessonType,
          room: parsed.room,
          teacher: parsed.teacher,
          jointGroups: joint
        });
      }
    }
  }

  return {
    lastUpdated: Date.now(),
    fileName: fileName,
    yandexDiskUrl: CONFIG.yandexPublicKey,
    groups: allGroups,
    coveredDates: Object.keys(rawRowsByDate).sort(),
    schedules: schedules
  };
}

function parseDayOfWeekClient(text) {
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

function parseDateFromCellClient(cell) {
  if (!cell || cell.v === undefined || cell.v === null) return null;
  if (cell.t === 'd' || cell.v instanceof Date) {
    const d = cell.v instanceof Date ? cell.v : new Date(cell.w || cell.v);
    return d.toISOString().split('T')[0];
  }
  if (typeof cell.v === 'number' && cell.v > 40000 && cell.v < 60000) {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    excelEpoch.setUTCDate(excelEpoch.getUTCDate() + Math.floor(cell.v));
    return excelEpoch.toISOString().split('T')[0];
  }
  const text = String(cell.w || cell.v || '').trim();
  const m1 = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m1) return `${m1[1]}-${m1[2].padStart(2, '0')}-${m1[3].padStart(2, '0')}`;
  const m2 = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})/);
  if (m2) {
    let y = parseInt(m2[3], 10);
    if (y < 100) y += 2000;
    return `${y}-${m2[2].padStart(2, '0')}-${m2[1].padStart(2, '0')}`;
  }
  return null;
}

function getPairNumberClient(timeStr) {
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

function parseLessonContentClient(rawText) {
  let text = String(rawText || '').trim();
  let room = '';
  let teacher = '';
  let lessonType = '';

  // Extract room from end first
  const roomMatch = text.match(/(?:,\s*)?(?:(?:Ауд\.?|АУд\.?)\s*[\w\d/]*\s*(?:\([^)]+\))?|(?:Актовый\s*зал|Конференц[- ]*зал|Спорт\.?\s*зал|Он[- ]*лайн|ДО)(?:\s*\([^)]+\))?|\((?:Конференц[- ]*зал|Актовый\s*зал)\))\s*$/i);
  if (roomMatch) {
    room = roomMatch[0].replace(/^,\s*/, '').trim();
    text = text.substring(0, text.length - roomMatch[0].length).trim();
  }

  // Extract and strip lesson-type parentheticals (e.g. "(Лекция)", "(Практическое занятие)")
  text = text.replace(LESSON_TYPE_REGEX, (_, rawType) => {
    if (!lessonType) lessonType = normalizeLessonTypeLabel(rawType);
    return ' ';
  });
  text = cleanSubjectTextClient(text);

  // Extract teacher
  const teacherMatch = text.match(TEACHER_END_REGEX);
  if (teacherMatch) {
    teacher = teacherMatch[1].trim();
    text = text.substring(0, text.length - teacherMatch[0].length).trim();
  }

  // Subject is cleaned remainder
  let cleanSubject = cleanSubjectTextClient(text);
  if (!cleanSubject) cleanSubject = String(rawText || '').trim();

  return { cleanSubject, room, teacher, lessonType };
}

// --- Sync with Yandex.Disk ---

let isYandexSyncRunning = false;

async function downloadAndParseXlsxClient(fileItem) {
  if (!fileItem || !fileItem.file) return null;

  let arrayBuffer = null;
  try {
    const fileResp = await fetchWithTimeout(fileItem.file, { referrerPolicy: 'no-referrer', cache: 'no-store' }, 12000);
    if (fileResp.ok) {
      arrayBuffer = await fileResp.arrayBuffer();
    }
  } catch (_) {}

  if (!arrayBuffer || arrayBuffer.byteLength === 0) {
    const proxies = [
      `https://api.allorigins.win/raw?url=${encodeURIComponent(fileItem.file)}`,
      `https://corsproxy.org/?${encodeURIComponent(fileItem.file)}`
    ];
    for (const pUrl of proxies) {
      try {
        const fileResp = await fetchWithTimeout(pUrl, { referrerPolicy: 'no-referrer' }, 7000);
        if (fileResp.ok) {
          arrayBuffer = await fileResp.arrayBuffer();
          if (arrayBuffer && arrayBuffer.byteLength > 0) break;
        }
      } catch (_) {}
    }
  }

  if (!arrayBuffer || arrayBuffer.byteLength === 0) return null;

  try {
    const parsed = parseWorkbookClient(arrayBuffer, fileItem.name || 'schedule.xlsx');
    if (parsed && parsed.groups && parsed.groups.length > 0) {
      return normalizeScheduleData(parsed);
    }
  } catch (err) {
    console.warn(`Failed to parse workbook ${fileItem.name}:`, err);
  }
  return null;
}

async function checkYandexDiskUpdates(isManual = false) {
  if (isYandexSyncRunning) {
    if (isManual) showToast('Синхронизация уже выполняется...');
    return;
  }

  isYandexSyncRunning = true;
  if (refreshBtn) refreshBtn.classList.add('spinning');
  if (isManual) showToast('Проверка Яндекс.Диска...');

  try {
    let apiData = null;
    // 1. First try Vercel Serverless API (/api/schedule)
    try {
      const resp = await fetchWithTimeout('./api/schedule', {}, 4000);
      const contentType = resp.headers.get('content-type') || '';
      if (resp.ok && contentType.includes('json')) {
        const raw = await resp.json();
        if (raw && raw.groups && raw.groups.length > 0 && raw.schedules) {
          apiData = pruneScheduleData(normalizeScheduleData(raw));
        }
      }
    } catch (e) {
      // Local or static GitHub Pages fallback
    }

    let nextSchedule = null;
    const baseBeforeSync = apiData ? mergeScheduleData(scheduleData, apiData) : scheduleData;

    if (apiData && !needsRecentCoverage(baseBeforeSync)) {
      nextSchedule = baseBeforeSync;
    }

    // 2. Direct Yandex API fetch or fallback to CORS proxy
    if (!nextSchedule) {
      const yandexApi = `https://cloud-api.yandex.net/v1/disk/public/resources?public_key=${encodeURIComponent(CONFIG.yandexPublicKey)}&limit=100`;
      let diskJson = null;

      try {
        const resp = await fetchWithTimeout(yandexApi, { referrerPolicy: 'no-referrer' }, 10000);
        if (resp.ok) {
          diskJson = await resp.json();
        }
      } catch (e) {
        // Direct fetch failed, try cors proxy
      }

      if (!diskJson) {
        try {
          const proxyResp = await fetchWithTimeout(CONFIG.corsProxy + encodeURIComponent(yandexApi), {}, 7000);
          if (proxyResp.ok) {
            const rawJson = await proxyResp.json();
            diskJson = typeof rawJson.contents === 'string' ? JSON.parse(rawJson.contents) : rawJson;
          }
        } catch (_) {}
      }

      if (diskJson) {
        const items = diskJson._embedded?.items || [];
        const xlsxFiles = items.filter(it => it.name && it.name.match(/\.(xlsx|xls)$/i) && it.file);
        if (xlsxFiles.length > 0) {
          xlsxFiles.sort((a, b) => new Date(b.modified || 0) - new Date(a.modified || 0));
          const latestFile = xlsxFiles[0];
          const latestSig = `${latestFile.name}|${latestFile.modified || ''}`;

          const alreadyUpToDate = (
            !isManual &&
            baseBeforeSync &&
            baseBeforeSync.sourceSignature === latestSig &&
            !needsRecentCoverage(baseBeforeSync)
          );

          if (alreadyUpToDate) {
            if (apiData) nextSchedule = baseBeforeSync;
          } else {
            const latestParsed = await downloadAndParseXlsxClient(latestFile);
            if (latestParsed) {
              latestParsed.sourceSignature = latestSig;
              let mergedCandidate = mergeScheduleData(baseBeforeSync, latestParsed);

              const todayStr = getTodayDateStr();
              const [ty, tm, td] = todayStr.split('-').map(Number);
              const isTodaySunday = new Date(Date.UTC(ty, tm - 1, td)).getUTCDay() === 0;
              const latestMissesToday = !isTodaySunday && !collectCoveredDates(latestParsed).has(todayStr);
              const prevFile = xlsxFiles[1];
              const prevModMs = prevFile ? (new Date(prevFile.modified || 0).getTime() || 0) : 0;
              const baseUpdatedMs = baseBeforeSync?.lastUpdated || 0;

              const shouldBackfill = xlsxFiles.length > 1 && (
                needsRecentCoverage(mergedCandidate, todayStr) ||
                (latestMissesToday && prevModMs > baseUpdatedMs && baseBeforeSync?.sourceSignature !== latestSig)
              );

              if (shouldBackfill) {
                const maxAgeMs = 30 * 24 * 60 * 60 * 1000;
                const recentFallbacks = xlsxFiles
                  .slice(1, 4)
                  .filter(it => !it.modified || (Date.now() - new Date(it.modified).getTime()) <= maxAgeMs);

                const olderParsedList = [];
                for (const fbItem of recentFallbacks) {
                  const fbParsed = await downloadAndParseXlsxClient(fbItem);
                  if (fbParsed) {
                    olderParsedList.push(fbParsed);
                    let rebuilt = baseBeforeSync;
                    for (let idx = olderParsedList.length - 1; idx >= 0; idx--) {
                      rebuilt = mergeScheduleData(rebuilt, olderParsedList[idx]);
                    }
                    mergedCandidate = mergeScheduleData(rebuilt, latestParsed);
                    const curMissesToday = !isTodaySunday && !collectCoveredDates(mergedCandidate).has(todayStr);
                    if (!curMissesToday && countRecentWorkdays(mergedCandidate, todayStr) >= 5) {
                      break;
                    }
                  }
                }
              }

              mergedCandidate.fileName = latestParsed.fileName;
              mergedCandidate.sourceSignature = latestSig;
              mergedCandidate.lastUpdated = Date.now();
              nextSchedule = mergedCandidate;
            } else if (apiData) {
              nextSchedule = baseBeforeSync;
            }
          }
        }
      } else if (apiData) {
        nextSchedule = baseBeforeSync;
      }
    }

    if (nextSchedule && nextSchedule.groups && nextSchedule.groups.length > 0) {
      scheduleData = pruneScheduleData(normalizeScheduleData(nextSchedule));
      saveScheduleDataSafe(scheduleData);
      if (currentGroup && !scheduleData.groups.includes(currentGroup)) {
        currentGroup = null;
        localStorage.removeItem('mabiu_group');
        updateGroupLabelUI();
      }
      updateSettingsLabels();
      setupDaysCarousel();
      renderLessons();
      updateTickerOnce();
      if (groupModal && groupModal.classList.contains('open')) {
        renderGroupList(groupSearchInput ? groupSearchInput.value : '');
      }

      if (isManual) showToast('Расписание успешно обновлено с Яндекс.Диска!');
    } else {
      if (isManual) showToast('Расписание актуально.');
    }
  } catch (err) {
    console.warn('Sync attempt finished:', err);
    if (isManual) showToast('Синхронизация завершена. Используются актуальные данные.');
  } finally {
    isYandexSyncRunning = false;
    if (refreshBtn) refreshBtn.classList.remove('spinning');
  }
}

// Reset schedule to bundled defaults
async function resetSchedule() {
  localStorage.removeItem('mabiu_schedule_data');
  showToast('Сброс к исходному расписанию...');
  try {
    let resp = await fetch('./schedule.json');
    if (!resp.ok) resp = await fetch('./data/schedule.json');
    if (resp.ok) {
      scheduleData = pruneScheduleData(normalizeScheduleData(await resp.json()));
      saveScheduleDataSafe(scheduleData);
      if (currentGroup && scheduleData.groups && !scheduleData.groups.includes(currentGroup)) {
        currentGroup = null;
        localStorage.removeItem('mabiu_group');
        updateGroupLabelUI();
      }
      updateSettingsLabels();
      setupDaysCarousel();
      renderLessons();
      updateTickerOnce();
      closeSettingsModal();
      showToast('Расписание успешно сброшено!');
      if (!currentGroup) {
        openGroupModal({ autoFocus: false });
      }
    }
  } catch (err) {
    showToast('Не удалось загрузить исходное расписание');
  }
}

// --- Touch Swiping Gestures on Schedule List ---

function setupSwipeGestures() {
  let touchStartX = 0;
  let touchStartY = 0;
  let touchEndX = 0;
  let touchEndY = 0;

  scheduleContainer.addEventListener('touchstart', (e) => {
    touchStartX = e.changedTouches[0].screenX;
    touchStartY = e.changedTouches[0].screenY;
  }, { passive: true });

  scheduleContainer.addEventListener('touchend', (e) => {
    touchEndX = e.changedTouches[0].screenX;
    touchEndY = e.changedTouches[0].screenY;
    handleSwipe();
  }, { passive: true });

  function handleSwipe() {
    const deltaX = touchEndX - touchStartX;
    const deltaY = touchEndY - touchStartY;

    // Only horizontal swipe (not scrolling vertically)
    if (Math.abs(deltaX) > 55 && Math.abs(deltaX) > Math.abs(deltaY) * 1.5) {
      const dates = getSortedAvailableDates();
      const currentIndex = dates.indexOf(selectedDateStr);
      if (currentIndex === -1) return;

      if (deltaX < 0 && currentIndex < dates.length - 1) {
        // Swiped left -> next day
        selectDay(dates[currentIndex + 1]);
      } else if (deltaX > 0 && currentIndex > 0) {
        // Swiped right -> previous day
        selectDay(dates[currentIndex - 1]);
      }
    }
  }
}

// --- Event Listeners Setup ---

function setupEventListeners() {
  // Group modal
  groupSelectorBtn.addEventListener('click', openGroupModal);
  closeGroupModalBtn.addEventListener('click', closeGroupModal);
  groupModal.addEventListener('click', (e) => {
    if (e.target === groupModal) closeGroupModal();
  });
  groupSearchInput.addEventListener('input', (e) => {
    renderGroupList(e.target.value);
  });

  // Settings modal
  settingsBtn.addEventListener('click', openSettingsModal);
  closeSettingsModalBtn.addEventListener('click', closeSettingsModal);
  settingsModal.addEventListener('click', (e) => {
    if (e.target === settingsModal) closeSettingsModal();
  });
  if (emojisToggle) {
    emojisToggle.addEventListener('change', (e) => {
      setEmojisEnabled(e.target.checked, true);
    });
  }

  // Action buttons
  refreshBtn.addEventListener('click', () => checkYandexDiskUpdates(true));
  syncNowBtn.addEventListener('click', () => checkYandexDiskUpdates(true));
  resetScheduleBtn.addEventListener('click', resetSchedule);

  todayJumpBtn.addEventListener('click', () => {
    selectDay(getTodayDateStr());
  });

  // Themes Modal & Clickable Banner Prompt
  if (statusBanner) {
    statusBanner.addEventListener('click', () => {
      if (!currentGroup) openGroupModal();
    });
  }
  if (themeBtn) {
    themeBtn.addEventListener('click', openThemeModal);
  }
  if (closeThemeModalBtn) {
    closeThemeModalBtn.addEventListener('click', closeThemeModal);
  }
  if (themeModal) {
    themeModal.addEventListener('click', (e) => {
      if (e.target === themeModal) closeThemeModal();
    });
  }

  // Re-layout floating emojis smoothly across orientation/width mode changes
  let lastWideViewport = (window.innerWidth || 390) >= 760;
  let resizeEmojiTimer = null;
  window.addEventListener('resize', () => {
    const isWide = (window.innerWidth || 390) >= 760;
    if (isWide !== lastWideViewport) {
      lastWideViewport = isWide;
      clearTimeout(resizeEmojiTimer);
      resizeEmojiTimer = setTimeout(() => {
        renderFloatingEmojisForTheme(getThemeById(currentTheme));
      }, 200);
    }
  });

  setupDropzone();
  setupSwipeGestures();

  // Auto-check for updates when tab becomes visible or app is unlocked
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible') {
      checkYandexDiskUpdates(false);
    }
  });
  window.addEventListener('focus', () => {
    checkYandexDiskUpdates(false);
  });

  // Background auto-check every 15 minutes
  setInterval(() => {
    checkYandexDiskUpdates(false);
  }, 15 * 60 * 1000);

  // Register PWA Service Worker
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('./service-worker.js')
        .then((reg) => console.log('ServiceWorker registered:', reg.scope))
        .catch((err) => console.log('ServiceWorker registration skipped:', err));
    });
  }
}

// --- Initialize App on Load ---
document.addEventListener('DOMContentLoaded', () => {
  setupEventListeners();
  initApp();
});
