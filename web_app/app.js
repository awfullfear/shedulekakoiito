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
  defaultGroup: 'ЛОМО26/СО',
  yandexPublicKey: 'https://disk.yandex.ru/d/7XafcYMVe1ec7g',
  yandexKeyId: '7XafcYMVe1ec7g',
  corsProxy: 'https://api.allorigins.win/raw?url='
};

let scheduleData = null;
let currentGroup = localStorage.getItem('mabiu_group') || CONFIG.defaultGroup;
let selectedDateStr = getTodayDateStr();
let tickerTimerId = null;

// DOM Elements
const currentGroupLabel = document.getElementById('currentGroupLabel');
const groupSelectorBtn = document.getElementById('groupSelectorBtn');
const refreshBtn = document.getElementById('refreshBtn');
const settingsBtn = document.getElementById('settingsBtn');

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

// --- Data Loading & Persistence ---

async function initApp() {
  currentGroupLabel.textContent = currentGroup;
  if (diskSourceLabel) diskSourceLabel.textContent = CONFIG.yandexKeyId;

  // 1. Try loading cached custom schedule from localStorage
  const cachedJson = localStorage.getItem('mabiu_schedule_data');
  if (cachedJson) {
    try {
      scheduleData = JSON.parse(cachedJson);
    } catch (e) {
      console.warn('Failed to parse cached schedule data, resetting:', e);
      localStorage.removeItem('mabiu_schedule_data');
    }
  }

  // 2. If no valid cache, load bundled data/schedule.json
  if (!scheduleData) {
    try {
      const resp = await fetch('./data/schedule.json');
      if (resp.ok) {
        scheduleData = await resp.json();
      }
    } catch (err) {
      console.error('Failed to load bundled schedule.json:', err);
    }
  }

  // 3. Fallback structure if both failed
  if (!scheduleData) {
    scheduleData = {
      groups: [currentGroup],
      schedules: {},
      fileName: 'default',
      lastUpdated: Date.now()
    };
  }

  // Ensure current group exists or fallback
  if (!scheduleData.groups || !scheduleData.groups.includes(currentGroup)) {
    if (scheduleData.groups && scheduleData.groups.length > 0) {
      // Pick first group ending in 26 or matching LOMO
      const found = scheduleData.groups.find(g => g.includes('26')) || scheduleData.groups[0];
      currentGroup = found;
      localStorage.setItem('mabiu_group', currentGroup);
      currentGroupLabel.textContent = currentGroup;
    }
  }

  updateSettingsLabels();
  setupDaysCarousel();
  renderLessons();
  startLiveTicker();

  // Background check for updates from Yandex.Disk
  checkYandexDiskUpdates(false);
}

function updateSettingsLabels() {
  if (currentFileNameLabel && scheduleData) {
    currentFileNameLabel.textContent = scheduleData.fileName || '26.10-31.10.xlsx';
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
  const groupSchedule = scheduleData.schedules?.[currentGroup];
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

  const groupSchedule = scheduleData?.schedules?.[currentGroup];
  const lessons = groupSchedule?.days?.[selectedDateStr] || [];

  lessonsList.innerHTML = '';

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
        <div class="empty-title">${isWeekend ? 'Выходной день 🌴' : 'Пар нет 🎉'}</div>
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
    `;

    lessonsList.appendChild(card);
  });
}

function findNextDayWithLessons(fromDateStr) {
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

// --- Live Countdown Ticker (Real-Time 1s Tick) ---

function startLiveTicker() {
  if (tickerTimerId) clearInterval(tickerTimerId);
  updateTickerOnce();
  tickerTimerId = setInterval(updateTickerOnce, 1000);
}

function updateTickerOnce() {
  const now = new Date();
  const todayStr = getTodayDateStr();
  const groupSchedule = scheduleData?.schedules?.[currentGroup];
  const todayLessons = groupSchedule?.days?.[todayStr] || [];

  if (!todayLessons || todayLessons.length === 0) {
    // No lessons today. Find next lesson day!
    const nextInfo = findNextDayWithLessons(todayStr);
    statusBadge.textContent = 'Выходной';
    statusBadge.style.color = 'var(--text-muted)';
    statusBadge.style.backgroundColor = 'var(--surface-variant)';
    statusTitle.textContent = 'Сегодня занятий нет 🎉';
    statusSubtitle.textContent = nextInfo || 'Расписание на следующие дни пока не загружено';
    return;
  }

  // Parse lesson intervals for today
  const currentMinutes = now.getHours() * 60 + now.getMinutes();
  const currentSeconds = now.getSeconds();
  const currentTimeInSec = currentMinutes * 60 + currentSeconds;

  const parsedLessons = todayLessons.map((l) => {
    const rawTime = (l.time || '').replace(/\s+/g, '');
    const m = rawTime.match(/(\d{1,2})[.:](\d{2})[-–—](\d{1,2})[.:](\d{2})/);
    let startSec = 0;
    let endSec = 0;
    let startTimeStr = '00:00';
    let endTimeStr = '00:00';

    if (m) {
      const sh = parseInt(m[1], 10);
      const sm = parseInt(m[2], 10);
      const eh = parseInt(m[3], 10);
      const em = parseInt(m[4], 10);
      startSec = (sh * 60 + sm) * 60;
      endSec = (eh * 60 + em) * 60;
      startTimeStr = `${String(sh).padStart(2, '0')}:${String(sm).padStart(2, '0')}`;
      endTimeStr = `${String(eh).padStart(2, '0')}:${String(em).padStart(2, '0')}`;
    }

    return {
      lesson: l,
      startSec,
      endSec,
      startTimeStr,
      endTimeStr
    };
  }).filter(p => p.startSec > 0 && p.endSec > 0);

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

function openGroupModal() {
  groupSearchInput.value = '';
  renderGroupList('');
  groupModal.classList.add('open');
  setTimeout(() => groupSearchInput.focus(), 150);
}

function closeGroupModal() {
  groupModal.classList.remove('open');
}

function renderGroupList(query) {
  groupList.innerHTML = '';
  const q = query.trim().toLowerCase();
  const allGroups = scheduleData?.groups || [currentGroup];

  // Filter 26 cohort groups
  const filtered = allGroups.filter((g) => {
    return g.toLowerCase().includes(q);
  });

  if (filtered.length === 0) {
    groupList.innerHTML = '<div style="text-align: center; padding: 24px; color: var(--text-muted); font-size: 14px;">Группы не найдены</div>';
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
  if (currentGroup === groupName) return;
  currentGroup = groupName;
  localStorage.setItem('mabiu_group', currentGroup);
  currentGroupLabel.textContent = currentGroup;

  setupDaysCarousel();
  renderLessons();
  updateTickerOnce();
  showToast(`Выбрана группа: ${groupName}`);
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

    scheduleData = parsedData;
    localStorage.setItem('mabiu_schedule_data', JSON.stringify(parsedData));

    // Ensure selected group is valid
    if (!scheduleData.groups.includes(currentGroup)) {
      currentGroup = scheduleData.groups[0];
      localStorage.setItem('mabiu_group', currentGroup);
      currentGroupLabel.textContent = currentGroup;
    }

    updateSettingsLabels();
    setupDaysCarousel();
    renderLessons();
    updateTickerOnce();
    closeSettingsModal();
    showToast(`Файл «${file.name}» успешно загружен!`);
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

    for (let c = 0; c <= Math.min(2, range.e.c); c++) {
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
    for (let c = 0; c <= Math.min(2, range.e.c); c++) {
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

  // Extract lesson type
  const typeMatch = text.match(/\(\s*(лек|лекц|лекция|пр|практ|практика|семинар|лаб|зачет|зачёт|экзамен|конс)\.?\s*\)/i);
  if (typeMatch) {
    const t = typeMatch[1].toLowerCase();
    if (t.startsWith('лек')) lessonType = 'Лекция';
    else if (t.startsWith('пр')) lessonType = 'Практическое занятие';
    else if (t.startsWith('сем')) lessonType = 'Семинар';
    else if (t.startsWith('лаб')) lessonType = 'Лабораторная';
    else if (t.startsWith('зач')) lessonType = 'Зачёт';
    else if (t.startsWith('экз')) lessonType = 'Экзамен';
    else lessonType = typeMatch[1];
  }

  // Extract room
  const roomMatch = text.match(/(?:,\s*)?(?:(?:Ауд\.?|АУд\.?)\s*[\w\d/]*\s*(?:\([^)]+\))?|(?:Актовый\s*зал|Конференц[- ]*зал|Спорт\.?\s*зал|Он[- ]*лайн|ДО)(?:\s*\([^)]+\))?|\((?:Конференц[- ]*зал|Актовый\s*зал)\))\s*$/i);
  if (roomMatch) {
    room = roomMatch[0].replace(/^,\s*/, '').trim();
    text = text.substring(0, text.length - roomMatch[0].length).trim();
  }

  // Extract teacher
  const teacherMatch = text.match(/([А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ]\.+\s*[А-ЯЁ]\.*|\s+[А-ЯЁ][а-яё]+))\s*$/);
  if (teacherMatch) {
    teacher = teacherMatch[1].trim();
    text = text.substring(0, text.length - teacherMatch[0].length).trim();
  }

  // Subject is cleaned remainder
  let cleanSubject = text.replace(/^[,.\s-]+|[,.\s-]+$/g, '').trim();
  if (!cleanSubject) cleanSubject = rawText;

  return { cleanSubject, room, teacher, lessonType };
}

// --- Sync with Yandex.Disk ---

async function checkYandexDiskUpdates(isManual = false) {
  if (refreshBtn) refreshBtn.classList.add('spinning');
  if (isManual) showToast('Проверка Яндекс.Диска...');

  try {
    // 1. First try Vercel Serverless API (/api/schedule)
    let fetchedData = null;
    try {
      const resp = await fetch('./api/schedule');
      if (resp.ok) {
        fetchedData = await resp.json();
      }
    } catch (e) {
      // Local or static GitHub Pages fallback
    }

    // 2. Fallback to direct Yandex API via public CORS proxy if /api/schedule not available
    if (!fetchedData) {
      const yandexApi = `https://cloud-api.yandex.net/v1/disk/public/resources?public_key=${encodeURIComponent(CONFIG.yandexPublicKey)}&limit=100`;
      const proxyUrl = CONFIG.corsProxy + encodeURIComponent(yandexApi);
      const resp = await fetch(proxyUrl);
      if (resp.ok) {
        const rawJson = await resp.json();
        const diskJson = typeof rawJson.contents === 'string' ? JSON.parse(rawJson.contents) : rawJson;

        const items = diskJson._embedded?.items || [];
        const xlsxFiles = items.filter(it => it.name && it.name.match(/\.(xlsx|xls)$/i) && it.file);
        if (xlsxFiles.length > 0) {
          // Sort latest file
          xlsxFiles.sort((a, b) => new Date(b.modified || 0) - new Date(a.modified || 0));
          const latestFile = xlsxFiles[0];

          // Download latest xlsx
          const fileResp = await fetch(CONFIG.corsProxy + encodeURIComponent(latestFile.file));
          if (fileResp.ok) {
            const rawBase64OrText = await fileResp.text();
            // Convert to buffer or blob
            const arrayBuffer = await fileResp.arrayBuffer();
            fetchedData = parseWorkbookClient(arrayBuffer, latestFile.name);
          }
        }
      }
    }

    if (fetchedData && fetchedData.groups && fetchedData.groups.length > 0) {
      scheduleData = fetchedData;
      localStorage.setItem('mabiu_schedule_data', JSON.stringify(fetchedData));
      updateSettingsLabels();
      setupDaysCarousel();
      renderLessons();
      updateTickerOnce();

      if (isManual) showToast('Расписание успешно обновлено с Яндекс.Диска!');
    } else {
      if (isManual) showToast('Расписание актуально.');
    }
  } catch (err) {
    console.warn('Sync attempt finished:', err);
    if (isManual) showToast('Синхронизация завершена. Используются актуальные данные.');
  } finally {
    if (refreshBtn) refreshBtn.classList.remove('spinning');
  }
}

// Reset schedule to bundled defaults
async function resetSchedule() {
  localStorage.removeItem('mabiu_schedule_data');
  showToast('Сброс к исходному расписанию...');
  try {
    const resp = await fetch('./data/schedule.json');
    if (resp.ok) {
      scheduleData = await resp.json();
      updateSettingsLabels();
      setupDaysCarousel();
      renderLessons();
      updateTickerOnce();
      closeSettingsModal();
      showToast('Расписание успешно сброшено!');
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

  // Action buttons
  refreshBtn.addEventListener('click', () => checkYandexDiskUpdates(true));
  syncNowBtn.addEventListener('click', () => checkYandexDiskUpdates(true));
  resetScheduleBtn.addEventListener('click', resetSchedule);

  todayJumpBtn.addEventListener('click', () => {
    selectDay(getTodayDateStr());
  });

  setupDropzone();
  setupSwipeGestures();

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
