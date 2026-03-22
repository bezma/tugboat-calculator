const STORAGE_PREFIX = 'standalone_tug_service:';
const STORAGE_KEYS = {
  vesselName: 'vessel_name',
  gt: 'gt',
  globalImoTransport: 'global_imo_transport',
  mobileHomeCompact: 'mobile_home_compact',
  towageTotal: 'towage_total',
  towageArrivalCount: 'towage_arrival_count',
  towageDepartureCount: 'towage_departure_count',
  tugsState: 'tugs_state'
};

const ICON_PRESS_ANIMATION_CLASS = 'icon-animating';
const ICON_PRESS_ANIMATION_MS = 176;
const iconPressAnimationTimers = new WeakMap();

let printRestoreTitle = null;
let mobilePrintPageStyle = null;
let tugCount = 0;
let isRestoringTugs = false;
let isResettingTugs = false;

const MIN_VOYAGE = 1;
const MIN_ASSIST = 0.5;
const TUG_OT_RATE_NORMAL = '0';
const TUG_OT_RATE_25 = '0.25';
const TUG_OT_RATE_50 = '0.50';
const TUG_STANDBY_RATE = 168;
const TUG_STANDBY_MIN_GT = 10000;

let imoMaster = null;
let linesMaster = null;
let arrivalOtMaster = null;
let arrivalOtSunday = null;
let departureOtMaster = null;
let departureOtSunday = null;
let discountEnabled = null;
let discountPercentInput = null;
let discountFieldWrap = null;
let standbyEnabled = null;
let standbyHoursInput = null;
let standbyWrap = null;
let standbyCard = null;
let standbyTotalEl = null;
let standbyBreakdownEl = null;
let tugOptionsToggleEl = null;
let tugOptionsBodyEl = null;
let tugOptionsExpanded = false;

function getScopedStorageKey(key) {
  return `${STORAGE_PREFIX}${key}`;
}

function safeStorageGet(key) {
  try {
    return localStorage.getItem(getScopedStorageKey(key));
  } catch (error) {
    return null;
  }
}

function safeStorageSet(key, value) {
  try {
    localStorage.setItem(getScopedStorageKey(key), String(value));
  } catch (error) {
    // ignore storage errors
  }
}

function safeStorageRemove(key) {
  try {
    localStorage.removeItem(getScopedStorageKey(key));
  } catch (error) {
    // ignore storage errors
  }
}

function parseStoredBoolean(raw) {
  if (raw === null || raw === undefined) return null;
  const normalized = String(raw).trim().toLowerCase();
  if (normalized === '1' || normalized === 'true' || normalized === 'yes' || normalized === 'on') return true;
  if (normalized === '0' || normalized === 'false' || normalized === 'no' || normalized === 'off') return false;
  return null;
}

function getGlobalImoTransportState() {
  return parseStoredBoolean(safeStorageGet(STORAGE_KEYS.globalImoTransport));
}

function setGlobalImoTransportState(checked) {
  safeStorageSet(STORAGE_KEYS.globalImoTransport, checked ? '1' : '0');
}

function getStoredMobileHomeCompactState() {
  return parseStoredBoolean(safeStorageGet(STORAGE_KEYS.mobileHomeCompact)) === true;
}

function setStoredMobileHomeCompactState(compact) {
  safeStorageSet(STORAGE_KEYS.mobileHomeCompact, compact ? '1' : '0');
}

function parseMoneyValue(raw) {
  if (!raw) return 0;
  const cleaned = String(raw).trim().replace(/[^0-9,.-]/g, '');
  if (!cleaned) return 0;

  const hasComma = cleaned.includes(',');
  const hasDot = cleaned.includes('.');

  let normalized = cleaned;
  if (hasComma && hasDot) {
    const lastComma = cleaned.lastIndexOf(',');
    const lastDot = cleaned.lastIndexOf('.');
    if (lastComma > lastDot) {
      normalized = cleaned.replace(/\./g, '').replace(/,/g, '.');
    } else {
      normalized = cleaned.replace(/,/g, '');
    }
  } else if (hasComma && !hasDot) {
    if (/^-?\d{1,3}(,\d{3})+$/.test(cleaned)) {
      normalized = cleaned.replace(/,/g, '');
    } else {
      normalized = cleaned.replace(/,/g, '.');
    }
  } else if (hasDot && !hasComma) {
    if (/^-?\d{1,3}(\.\d{3})+$/.test(cleaned)) {
      normalized = cleaned.replace(/\./g, '');
    }
  }

  const value = Number(normalized);
  return Number.isFinite(value) ? value : 0;
}

function parsePercentValue(raw) {
  return parseMoneyValue(raw);
}

function formatMoneyValue(value) {
  const fixed = Number.isFinite(value) ? value.toFixed(2) : '0.00';
  const [whole, decimal] = fixed.split('.');
  const withThousands = whole.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return `${withThousands},${decimal}`;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function isLikelyMobileViewport() {
  const ua = navigator && navigator.userAgent ? navigator.userAgent : '';
  if (/Android|webOS|iPhone|iPad|iPod|Mobile|CriOS|FxiOS/i.test(ua)) return true;
  const viewportWidth = window.visualViewport ? window.visualViewport.width : window.innerWidth;
  if (Number.isFinite(viewportWidth) && viewportWidth > 0 && viewportWidth <= 900) return true;
  const screenWidth = window.screen && Number.isFinite(window.screen.width) ? window.screen.width : 0;
  if (screenWidth > 0 && screenWidth <= 900) return true;
  return window.matchMedia('(max-width: 900px), (pointer: coarse)').matches;
}

function enableMobilePrintFooterSuppression() {
  if (mobilePrintPageStyle) return;
  const style = document.createElement('style');
  style.id = 'mobilePrintPageStyle';
  style.textContent = [
    '@page { margin: 0 !important; size: A4; }',
    'body.page-tugs { padding: 10mm 6mm 10mm 6mm !important; }'
  ].join('\n');
  document.head.appendChild(style);
  mobilePrintPageStyle = style;
}

function disableMobilePrintFooterSuppression() {
  if (!mobilePrintPageStyle) return;
  mobilePrintPageStyle.remove();
  mobilePrintPageStyle = null;
}

function getVesselNameForPrint() {
  const stored = safeStorageGet(STORAGE_KEYS.vesselName);
  if (stored) return String(stored).trim();
  const tugField = document.getElementById('vesselName');
  if (tugField && tugField.value) return String(tugField.value).trim();
  return '';
}

function setPrintTitleFromVessel() {
  if (printRestoreTitle === null) {
    printRestoreTitle = document.title;
  }
  const vesselName = getVesselNameForPrint();
  document.title = vesselName ? `Tug Service Calculator ${vesselName}` : 'Tug Service Calculator';
}

function restorePrintTitleIfNeeded() {
  if (printRestoreTitle === null) return;
  document.title = printRestoreTitle;
  printRestoreTitle = null;
}

function syncPrintSummaryValues() {
  const vesselInput = document.getElementById('vesselName');
  const gtInput = document.getElementById('gt');
  const tariffInput = document.getElementById('tariff');
  const vesselPrint = document.getElementById('vesselNamePrint');
  const gtPrint = document.getElementById('gtPrint');
  const tariffPrint = document.getElementById('tariffPrint');

  if (vesselPrint) {
    const vesselValue = vesselInput && vesselInput.value && vesselInput.value.trim()
      ? vesselInput.value.trim()
      : (vesselInput && vesselInput.placeholder ? vesselInput.placeholder : '');
    vesselPrint.textContent = vesselValue;
  }
  if (gtPrint) gtPrint.textContent = gtInput && gtInput.value ? gtInput.value : '';
  if (tariffPrint) tariffPrint.textContent = tariffInput && tariffInput.value ? tariffInput.value : '';
}

function printWithTitle() {
  syncPrintSummaryValues();
  setPrintTitleFromVessel();
  window.requestAnimationFrame(() => {
    window.setTimeout(() => {
      window.print();
    }, 60);
  });
  window.setTimeout(() => {
    restorePrintTitleIfNeeded();
  }, 1500);
}

function isAnimatedIconButton(button) {
  if (!button || !(button instanceof HTMLElement)) return false;
  return button.matches('.icon-btn');
}

function triggerIconPressAnimation(button) {
  if (!button || !isAnimatedIconButton(button) || button.disabled) return;
  const existingTimer = iconPressAnimationTimers.get(button);
  if (existingTimer) window.clearTimeout(existingTimer);

  button.classList.remove(ICON_PRESS_ANIMATION_CLASS);
  void button.offsetWidth;
  button.classList.add(ICON_PRESS_ANIMATION_CLASS);

  const nextTimer = window.setTimeout(() => {
    button.classList.remove(ICON_PRESS_ANIMATION_CLASS);
    iconPressAnimationTimers.delete(button);
  }, ICON_PRESS_ANIMATION_MS);
  iconPressAnimationTimers.set(button, nextTimer);
}

function clearTransientIconAnimationState(root = document) {
  if (!root || !root.querySelectorAll) return;
  root
    .querySelectorAll(`button.${ICON_PRESS_ANIMATION_CLASS}`)
    .forEach((button) => button.classList.remove(ICON_PRESS_ANIMATION_CLASS));
}

function initIconClickAnimationCompletion() {
  document.addEventListener('pointerdown', (event) => {
    if (!event.isTrusted || event.button !== 0) return;
    const button = event.target && event.target.closest ? event.target.closest('button') : null;
    if (!button) return;
    triggerIconPressAnimation(button);
  }, true);

  document.addEventListener('keydown', (event) => {
    if (!event.isTrusted) return;
    if (event.key !== 'Enter' && event.key !== ' ') return;
    const button = event.target && event.target.closest ? event.target.closest('button') : null;
    if (!button) return;
    triggerIconPressAnimation(button);
  }, true);
}

function updateImoToggleLabelColor(input) {
  if (!input) return;
  const toggleLabel = input.closest('label');
  if (!toggleLabel) return;
  if (!toggleLabel.classList.contains('imo-master-label')) return;
  toggleLabel.classList.toggle('is-checked', Boolean(input.checked));
}

function updateVesselNameFromStorage(targetInput) {
  if (!targetInput) return;
  const storedName = safeStorageGet(STORAGE_KEYS.vesselName);
  if (!storedName) return;
  if (targetInput.value !== storedName) {
    targetInput.value = storedName;
  }
}

function getTariffFromGT(gt) {
  if (!gt || gt <= 0) return 0;
  if (gt <= 2000) return 698;
  if (gt <= 5000) return 908;
  if (gt <= 10000) return 1133;
  if (gt <= 15000) return 1294;
  if (gt <= 25000) return 1358;
  if (gt <= 30000) return 1646;
  if (gt <= 35000) return 1863;
  const extraThousands = Math.ceil((gt - 35000) / 1000);
  return 1863 + (extraThousands * 16);
}

function getTugServiceCards(root = document.getElementById('tugCards')) {
  if (!root) return [];
  return Array.from(root.querySelectorAll('.tug-service-card'));
}

function saveTugsState() {
  if (isRestoringTugs || isResettingTugs) return;
  const tugCards = document.getElementById('tugCards');
  if (!tugCards) return;

  const tugs = getTugServiceCards(tugCards).map((card) => {
    const id = card.id.split('_')[1];
    return {
      op: document.getElementById(`op_${id}`)?.value || 'arrival',
      voyage: document.getElementById(`voyage_${id}`)?.value || '1',
      assist: document.getElementById(`assist_${id}`)?.value || '0.5',
      imo: document.getElementById(`imo_${id}`)?.checked || false,
      lines: document.getElementById(`lines_${id}`)?.checked || false,
      kw: document.getElementById(`kw_${id}`)?.checked || false,
      voy_ot: document.getElementById(`voy_ot_${id}`)?.value || '0',
      assist_ot: document.getElementById(`assist_ot_${id}`)?.value || '0'
    };
  });

  const state = {
    vesselName: document.getElementById('vesselName')?.value || '',
    gt: document.getElementById('gt')?.value || '',
    imoMaster: document.getElementById('imoMaster')?.checked || false,
    linesMaster: document.getElementById('linesMaster')?.checked || false,
    arrivalOtMaster: document.getElementById('arrivalOtMaster')?.checked || false,
    arrivalOtSunday: document.getElementById('arrivalOtSunday')?.checked || false,
    departureOtMaster: document.getElementById('departureOtMaster')?.checked || false,
    departureOtSunday: document.getElementById('departureOtSunday')?.checked || false,
    discountEnabled: document.getElementById('discountEnabled')?.checked || false,
    discountPercent: document.getElementById('discountPercent')?.value || '',
    standbyEnabled: document.getElementById('standbyEnabled')?.checked || false,
    standbyHours: document.getElementById('standbyHours')?.value || '',
    tugs
  };

  safeStorageSet(STORAGE_KEYS.tugsState, JSON.stringify(state));
}

function restoreTugsState() {
  const tugCards = document.getElementById('tugCards');
  if (!tugCards) return false;

  const raw = safeStorageGet(STORAGE_KEYS.tugsState);
  if (!raw) return false;

  let state = null;
  try {
    state = JSON.parse(raw);
  } catch (error) {
    return false;
  }
  if (!state || !Array.isArray(state.tugs)) return false;

  isRestoringTugs = true;
  getTugServiceCards(tugCards).forEach((card) => card.remove());
  tugCount = 0;

  state.tugs.forEach((tug) => {
    addTug();
    const id = tugCount;
    const op = document.getElementById(`op_${id}`);
    const voyage = document.getElementById(`voyage_${id}`);
    const assist = document.getElementById(`assist_${id}`);
    const imo = document.getElementById(`imo_${id}`);
    const lines = document.getElementById(`lines_${id}`);
    const kw = document.getElementById(`kw_${id}`);
    const voyOt = document.getElementById(`voy_ot_${id}`);
    const assistOt = document.getElementById(`assist_ot_${id}`);

    if (op) op.value = tug.op || 'arrival';
    if (voyage) voyage.value = tug.voyage || '1';
    if (assist) assist.value = tug.assist || '0.5';
    if (imo) imo.checked = Boolean(tug.imo);
    if (lines) lines.checked = Boolean(tug.lines);
    if (kw) kw.checked = Boolean(tug.kw);
    if (voyOt) voyOt.value = tug.voy_ot || '0';
    if (assistOt) assistOt.value = tug.assist_ot || '0';
    if (kw && kw.checked && voyage) voyage.value = '1.5';
  });

  const vesselName = document.getElementById('vesselName');
  if (vesselName) vesselName.value = state.vesselName || '';

  const gtInput = document.getElementById('gt');
  if (gtInput) {
    gtInput.value = state.gt || '';
    const tariffInput = document.getElementById('tariff');
    if (tariffInput) {
      tariffInput.value = getTariffFromGT(Number(gtInput.value)) || '';
    }
  }

  if (imoMaster) imoMaster.checked = Boolean(state.imoMaster);
  if (linesMaster) linesMaster.checked = Boolean(state.linesMaster);
  if (arrivalOtMaster) arrivalOtMaster.checked = Boolean(state.arrivalOtMaster);
  if (arrivalOtSunday) arrivalOtSunday.checked = Boolean(state.arrivalOtSunday);
  if (departureOtMaster) departureOtMaster.checked = Boolean(state.departureOtMaster);
  if (departureOtSunday) departureOtSunday.checked = Boolean(state.departureOtSunday);
  if (discountEnabled) discountEnabled.checked = Boolean(state.discountEnabled);
  if (discountPercentInput) discountPercentInput.value = state.discountPercent == null ? '' : String(state.discountPercent);
  if (standbyEnabled) standbyEnabled.checked = Boolean(state.standbyEnabled);
  if (standbyHoursInput) standbyHoursInput.value = state.standbyHours == null ? '' : String(state.standbyHours);

  isRestoringTugs = false;
  updateTugTitles();
  syncImoMaster();
  syncLinesMaster();
  updateOvertimeMasterVisibility();
  updateDiscountEnabledState();
  updateTugStandbyVisibility();
  calculate();
  saveTugsState();
  return true;
}

function addTug() {
  const tugCards = document.getElementById('tugCards');
  if (!tugCards) return;
  tugCount += 1;

  const markup = `
    <div class="card tug-service-card" id="tug_${tugCount}" data-operation="arrival">
      <div class="tug-header">
        <button class="icon-btn" onclick="removeTug(${tugCount})" aria-label="Remove tugboat"><img src="assets/icons/trash.svg" alt="Remove"></button>
        <h3 class="tug-title">Tugboat</h3>
        <button class="icon-btn" onclick="duplicateTug(${tugCount})" aria-label="Duplicate tugboat"><img src="assets/icons/copy-plus.svg" alt="Duplicate"></button>
      </div>

      <label>Operation</label>
      <select id="op_${tugCount}">
        <option value="arrival">Arrival (IN)</option>
        <option value="departure">Departure (OUT)</option>
      </select>

      <label>Voyage Time</label>
      <select id="voyage_${tugCount}">
        <option value="1">1 hour</option>
        <option value="1.5">1.5 hours</option>
      </select>

      <label>Assistance Time</label>
      <select id="assist_${tugCount}">
        <option value="0.5">Up to 30 min</option>
        <option value="1">Within 1 hour</option>
        <option value="1.5">Within 1.5 hours</option>
      </select>

      <div class="section-title">Assistance Surcharges</div>
      <label class="checkbox"><input type="checkbox" id="imo_${tugCount}" /> 20% IMO Code Classes I-III</label>
      <label class="checkbox"><input type="checkbox" id="lines_${tugCount}" /> 15% Usage of tug lines</label>
      <label class="checkbox"><input type="checkbox" id="kw_${tugCount}" /> 30% Tug above 2000 kW</label>

      <div class="section-title">Overtime</div>
      <label>Voyage Overtime</label>
      <select id="voy_ot_${tugCount}">
        <option value="0">Normal working hours</option>
        <option value="0.25">25% (Mon–Sat)</option>
        <option value="0.50">50% (Sun & Holidays)</option>
      </select>

      <label>Assistance Overtime</label>
      <select id="assist_ot_${tugCount}">
        <option value="0">Normal working hours</option>
        <option value="0.25">25% (Mon–Sat)</option>
        <option value="0.50">50% (Sun & Holidays)</option>
      </select>

      <div class="tug-total" id="tugTotal_${tugCount}">Tug total: €0.00</div>
      <div class="tug-breakdown-print" id="tugBreakdown_${tugCount}"></div>
    </div>
  `;

  if (standbyCard && standbyCard.parentElement === tugCards) {
    standbyCard.insertAdjacentHTML('beforebegin', markup);
  } else {
    tugCards.insertAdjacentHTML('beforeend', markup);
  }

  if (isRestoringTugs) return;

  applyImoMaster();
  applyLinesMaster();
  syncImoMaster();
  syncLinesMaster();
  applyOvertimeMastersToCard(tugCount);
  updateTugTitles();
  calculate();
}

function removeTug(id) {
  document.getElementById(`tug_${id}`)?.remove();
  syncImoMaster();
  syncLinesMaster();
  updateTugTitles();
  calculate();
}

function duplicateTug(id) {
  const source = document.getElementById(`tug_${id}`);
  if (!source) return;

  addTug();
  const newId = tugCount;
  const fields = [
    ['op', 'value'],
    ['voyage', 'value'],
    ['assist', 'value'],
    ['imo', 'checked'],
    ['lines', 'checked'],
    ['kw', 'checked'],
    ['voy_ot', 'value'],
    ['assist_ot', 'value']
  ];

  fields.forEach(([key, prop]) => {
    const from = document.getElementById(`${key}_${id}`);
    const to = document.getElementById(`${key}_${newId}`);
    if (!from || !to) return;
    to[prop] = from[prop];
  });

  applyKwVoyageRule(newId);
  syncImoMaster();
  syncLinesMaster();
  updateTugTitles();
  calculate();
}

function duplicateLastTug() {
  const selectedId = getSelectedTugId();
  if (selectedId) {
    duplicateTug(selectedId);
    return;
  }
  const cards = getTugServiceCards();
  if (cards.length === 0) {
    addTug();
    return;
  }
  const last = cards[cards.length - 1];
  const id = last.id.split('_')[1];
  duplicateTug(Number(id));
}

function setSelectedTug(id) {
  getTugServiceCards().forEach((card) => card.classList.remove('selected'));
  const target = document.getElementById(`tug_${id}`);
  if (target) target.classList.add('selected');
}

function getSelectedTugId() {
  const selected = document.querySelector('#tugCards .tug-service-card.selected');
  if (!selected) return null;
  const parts = selected.id.split('_');
  return parts.length > 1 ? Number(parts[1]) : null;
}

function applyImoMaster() {
  if (!imoMaster) return;
  const checked = imoMaster.checked;
  imoMaster.indeterminate = false;
  for (let i = 1; i <= tugCount; i += 1) {
    const cb = document.getElementById(`imo_${i}`);
    if (cb) cb.checked = checked;
  }
}

function applyLinesMaster() {
  if (!linesMaster) return;
  const checked = linesMaster.checked;
  linesMaster.indeterminate = false;
  for (let i = 1; i <= tugCount; i += 1) {
    const cb = document.getElementById(`lines_${i}`);
    if (cb) cb.checked = checked;
  }
}

function syncImoMaster() {
  if (!imoMaster) return;
  let total = 0;
  let checked = 0;
  for (let i = 1; i <= tugCount; i += 1) {
    const cb = document.getElementById(`imo_${i}`);
    if (!cb) continue;
    total += 1;
    if (cb.checked) checked += 1;
  }
  if (total === 0) {
    imoMaster.indeterminate = false;
    updateImoToggleLabelColor(imoMaster);
    return;
  }
  if (checked === 0) {
    imoMaster.checked = false;
    imoMaster.indeterminate = false;
  } else if (checked === total) {
    imoMaster.checked = true;
    imoMaster.indeterminate = false;
  } else {
    imoMaster.checked = false;
    imoMaster.indeterminate = true;
  }
  updateImoToggleLabelColor(imoMaster);
}

function syncTugImoWithGlobalState(forceInitialize = false) {
  if (!imoMaster) return;
  const globalState = getGlobalImoTransportState();
  if (globalState === null) {
    if (forceInitialize && !imoMaster.indeterminate) {
      setGlobalImoTransportState(Boolean(imoMaster.checked));
    }
    updateImoToggleLabelColor(imoMaster);
    return;
  }

  if (imoMaster.checked === globalState && !imoMaster.indeterminate) return;
  imoMaster.checked = globalState;
  imoMaster.indeterminate = false;
  applyImoMaster();
  syncImoMaster();
  updateImoToggleLabelColor(imoMaster);
}

function syncLinesMaster() {
  if (!linesMaster) return;
  let total = 0;
  let checked = 0;
  for (let i = 1; i <= tugCount; i += 1) {
    const cb = document.getElementById(`lines_${i}`);
    if (!cb) continue;
    total += 1;
    if (cb.checked) checked += 1;
  }
  if (total === 0) {
    linesMaster.checked = false;
    linesMaster.indeterminate = false;
    return;
  }
  if (checked === 0) {
    linesMaster.checked = false;
    linesMaster.indeterminate = false;
  } else if (checked === total) {
    linesMaster.checked = true;
    linesMaster.indeterminate = false;
  } else {
    linesMaster.checked = false;
    linesMaster.indeterminate = true;
  }
}

function getMasterOvertimeRate(masterInput, sundayInput) {
  if (!masterInput || !masterInput.checked) return TUG_OT_RATE_NORMAL;
  if (sundayInput && sundayInput.checked) return TUG_OT_RATE_50;
  return TUG_OT_RATE_25;
}

function applyOvertimeMaster(operation, rate, onlyId = null) {
  for (let i = 1; i <= tugCount; i += 1) {
    if (onlyId && i !== onlyId) continue;
    const opValue = document.getElementById(`op_${i}`)?.value;
    if (!opValue) continue;
    const op = opValue === 'departure' ? 'departure' : 'arrival';
    if (op !== operation) continue;
    const voyOt = document.getElementById(`voy_ot_${i}`);
    const assistOt = document.getElementById(`assist_ot_${i}`);
    if (voyOt) voyOt.value = rate;
    if (assistOt) assistOt.value = rate;
  }
}

function applyOvertimeMastersToCard(id) {
  const opValue = document.getElementById(`op_${id}`)?.value;
  if (!opValue) return;
  const op = opValue === 'departure' ? 'departure' : 'arrival';
  if (op === 'departure') {
    applyOvertimeMaster('departure', getMasterOvertimeRate(departureOtMaster, departureOtSunday), id);
    return;
  }
  applyOvertimeMaster('arrival', getMasterOvertimeRate(arrivalOtMaster, arrivalOtSunday), id);
}

function applyKwVoyageRule(id) {
  const kwCheckbox = document.getElementById(`kw_${id}`);
  const voyageSelect = document.getElementById(`voyage_${id}`);
  if (!voyageSelect) return;
  voyageSelect.value = kwCheckbox && kwCheckbox.checked ? '1.5' : '1';
}

function updateOvertimeMasterVisibility() {
  const arrivalWrap = document.getElementById('arrivalOtSundayWrap');
  if (arrivalWrap) arrivalWrap.hidden = !arrivalOtMaster?.checked;
  if (!arrivalOtMaster?.checked && arrivalOtSunday) arrivalOtSunday.checked = false;
  const departureWrap = document.getElementById('departureOtSundayWrap');
  if (departureWrap) departureWrap.hidden = !departureOtMaster?.checked;
  if (!departureOtMaster?.checked && departureOtSunday) departureOtSunday.checked = false;
  syncTopHomeBarOffset();
}

function updateDiscountEnabledState() {
  if (!discountPercentInput) return;
  const enabled = Boolean(discountEnabled && discountEnabled.checked);
  discountPercentInput.disabled = !enabled;
  if (discountFieldWrap) discountFieldWrap.hidden = !enabled;
  syncTopHomeBarOffset();
}

function getApprovedDiscountPercent() {
  if (!discountEnabled || !discountEnabled.checked) return 0;
  const percent = parsePercentValue(discountPercentInput ? discountPercentInput.value : '');
  if (!Number.isFinite(percent) || percent <= 0) return 0;
  return Math.min(percent, 100);
}

function formatDiscountLabel(percent) {
  if (!Number.isFinite(percent) || percent <= 0) return '';
  let text = (Math.round(percent * 1000) / 1000).toFixed(3);
  text = text.replace(/\.?0+$/, '');
  return text;
}

function getTugStandbyCharge(hoursRaw, gtValue) {
  const gt = Number(gtValue);
  if (!Number.isFinite(gt) || gt <= TUG_STANDBY_MIN_GT) return 0;
  const hours = parseMoneyValue(hoursRaw);
  if (!Number.isFinite(hours) || hours <= 0) return 0;
  return Math.ceil(hours) * TUG_STANDBY_RATE;
}

function updateTugStandbyVisibility() {
  if (!standbyWrap || !standbyEnabled) return;
  const gtValue = Number(document.getElementById('gt')?.value);
  const isEligible = Number.isFinite(gtValue) && gtValue > TUG_STANDBY_MIN_GT;
  standbyWrap.hidden = !isEligible;
  if (standbyCard) standbyCard.hidden = !isEligible || !standbyEnabled.checked;
  if (standbyHoursInput) standbyHoursInput.disabled = !isEligible || !standbyEnabled.checked;
  syncTopHomeBarOffset();
}

function updateTugSummaryDiscountNote(baseTariff, discountPercent) {
  const note = document.getElementById('tugSummaryDiscountNote');
  if (!note) return;
  const safeTariff = Number(baseTariff);
  const safePercent = Number(discountPercent);
  if (!Number.isFinite(safeTariff) || safeTariff <= 0 || !Number.isFinite(safePercent) || safePercent <= 0) {
    note.hidden = true;
    note.textContent = '';
    return;
  }

  const effectiveTariff = safeTariff * (1 - (safePercent / 100));
  note.textContent = `Approved discount: ${formatDiscountLabel(safePercent)}% | Effective tariff: €${formatMoneyValue(effectiveTariff)} / hour`;
  note.hidden = false;
}

function getOrdinal(n) {
  if (n % 10 === 1 && n % 100 !== 11) return 'st';
  if (n % 10 === 2 && n % 100 !== 12) return 'nd';
  if (n % 10 === 3 && n % 100 !== 13) return 'rd';
  return 'th';
}

function updateTugTitles() {
  const cards = getTugServiceCards();
  const operations = cards.map((card) => {
    const id = card.id.split('_')[1];
    return document.getElementById(`op_${id}`)?.value === 'departure' ? 'departure' : 'arrival';
  });
  const operationCounts = operations.reduce((acc, operation) => {
    acc[operation] = (acc[operation] || 0) + 1;
    return acc;
  }, {});

  let arrivalIndex = 0;
  let departureIndex = 0;

  cards.forEach((card) => {
    const id = card.id.split('_')[1];
    const op = document.getElementById(`op_${id}`)?.value;
    card.setAttribute('data-operation', op === 'departure' ? 'departure' : 'arrival');
    const title = card.querySelector('.tug-title');
    if (!title) return;

    if (op === 'arrival') {
      arrivalIndex += 1;
      card.style.setProperty('--print-row', String(arrivalIndex));
      const prefix = operationCounts.arrival > 1 ? `${arrivalIndex}${getOrdinal(arrivalIndex)} ` : '';
      title.innerText = `${prefix}Tugboat on arrival`;
    } else {
      departureIndex += 1;
      card.style.setProperty('--print-row', String(departureIndex));
      const prefix = operationCounts.departure > 1 ? `${departureIndex}${getOrdinal(departureIndex)} ` : '';
      title.innerText = `${prefix}Tugboat on departure`;
    }
  });
}

function getSelectedOptionText(selectId) {
  const select = document.getElementById(selectId);
  if (!select) return '';
  const option = select.options[select.selectedIndex];
  return option ? String(option.textContent || '').trim() : '';
}

function getTugChargeBreakdown(source, tariff) {
  const voyage = Math.max(Number(source?.voyage) || 0, MIN_VOYAGE);
  const assist = Math.max(Number(source?.assist) || 0, MIN_ASSIST);
  const voyageOT = Number(source?.voyageOT ?? source?.voy_ot) || 0;
  const assistOT = Number(source?.assistOT ?? source?.assist_ot) || 0;

  const voyageBase = tariff * voyage;
  const voyageOvertimeCharge = voyageBase * voyageOT;
  const voyageSubtotal = voyageBase + voyageOvertimeCharge;

  const assistBase = tariff * assist;
  const assistOvertimeCharge = assistBase * assistOT;
  const assistSubtotal = assistBase + assistOvertimeCharge;

  const imoCharge = source?.imo ? assistSubtotal * 0.20 : 0;
  const linesCharge = source?.lines ? assistSubtotal * 0.15 : 0;
  const kwCharge = source?.kw ? assistSubtotal * 0.30 : 0;

  return {
    voyage,
    assist,
    voyageOT,
    assistOT,
    voyageBase,
    voyageOvertimeCharge,
    voyageSubtotal,
    assistBase,
    assistOvertimeCharge,
    assistSubtotal,
    imoCharge,
    linesCharge,
    kwCharge,
    total: voyageSubtotal + assistSubtotal + imoCharge + linesCharge + kwCharge
  };
}

function getTugCardCalculation(id, tariff) {
  const opSelect = document.getElementById(`op_${id}`);
  if (!opSelect) return null;

  const operation = opSelect.value === 'departure' ? 'departure' : 'arrival';
  const calculation = getTugChargeBreakdown({
    voyage: document.getElementById(`voyage_${id}`)?.value,
    assist: document.getElementById(`assist_${id}`)?.value,
    voyageOT: document.getElementById(`voy_ot_${id}`)?.value,
    assistOT: document.getElementById(`assist_ot_${id}`)?.value,
    imo: document.getElementById(`imo_${id}`)?.checked,
    lines: document.getElementById(`lines_${id}`)?.checked,
    kw: document.getElementById(`kw_${id}`)?.checked
  }, tariff);

  return { id, operation, ...calculation };
}

function renderTugPrintBreakdown(id, calc) {
  const container = document.getElementById(`tugBreakdown_${id}`);
  if (!container) return;
  if (!calc || !Number.isFinite(calc.total) || calc.total <= 0) {
    container.innerHTML = '';
    return;
  }

  const card = document.getElementById(`tug_${id}`);
  const cardTitle = card?.querySelector('.tug-title')?.textContent?.trim() || `Tugboat ${id}`;
  const rows = [
    { amount: calc.voyageBase, label: `Voyage base to base (${getSelectedOptionText(`voyage_${id}`) || `${calc.voyage} hour`})` },
    { amount: calc.assistBase, label: `Assistance (${getSelectedOptionText(`assist_${id}`) || `${calc.assist} hour`})` }
  ];

  if (calc.voyageOvertimeCharge > 0) rows.push({ amount: calc.voyageOvertimeCharge, label: `Voyage overtime (${getSelectedOptionText(`voy_ot_${id}`)})` });
  if (calc.assistOvertimeCharge > 0) rows.push({ amount: calc.assistOvertimeCharge, label: `Assistance overtime (${getSelectedOptionText(`assist_ot_${id}`)})` });
  if (calc.linesCharge > 0) rows.push({ amount: calc.linesCharge, label: '15% usage of tug lines' });
  if (calc.imoCharge > 0) rows.push({ amount: calc.imoCharge, label: '20% towing of ships transporting inflammable materials (IMO Code Classes I-III)' });
  if (calc.kwCharge > 0) rows.push({ amount: calc.kwCharge, label: '30% for the work of the tug above 2000 kW' });

  const rowsHtml = rows.map((row) => `
    <tr>
      <td class="amount">${escapeHtml(formatMoneyValue(row.amount))}</td>
      <td class="currency">EUR</td>
      <td class="label">${escapeHtml(row.label)}</td>
    </tr>
  `).join('');

  container.innerHTML = `
    <table class="tug-breakdown-print-table">
      <thead>
        <tr><th colspan="3">${escapeHtml(cardTitle)}</th></tr>
      </thead>
      <tbody>
        ${rowsHtml}
        <tr class="total-row">
          <td class="amount">${escapeHtml(formatMoneyValue(calc.total))}</td>
          <td class="currency">EUR</td>
          <td class="label">Total</td>
        </tr>
      </tbody>
    </table>
  `;
}

function renderTugStandbyPrintBreakdown(charge, hoursRaw) {
  if (!standbyBreakdownEl) return;
  if (!Number.isFinite(charge) || charge <= 0) {
    standbyBreakdownEl.innerHTML = '';
    return;
  }

  const hours = parseMoneyValue(hoursRaw);
  const startedHours = Number.isFinite(hours) && hours > 0 ? Math.ceil(hours) : 0;
  const enteredHoursLabel = Number.isFinite(hours) && hours > 0 ? formatMoneyValue(hours) : '0,00';

  standbyBreakdownEl.innerHTML = `
    <table class="tug-breakdown-print-table">
      <thead>
        <tr><th colspan="3">Stand By Tugboat</th></tr>
      </thead>
      <tbody>
        <tr>
          <td class="amount">${escapeHtml(formatMoneyValue(TUG_STANDBY_RATE))}</td>
          <td class="currency">EUR</td>
          <td class="label">Tariff per started hour</td>
        </tr>
        <tr>
          <td class="amount">${escapeHtml(String(startedHours))}</td>
          <td class="currency">h</td>
          <td class="label">Started hours charged (entered ${escapeHtml(enteredHoursLabel)} h)</td>
        </tr>
        <tr class="total-row">
          <td class="amount">${escapeHtml(formatMoneyValue(charge))}</td>
          <td class="currency">EUR</td>
          <td class="label">Total</td>
        </tr>
      </tbody>
    </table>
  `;
}

function calculate() {
  saveTugsState();
  syncPrintSummaryValues();
  const tariffInput = document.getElementById('tariff');
  const final = document.getElementById('finalTotal');
  if (!tariffInput || !final) return;

  const tariff = Number(tariffInput.value) || 0;
  if (!tariff) {
    for (let i = 1; i <= tugCount; i += 1) {
      if (!document.getElementById(`tug_${i}`)) continue;
      const totalEl = document.getElementById(`tugTotal_${i}`);
      if (totalEl) totalEl.innerText = 'Tug total: €0.00';
      renderTugPrintBreakdown(i, null);
    }
    if (standbyTotalEl) standbyTotalEl.innerText = 'Stand by Tugboat total: €0.00';
    renderTugStandbyPrintBreakdown(0, '');
    final.style.display = 'none';
    safeStorageSet(STORAGE_KEYS.towageArrivalCount, 0);
    safeStorageSet(STORAGE_KEYS.towageDepartureCount, 0);
    safeStorageRemove(STORAGE_KEYS.towageTotal);
    return;
  }

  let arrivalTotal = 0;
  let departureTotal = 0;
  let arrivalCount = 0;
  let departureCount = 0;

  const discountPercent = getApprovedDiscountPercent();
  const discountFactor = discountPercent > 0 ? (1 - (discountPercent / 100)) : 1;
  const discountLabel = formatDiscountLabel(discountPercent);
  const discountSuffix = discountLabel ? ` (discounted ${discountLabel}%)` : '';
  const effectiveTariff = tariff * discountFactor;
  const gtValue = Number(document.getElementById('gt')?.value);
  const standbyCharge = standbyEnabled && standbyEnabled.checked
    ? getTugStandbyCharge(standbyHoursInput ? standbyHoursInput.value : '', gtValue)
    : 0;

  updateTugSummaryDiscountNote(tariff, discountPercent);
  if (standbyTotalEl) standbyTotalEl.innerText = `Stand by Tugboat total: €${standbyCharge.toFixed(2)}`;
  renderTugStandbyPrintBreakdown(standbyCharge, standbyHoursInput ? standbyHoursInput.value : '');

  for (let i = 1; i <= tugCount; i += 1) {
    if (!document.getElementById(`tug_${i}`)) continue;
    const calc = getTugCardCalculation(i, effectiveTariff);
    if (!calc) continue;

    if (calc.operation === 'arrival') arrivalCount += 1;
    else departureCount += 1;

    const totalEl = document.getElementById(`tugTotal_${i}`);
    if (totalEl) totalEl.innerText = `Tug total: €${calc.total.toFixed(2)}${discountSuffix}`;
    renderTugPrintBreakdown(i, calc);

    if (calc.operation === 'arrival') arrivalTotal += calc.total;
    else departureTotal += calc.total;
  }

  if (arrivalTotal === 0 && departureTotal === 0 && standbyCharge <= 0) {
    final.style.display = 'none';
    safeStorageSet(STORAGE_KEYS.towageArrivalCount, arrivalCount);
    safeStorageSet(STORAGE_KEYS.towageDepartureCount, departureCount);
    safeStorageRemove(STORAGE_KEYS.towageTotal);
    return;
  }

  const grandTotal = arrivalTotal + departureTotal + standbyCharge;
  const discountRow = discountLabel
    ? `<div class="summary-discount print-only">Approved discount: ${discountLabel}%</div>`
    : '';
  const standbyRow = standbyCharge > 0
    ? `<div><strong>Stand by Tugboat total</strong><br>€${standbyCharge.toFixed(2)}</div>`
    : '';

  final.style.display = 'block';
  final.innerHTML = `
    ${discountRow}
    <div class="summary">
      <div><strong>Arrival total${discountSuffix}</strong><br>€${arrivalTotal.toFixed(2)}</div>
      <div><strong>Departure total${discountSuffix}</strong><br>€${departureTotal.toFixed(2)}</div>
      ${standbyRow}
      <div><strong>Grand total${discountSuffix}</strong><br>€${grandTotal.toFixed(2)}</div>
    </div>
  `;

  safeStorageSet(STORAGE_KEYS.towageTotal, grandTotal.toFixed(2));
  safeStorageSet(STORAGE_KEYS.towageArrivalCount, arrivalCount);
  safeStorageSet(STORAGE_KEYS.towageDepartureCount, departureCount);
}

function resetTugs() {
  const shouldReset = window.confirm('Reset all tug data and refresh the page?');
  if (!shouldReset) return;
  isResettingTugs = true;
  Object.values(STORAGE_KEYS).forEach((key) => {
    safeStorageRemove(key);
  });
  window.location.reload();
}

function goHome() {
  saveTugsState();
  if (toggleMobileHomeCompactState()) return;
  window.location.href = 'index.html';
}

function syncTopHomeOptionsVisibility() {
  if (!tugOptionsToggleEl || !tugOptionsBodyEl) return;
  const topBar = document.querySelector('body.page-tugs .actions.top-home-actions');
  const isMobile = isLikelyMobileViewport();
  const isCompact = Boolean(isMobile && topBar && topBar.classList.contains('mobile-home-compact'));
  const showToggle = !isMobile;
  const showBody = isMobile ? !isCompact : tugOptionsExpanded;

  tugOptionsToggleEl.hidden = !showToggle;
  if (showToggle) tugOptionsToggleEl.removeAttribute('aria-hidden');
  else tugOptionsToggleEl.setAttribute('aria-hidden', 'true');

  tugOptionsToggleEl.setAttribute('aria-expanded', showBody ? 'true' : 'false');
  tugOptionsBodyEl.hidden = !showBody;
  syncTopHomeBarOffset();
}

function setTugOptionsExpanded(expanded) {
  tugOptionsExpanded = Boolean(expanded);
  syncTopHomeOptionsVisibility();
}

function applyMobileHomeCompactState(forceCompact = getStoredMobileHomeCompactState()) {
  const topBar = document.querySelector('body.page-tugs .actions.top-home-actions');
  const homeButton = topBar ? topBar.querySelector('.home-icon-btn') : null;
  const homeButtonLabel = homeButton ? homeButton.querySelector('.home-icon-toggle-label') : null;
  if (!topBar) return false;
  const nextCompact = isLikelyMobileViewport() && Boolean(forceCompact);
  topBar.classList.toggle('mobile-home-compact', nextCompact);
  if (homeButton) {
    homeButton.setAttribute('aria-pressed', nextCompact ? 'true' : 'false');
    homeButton.setAttribute('aria-label', nextCompact ? 'Show more fields' : 'Hide extra fields');
  }
  if (homeButtonLabel) {
    homeButtonLabel.textContent = nextCompact ? 'MORE' : 'LESS';
  }
  syncTopHomeOptionsVisibility();
  return nextCompact;
}

function toggleMobileHomeCompactState() {
  if (!isLikelyMobileViewport()) return false;
  const topBar = document.querySelector('body.page-tugs .actions.top-home-actions');
  const nextCompact = !(topBar && topBar.classList.contains('mobile-home-compact'));
  setStoredMobileHomeCompactState(nextCompact);
  applyMobileHomeCompactState(nextCompact);
  return true;
}

function syncTopHomeBarOffset() {
  const container = document.querySelector('body.page-tugs .container');
  const topBar = document.querySelector('body.page-tugs .actions.top-home-actions');
  if (!container || !topBar) return;
  if (window.matchMedia && window.matchMedia('print').matches) {
    container.style.paddingTop = '';
    return;
  }
  const topInset = Number.parseFloat(window.getComputedStyle(topBar).top) || 0;
  const nextPadding = Math.ceil(topBar.offsetHeight + topInset + 12);
  container.style.paddingTop = `${nextPadding}px`;
}

function syncTopHomeSummaryColumns() {
  const row = document.querySelector('body.page-tugs .top-home-summary-row');
  const vesselInput = document.getElementById('vesselName');
  if (!row || !vesselInput) return;
  if ((window.innerWidth || 0) <= 900) {
    row.style.removeProperty('--tug-vessel-col');
    return;
  }
  const value = vesselInput.value && vesselInput.value.trim() ? vesselInput.value.trim() : (vesselInput.placeholder || '');
  const probe = document.createElement('span');
  const computed = window.getComputedStyle(vesselInput);
  probe.textContent = value;
  probe.style.position = 'absolute';
  probe.style.visibility = 'hidden';
  probe.style.whiteSpace = 'pre';
  probe.style.font = computed.font;
  probe.style.letterSpacing = computed.letterSpacing;
  document.body.appendChild(probe);
  const horizontalPadding = (Number.parseFloat(computed.paddingLeft) || 0) + (Number.parseFloat(computed.paddingRight) || 0) + (Number.parseFloat(computed.borderLeftWidth) || 0) + (Number.parseFloat(computed.borderRightWidth) || 0);
  const desired = Math.ceil(probe.getBoundingClientRect().width + horizontalPadding + 24);
  probe.remove();
  const minimum = 220;
  const maximum = 420;
  const clamped = Math.max(minimum, Math.min(maximum, desired));
  row.style.setProperty('--tug-vessel-col', `${clamped}px`);
}

function initTugs() {
  const tugCards = document.getElementById('tugCards');
  if (!tugCards) return;

  const vesselNameInput = document.getElementById('vesselName');
  if (vesselNameInput) {
    updateVesselNameFromStorage(vesselNameInput);
    vesselNameInput.addEventListener('input', () => {
      const value = vesselNameInput.value.trim();
      if (value) safeStorageSet(STORAGE_KEYS.vesselName, value);
      else safeStorageRemove(STORAGE_KEYS.vesselName);
      syncTopHomeSummaryColumns();
      syncPrintSummaryValues();
      saveTugsState();
    });
  }

  const gtInput = document.getElementById('gt');
  const tariffInput = document.getElementById('tariff');

  imoMaster = document.getElementById('imoMaster');
  linesMaster = document.getElementById('linesMaster');
  arrivalOtMaster = document.getElementById('arrivalOtMaster');
  arrivalOtSunday = document.getElementById('arrivalOtSunday');
  departureOtMaster = document.getElementById('departureOtMaster');
  departureOtSunday = document.getElementById('departureOtSunday');
  discountEnabled = document.getElementById('discountEnabled');
  discountPercentInput = document.getElementById('discountPercent');
  discountFieldWrap = document.querySelector('.tug-discount-field');
  standbyEnabled = document.getElementById('standbyEnabled');
  standbyHoursInput = document.getElementById('standbyHours');
  standbyWrap = document.getElementById('tugStandbyWrap');
  standbyCard = document.getElementById('tugStandbyCard');
  standbyTotalEl = document.getElementById('standbyTotal');
  standbyBreakdownEl = document.getElementById('tugStandbyBreakdown');

  if (gtInput && tariffInput) {
    const syncGtAndTariff = () => {
      const storedGt = safeStorageGet(STORAGE_KEYS.gt);
      if (storedGt !== null && gtInput.value !== storedGt) {
        gtInput.value = storedGt;
      }
      const gt = Number(gtInput.value);
      tariffInput.value = getTariffFromGT(gt) || '';
      updateTugStandbyVisibility();
      syncPrintSummaryValues();
    };

    syncGtAndTariff();
    gtInput.addEventListener('input', () => {
      tariffInput.value = getTariffFromGT(Number(gtInput.value)) || '';
      const value = gtInput.value.trim();
      if (value) safeStorageSet(STORAGE_KEYS.gt, value);
      else safeStorageRemove(STORAGE_KEYS.gt);
      updateTugStandbyVisibility();
      syncPrintSummaryValues();
      calculate();
    });
  }

  tugOptionsToggleEl = document.getElementById('tugOptionsToggle');
  tugOptionsBodyEl = document.getElementById('tugOptionsBody');
  if (tugOptionsToggleEl && tugOptionsBodyEl) {
    tugOptionsToggleEl.addEventListener('click', () => {
      setTugOptionsExpanded(tugOptionsToggleEl.getAttribute('aria-expanded') !== 'true');
    });
    setTugOptionsExpanded(false);
  }

  document.addEventListener('input', calculate);
  document.addEventListener('change', (event) => {
    const target = event.target;
    if (target && target.id && target.id.startsWith('op_')) {
      const id = Number(target.id.replace('op_', ''));
      if (Number.isFinite(id)) applyOvertimeMastersToCard(id);
    }
    if (target && target.id && target.id.startsWith('imo_')) {
      syncImoMaster();
      if (imoMaster && !imoMaster.indeterminate) setGlobalImoTransportState(Boolean(imoMaster.checked));
    }
    if (target && target.id && target.id.startsWith('lines_')) {
      syncLinesMaster();
    }
    if (target && target.id && target.id.startsWith('kw_')) {
      const id = Number(target.id.replace('kw_', ''));
      if (Number.isFinite(id)) applyKwVoyageRule(id);
    }
    updateTugTitles();
    calculate();
  });

  tugCards.addEventListener('click', (event) => {
    const card = event.target.closest('.tug-service-card');
    if (!card) return;
    const id = card.id.split('_')[1];
    setSelectedTug(id);
  });

  if (imoMaster) {
    imoMaster.addEventListener('change', () => {
      applyImoMaster();
      setGlobalImoTransportState(Boolean(imoMaster.checked));
      updateImoToggleLabelColor(imoMaster);
      calculate();
    });
    updateImoToggleLabelColor(imoMaster);
  }

  if (linesMaster) {
    linesMaster.addEventListener('change', () => {
      applyLinesMaster();
      calculate();
    });
  }

  if (arrivalOtMaster) {
    arrivalOtMaster.addEventListener('change', () => {
      updateOvertimeMasterVisibility();
      applyOvertimeMaster('arrival', getMasterOvertimeRate(arrivalOtMaster, arrivalOtSunday));
      calculate();
    });
  }
  if (arrivalOtSunday) {
    arrivalOtSunday.addEventListener('change', () => {
      applyOvertimeMaster('arrival', getMasterOvertimeRate(arrivalOtMaster, arrivalOtSunday));
      calculate();
    });
  }
  if (departureOtMaster) {
    departureOtMaster.addEventListener('change', () => {
      updateOvertimeMasterVisibility();
      applyOvertimeMaster('departure', getMasterOvertimeRate(departureOtMaster, departureOtSunday));
      calculate();
    });
  }
  if (departureOtSunday) {
    departureOtSunday.addEventListener('change', () => {
      applyOvertimeMaster('departure', getMasterOvertimeRate(departureOtMaster, departureOtSunday));
      calculate();
    });
  }

  if (discountEnabled) {
    discountEnabled.addEventListener('change', () => {
      updateDiscountEnabledState();
      calculate();
    });
  }
  if (discountPercentInput) {
    discountPercentInput.addEventListener('input', calculate);
  }
  if (standbyEnabled) {
    standbyEnabled.addEventListener('change', () => {
      updateTugStandbyVisibility();
      calculate();
    });
  }
  if (standbyHoursInput) {
    standbyHoursInput.addEventListener('input', calculate);
  }

  restoreTugsState();
  updateTugStandbyVisibility();
  syncTugImoWithGlobalState(true);
  updateOvertimeMasterVisibility();
  updateDiscountEnabledState();
  calculate();
  syncTopHomeSummaryColumns();
  syncPrintSummaryValues();
  applyMobileHomeCompactState();
  syncTopHomeBarOffset();
  window.addEventListener('resize', () => {
    syncTopHomeSummaryColumns();
    applyMobileHomeCompactState();
    syncTopHomeBarOffset();
  });

  window.addEventListener('storage', (event) => {
    if (event.key === getScopedStorageKey(STORAGE_KEYS.globalImoTransport)) {
      syncTugImoWithGlobalState();
      calculate();
    }
    if (event.key === getScopedStorageKey(STORAGE_KEYS.vesselName) && vesselNameInput) {
      updateVesselNameFromStorage(vesselNameInput);
      syncPrintSummaryValues();
    }
    if (event.key === getScopedStorageKey(STORAGE_KEYS.gt) && gtInput && tariffInput) {
      const storedGt = safeStorageGet(STORAGE_KEYS.gt) || '';
      if (gtInput.value !== storedGt) gtInput.value = storedGt;
      tariffInput.value = getTariffFromGT(Number(gtInput.value)) || '';
      updateTugStandbyVisibility();
      syncPrintSummaryValues();
      calculate();
    }
  });

  window.addEventListener('beforeunload', saveTugsState);
  window.addEventListener('pagehide', saveTugsState);
}

if (typeof window !== 'undefined' && window.addEventListener) {
  window.addEventListener('afterprint', () => {
    disableMobilePrintFooterSuppression();
    restorePrintTitleIfNeeded();
    clearTransientIconAnimationState(document);
    syncTopHomeBarOffset();
  });

  window.addEventListener('beforeprint', () => {
    syncPrintSummaryValues();
    setPrintTitleFromVessel();
    if (isLikelyMobileViewport()) {
      enableMobilePrintFooterSuppression();
    } else {
      disableMobilePrintFooterSuppression();
    }
    syncTopHomeBarOffset();
  });

  window.addEventListener('pageshow', () => {
    clearTransientIconAnimationState(document);
  });

  window.addEventListener('DOMContentLoaded', () => {
    initIconClickAnimationCompletion();
    clearTransientIconAnimationState(document);
    initTugs();
  });
}
