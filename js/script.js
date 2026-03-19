/* Shared helpers */
let draggingRow = null;
let outlaysTable = null;
let toggleSailing = null;
let densityComfortable = null;
let densityDense = null;
let printRestoreDensity = null;
let printRestoreTitle = null;
let defaultOutlaysHtml = null;
let mobilePrintPageStyle = null;
let outlaysCurrency = null;
let roundPdaPrices = null;
const STANDALONE_STORAGE_PREFIX = 'standalone_tug_service:';

function isLikelyMobileViewport() {
  const ua = (navigator && navigator.userAgent) ? navigator.userAgent : '';
  if (/Android|webOS|iPhone|iPad|iPod|Mobile|CriOS|FxiOS/i.test(ua)) return true;
  const viewportWidth = window.visualViewport ? window.visualViewport.width : window.innerWidth;
  if (Number.isFinite(viewportWidth) && viewportWidth > 0 && viewportWidth <= 900) return true;
  const screenWidth = window.screen && Number.isFinite(window.screen.width) ? window.screen.width : 0;
  if (screenWidth > 0 && screenWidth <= 900) return true;
  return window.matchMedia('(max-width: 900px), (pointer: coarse)').matches;
}
const STORAGE_KEYS = {
  vesselName: 'pda_vessel_name',
  gt: 'pda_gt',
  quantity: 'pda_quantity',
  indexState: 'pda_index_state',
  towageTotal: 'pda_towage_total',
  towageArrivalCount: 'pda_towage_arrival_count',
  towageDepartureCount: 'pda_towage_departure_count',
  tugsState: 'pda_tugs_state',
  towageTotalSailing: 'pda_towage_total_sailing',
  towageArrivalCountSailing: 'pda_towage_arrival_count_sailing',
  towageDepartureCountSailing: 'pda_towage_departure_count_sailing',
  tugsStateSailing: 'pda_tugs_state_sailing',
  lightDuesState: 'pda_light_dues_state',
  lightDuesStateSailing: 'pda_light_dues_state_sailing',
  lightDuesAmountPda: 'pda_light_dues_amount_pda',
  lightDuesTariffPda: 'pda_light_dues_tariff_pda',
  lightDuesAmountSailing: 'pda_light_dues_amount_sailing',
  lightDuesTariffSailing: 'pda_light_dues_tariff_sailing',
  portDuesState: 'pda_port_dues_state',
  portDuesStateSailing: 'pda_port_dues_state_sailing',
  portDuesAmountPda: 'pda_port_dues_amount_pda',
  portDuesAmountSailing: 'pda_port_dues_amount_sailing',
  portDuesCargoAmountPda: 'pda_port_dues_cargo_amount_pda',
  portDuesCargoAmountSailing: 'pda_port_dues_cargo_amount_sailing',
  portDuesBunkeringAmountPda: 'pda_port_dues_bunkering_amount_pda',
  portDuesBunkeringAmountSailing: 'pda_port_dues_bunkering_amount_sailing',
  bunkeringState: 'pda_bunkering_state',
  bunkeringStateSailing: 'pda_bunkering_state_sailing',
  berthageState: 'pda_berthage_state',
  berthageStateSailing: 'pda_berthage_state_sailing',
  berthageAmountPda: 'pda_berthage_amount_pda',
  berthageAmountSailing: 'pda_berthage_amount_sailing',
  mooringState: 'pda_mooring_state',
  mooringStateSailing: 'pda_mooring_state_sailing',
  mooringAmountPda: 'pda_mooring_amount_pda',
  mooringAmountSailing: 'pda_mooring_amount_sailing',
  globalImoTransport: 'pda_global_imo_transport',
  pilotBoatState: 'pda_pilot_boat_state',
  pilotBoatStateSailing: 'pda_pilot_boat_state_sailing',
  pilotageState: 'pda_pilotage_state',
  pilotageStateSailing: 'pda_pilotage_state_sailing',
  pilotageAmountPda: 'pda_pilotage_amount_pda',
  pilotageAmountSailing: 'pda_pilotage_amount_sailing',
  pilotBoatAmountPda: 'pda_pilot_boat_amount_pda',
  pilotBoatAmountSailing: 'pda_pilot_boat_amount_sailing'
};

const INDEX_FIELD_IDS = [
  'logoLeftNote',
  'titleNote',
  'dateInput',
  'vesselNameIndex',
  'grossTonnage',
  'lengthOverall',
  'bowThrusterFitted',
  'portInput',
  'berthTerminal',
  'operationsInput',
  'cargoInput',
  'quantityInput',
  'agentInput',
  'togglePda',
  'toggleSailing',
  'toggleLogoNote',
  'globalImoTransport',
  'roundPdaPrices',
  'outlaysCurrency',
  'bankRate'
];

const TUG_STORAGE = {
  standard: {
    towageTotal: STORAGE_KEYS.towageTotal,
    towageArrivalCount: STORAGE_KEYS.towageArrivalCount,
    towageDepartureCount: STORAGE_KEYS.towageDepartureCount,
    tugsState: STORAGE_KEYS.tugsState
  },
  sailing: {
    towageTotal: STORAGE_KEYS.towageTotalSailing,
    towageArrivalCount: STORAGE_KEYS.towageArrivalCountSailing,
    towageDepartureCount: STORAGE_KEYS.towageDepartureCountSailing,
    tugsState: STORAGE_KEYS.tugsStateSailing
  }
};

const LIGHT_DUES_TARIFFS = {
  cargo: {
    rate30: 0.5088,
    rate12: 1.696,
    label30: '0,5088',
    label12: '1,696'
  },
  tanker: {
    rate30: 0.579768,
    rate12: 1.940448,
    label30: '0,579768',
    label12: '1,940448'
  },
  roroCargo: {
    rate30: 0.2,
    rate12: 0.664,
    label30: '0,2',
    label12: '0,664'
  },
  passengerRoroFerry: {
    rate30: 0.2014,
    rate12: 0.6784,
    label30: '0,2014',
    label12: '0,6784'
  },
  supply: {
    rate30: 0.5088,
    rate12: 1.696,
    label30: '0,5088',
    label12: '1,696'
  },
  tugboatPusher: {
    rate30: 0.1325,
    rate12: 0.4399,
    label30: '0,1325',
    label12: '0,4399'
  },
  fishing: {
    rate30: 0.2014,
    rate12: 0.6784,
    label30: '0,2014',
    label12: '0,6784'
  },
  technicalCraft: {
    rate30: 0.2014,
    rate12: 0.6784,
    label30: '0,2014',
    label12: '0,6784'
  },
  nonSelfPropelled: {
    rate30: 0.2014,
    rate12: 0.6784,
    label30: '0,2014',
    label12: '0,6784'
  },
  otherUnidentified: {
    rate30: 0.5088,
    rate12: 1.696,
    label30: '0,5088',
    label12: '1,696'
  },
  bulk: {
    le30000: {
      rate30: 0.45792,
      rate12: 1.5264,
      label30: '0,45792',
      label12: '1,5264'
    },
    '30001to50000': {
      rate30: 0.4028,
      rate12: 1.378,
      label30: '0,4028',
      label12: '1,378'
    },
    gt50000: {
      rate30: 0.2968,
      rate12: 0.8904,
      label30: '0,2968',
      label12: '0,8904'
    }
  },
  container: {
    le40000: {
      rate30: 0.2332,
      rate12: 1.06,
      label30: '0,2332',
      label12: '1,06'
    },
    '40001to65000': {
      rate30: 0.10388,
      rate12: 0.34238,
      label30: '0,10388',
      label12: '0,34238'
    },
    '65001to100000': {
      rate30: 0.0742,
      rate12: 0.2862,
      label30: '0,0742',
      label12: '0,2862'
    },
    gt100000: {
      rate30: 0.053,
      rate12: 0.1908,
      label30: '0,053',
      label12: '0,1908'
    }
  },
  cruise: {
    le20000: {
      rate30: 0.25,
      rate12: 0.875,
      label30: '0,25',
      label12: '0,875'
    },
    '20001to50000': {
      rate30: 0.20125,
      rate12: 0.6875,
      label30: '0,20125',
      label12: '0,6875'
    },
    '50001to80000': {
      rate30: 0.1725,
      rate12: 0.575,
      label30: '0,1725',
      label12: '0,575'
    },
    gt80000: {
      rate30: 0.17125,
      rate12: 0.5,
      label30: '0,17125',
      label12: '0,5'
    }
  }
};

const LIGHT_DUES_TIER_OPTIONS = {
  bulk: [
    { value: 'le30000', label: 'if ≤ 30.000 GT', max: 30000 },
    { value: '30001to50000', label: 'if 30.001 - 50.000 GT', min: 30001, max: 50000 },
    { value: 'gt50000', label: 'if > 50.000 GT', min: 50001 }
  ],
  container: [
    { value: 'le40000', label: 'if ≤ 40.000 GT', max: 40000 },
    { value: '40001to65000', label: 'if 40.001 - 65.000 GT', min: 40001, max: 65000 },
    { value: '65001to100000', label: 'if 65.001 - 100.000 GT', min: 65001, max: 100000 },
    { value: 'gt100000', label: 'if > 100.000 GT', min: 100001 }
  ],
  cruise: [
    { value: 'le20000', label: 'if ≤ 20.000 GT', max: 20000 },
    { value: '20001to50000', label: 'if 20.001 - 50.000 GT', min: 20001, max: 50000 },
    { value: '50001to80000', label: 'if 50.001 - 80.000 GT', min: 50001, max: 80000 },
    { value: 'gt80000', label: 'if > 80.000 GT', min: 80001 }
  ]
};
const LIGHT_DUES_DEFAULT_TYPE = 'cargo';

const PORT_DUES_CARGO_TYPES = {
  bulkCargo: { rate: 0.6, label: '0,60' },
  cementGravelSawdustBulkStoneSlagPetroleumCokeClinkerCoal: { rate: 0.33, label: '0,33' },
  grains: { rate: 0.5, label: '0,50' },
  sulphur: { rate: 0.48, label: '0,48' },
  scrapIron: { rate: 0.33, label: '0,33' },
  copper: { rate: 0.4, label: '0,40' },
  liquidCargo: { rate: 1, label: '1,00' },
  generalCargo: { rate: 1, label: '1,00' },
  baggedCement: { rate: 0.8, label: '0,80' },
  explosiveCargo: { rate: 2.9, label: '2,90' },
  heavyCargo: { rate: 2, label: '2,00' },
  containerPerGt: { rate: 0.84, label: '0,84' }
};
const PORT_DUES_DEFAULT_CARGO_TYPE = 'bulkCargo';

const MOORING_GT_TARIFFS = [
  { min: 0, max: 250, amount: 8.4 },
  { min: 251, max: 500, amount: 14.7 },
  { min: 501, max: 1000, amount: 21.7 },
  { min: 1001, max: 2000, amount: 30.8 },
  { min: 2001, max: 4000, amount: 43.4 },
  { min: 4001, max: 7000, amount: 61.04 },
  { min: 7001, max: 11000, amount: 86.1 },
  { min: 11001, max: 15000, amount: 121.8 },
  { min: 15001, max: 20000, amount: 162.4 },
  { min: 20001, max: 25000, amount: 203 },
  { min: 25001, max: 30000, amount: 243.6 },
  { min: 30001, max: 35000, amount: 285.6 }
];

const PILOTAGE_GT_TARIFFS = [
  { min: 0, max: 5000, amount: 160 },
  { min: 5001, max: 10000, amount: 240 },
  { min: 10001, max: 20000, amount: 520 },
  { min: 20001, max: 30000, amount: 680 },
  { min: 30001, max: 50000, amount: 1100 }
];
const PILOT_BOAT_BASE_CHARGE = 360;
const PILOT_BOAT_EXTRA_NM_RATE = 85;
const PILOT_BOAT_WAITING_15MIN_RATE = 120;
const BUNKER_ROW_DESCRIPTION = 'BUNKER (EUR 0,30 x loaded bunker / MT )';
const BERTHAGE_QUAYAGE_ROW_DESCRIPTION = 'BERTHAGE/QUAYAGE/ANCHORAGE';
const BERTHAGE_TARIFF_RATE = 2;
const ANCHORAGE_TARIFF_RATE = 1;
const ICON_PRESS_ANIMATION_CLASS = 'icon-animating';
const ICON_PRESS_ANIMATION_MS = 176;
const iconPressAnimationTimers = new WeakMap();
const NEW_DRAFT_CALC_KEYS = [
  STORAGE_KEYS.towageTotal,
  STORAGE_KEYS.towageArrivalCount,
  STORAGE_KEYS.towageDepartureCount,
  STORAGE_KEYS.tugsState,
  STORAGE_KEYS.towageTotalSailing,
  STORAGE_KEYS.towageArrivalCountSailing,
  STORAGE_KEYS.towageDepartureCountSailing,
  STORAGE_KEYS.tugsStateSailing,
  STORAGE_KEYS.lightDuesState,
  STORAGE_KEYS.lightDuesStateSailing,
  STORAGE_KEYS.lightDuesAmountPda,
  STORAGE_KEYS.lightDuesTariffPda,
  STORAGE_KEYS.lightDuesAmountSailing,
  STORAGE_KEYS.lightDuesTariffSailing,
  STORAGE_KEYS.portDuesState,
  STORAGE_KEYS.portDuesStateSailing,
  STORAGE_KEYS.portDuesAmountPda,
  STORAGE_KEYS.portDuesAmountSailing,
  STORAGE_KEYS.portDuesCargoAmountPda,
  STORAGE_KEYS.portDuesCargoAmountSailing,
  STORAGE_KEYS.portDuesBunkeringAmountPda,
  STORAGE_KEYS.portDuesBunkeringAmountSailing,
  STORAGE_KEYS.bunkeringState,
  STORAGE_KEYS.bunkeringStateSailing,
  STORAGE_KEYS.berthageState,
  STORAGE_KEYS.berthageStateSailing,
  STORAGE_KEYS.berthageAmountPda,
  STORAGE_KEYS.berthageAmountSailing,
  STORAGE_KEYS.mooringState,
  STORAGE_KEYS.mooringStateSailing,
  STORAGE_KEYS.mooringAmountPda,
  STORAGE_KEYS.mooringAmountSailing,
  STORAGE_KEYS.pilotageState,
  STORAGE_KEYS.pilotageStateSailing,
  STORAGE_KEYS.pilotageAmountPda,
  STORAGE_KEYS.pilotageAmountSailing,
  STORAGE_KEYS.pilotBoatState,
  STORAGE_KEYS.pilotBoatStateSailing,
  STORAGE_KEYS.pilotBoatAmountPda,
  STORAGE_KEYS.pilotBoatAmountSailing,
  STORAGE_KEYS.globalImoTransport
];

function isNewDraftMode() {
  const params = new URLSearchParams(window.location.search);
  return params.get('new') === '1';
}

function clearCalculatorStorageForNewDraft() {
  NEW_DRAFT_CALC_KEYS.forEach((key) => safeStorageRemove(key));
}

function getScopedStorageKey(key) {
  return `${STANDALONE_STORAGE_PREFIX}${key}`;
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

function resolveLightDuesTypeForImo(type, globalImo, options = {}) {
  const currentType = typeof type === 'string' && type ? type : LIGHT_DUES_DEFAULT_TYPE;
  if (globalImo) return 'tanker';
  if (options.resetForced && currentType === 'tanker') return LIGHT_DUES_DEFAULT_TYPE;
  return currentType;
}

function resolvePortDuesCargoTypeForImo(cargoType, globalImo, options = {}) {
  const currentType = typeof cargoType === 'string' && cargoType ? cargoType : PORT_DUES_DEFAULT_CARGO_TYPE;
  const lastManualCargoType = (
    typeof options.lastManualCargoType === 'string'
    && PORT_DUES_CARGO_TYPES[options.lastManualCargoType]
  )
    ? options.lastManualCargoType
    : PORT_DUES_DEFAULT_CARGO_TYPE;
  if (globalImo) return 'liquidCargo';
  if (options.resetForced && (options.wasForcedByImo || currentType === 'liquidCargo')) {
    return lastManualCargoType;
  }
  return currentType;
}

function shouldApplyGlobalImoStateOnInit(storedState, checkboxChecked) {
  if (storedState === null) return Boolean(checkboxChecked);
  return Boolean(storedState);
}

function setLightDuesTypeState(type) {
  if (!type) return;
  const keys = [STORAGE_KEYS.lightDuesState, STORAGE_KEYS.lightDuesStateSailing];
  keys.forEach((key) => {
    const raw = safeStorageGet(key);
    let state = {};
    if (raw) {
      try {
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === 'object') state = parsed;
      } catch (error) {
        state = {};
      }
    }
    state.type = type;
    safeStorageSet(key, JSON.stringify(state));
  });
}

function syncLightDuesTypeStateFromImo(checked) {
  const keys = [STORAGE_KEYS.lightDuesState, STORAGE_KEYS.lightDuesStateSailing];
  keys.forEach((key) => {
    const raw = safeStorageGet(key);
    let state = {};
    if (raw) {
      try {
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === 'object') state = parsed;
      } catch (error) {
        state = {};
      }
    }
    state.type = resolveLightDuesTypeForImo(state.type, checked, { resetForced: true });
    safeStorageSet(key, JSON.stringify(state));
  });
}

function setPortDuesCargoTypeState(cargoType) {
  if (!cargoType) return;
  const keys = [STORAGE_KEYS.portDuesState, STORAGE_KEYS.portDuesStateSailing];
  keys.forEach((key) => {
    const raw = safeStorageGet(key);
    let state = {};
    if (raw) {
      try {
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === 'object') state = parsed;
      } catch (error) {
        state = {};
      }
    }
    state.cargoType = cargoType;
    state.cargoTypeForcedByImo = false;
    if (PORT_DUES_CARGO_TYPES[cargoType]) {
      state.lastManualCargoType = cargoType;
    }
    safeStorageSet(key, JSON.stringify(state));
  });
}

function syncPortDuesCargoTypeStateFromImo(checked) {
  const keys = [STORAGE_KEYS.portDuesState, STORAGE_KEYS.portDuesStateSailing];
  keys.forEach((key) => {
    const raw = safeStorageGet(key);
    let state = {};
    if (raw) {
      try {
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === 'object') state = parsed;
      } catch (error) {
        state = {};
      }
    }
    const currentCargoType = (
      typeof state.cargoType === 'string' && PORT_DUES_CARGO_TYPES[state.cargoType]
    )
      ? state.cargoType
      : PORT_DUES_DEFAULT_CARGO_TYPE;
    const lastManualCargoType = (
      typeof state.lastManualCargoType === 'string' && PORT_DUES_CARGO_TYPES[state.lastManualCargoType]
    )
      ? state.lastManualCargoType
      : PORT_DUES_DEFAULT_CARGO_TYPE;

    if (checked) {
      if (currentCargoType !== 'liquidCargo') {
        state.lastManualCargoType = currentCargoType;
      } else if (!PORT_DUES_CARGO_TYPES[state.lastManualCargoType]) {
        state.lastManualCargoType = lastManualCargoType;
      }
      state.cargoType = 'liquidCargo';
      state.cargoTypeForcedByImo = true;
    } else {
      state.cargoType = resolvePortDuesCargoTypeForImo(currentCargoType, false, {
        resetForced: true,
        wasForcedByImo: Boolean(state.cargoTypeForcedByImo),
        lastManualCargoType
      });
      state.cargoTypeForcedByImo = false;
      if (PORT_DUES_CARGO_TYPES[state.cargoType]) {
        state.lastManualCargoType = state.cargoType;
      }
    }
    safeStorageSet(key, JSON.stringify(state));
  });
}

function setGlobalImoTransportState(checked) {
  safeStorageSet(STORAGE_KEYS.globalImoTransport, checked ? '1' : '0');
  syncPortDuesCargoTypeStateFromImo(Boolean(checked));
  syncLightDuesTypeStateFromImo(Boolean(checked));
}

function getVesselNameForPrint() {
  const stored = safeStorageGet(STORAGE_KEYS.vesselName);
  if (stored) return String(stored).trim();
  const indexField = document.getElementById('vesselNameIndex');
  if (indexField && indexField.value) return String(indexField.value).trim();
  const tugField = document.getElementById('vesselName');
  if (tugField && tugField.value) return String(tugField.value).trim();
  return '';
}

function setPrintTitleFromVessel() {
  if (printRestoreTitle === null) {
    printRestoreTitle = document.title;
  }
  const vesselName = getVesselNameForPrint();
  const isTugsPage = document.body && document.body.classList.contains('page-tugs');
  const baseTitle = isTugsPage ? 'Tug Service Calculator' : 'PRO-FORMA D/A';
  document.title = vesselName ? `${baseTitle} ${vesselName}` : baseTitle;
}

function restorePrintTitleIfNeeded() {
  if (printRestoreTitle === null) return;
  document.title = printRestoreTitle;
  printRestoreTitle = null;
}

function printWithTitle() {
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
  return button.matches('.row-edit, .row-remove, .pda-db-edit-btn, .pda-db-delete-btn, .icon-btn');
}

function triggerIconPressAnimation(button) {
  if (!button || !isAnimatedIconButton(button) || button.disabled) return;
  const existingTimer = iconPressAnimationTimers.get(button);
  if (existingTimer) window.clearTimeout(existingTimer);

  button.classList.remove(ICON_PRESS_ANIMATION_CLASS);
  // Force reflow so repeated presses retrigger animation reliably.
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

function getPersistableOutlaysHtml(outlaysBody) {
  if (!outlaysBody) return '';
  const clone = outlaysBody.cloneNode(true);
  clearTransientIconAnimationState(clone);
  clone.querySelectorAll('tr.dragging').forEach((row) => row.classList.remove('dragging'));
  return clone.innerHTML;
}

function getDefaultOutlaysHtml() {
  if (defaultOutlaysHtml) return defaultOutlaysHtml;
  const outlaysBody = document.getElementById('outlaysBody');
  if (!outlaysBody) return '';
  defaultOutlaysHtml = outlaysBody.innerHTML;
  return defaultOutlaysHtml;
}

function ensureBaselineOutlayRows(outlaysBody) {
  if (!outlaysBody) return;
  const baselineHtml = getDefaultOutlaysHtml();
  if (!baselineHtml) return;

  const existingKeys = new Set();
  outlaysBody.querySelectorAll('tr').forEach((row) => {
    const key = String(row.dataset.row || '').trim();
    if (key) existingKeys.add(key);
  });

  const baseline = document.createElement('tbody');
  baseline.innerHTML = baselineHtml;
  baseline.querySelectorAll('tr').forEach((row) => {
    const key = String(row.dataset.row || '').trim();
    if (!key || key === 'bank-charges') return;
    if (existingKeys.has(key)) return;
    const clone = row.cloneNode(true);
    clone.hidden = true;
    clone.dataset.outlayHidden = '1';
    const bankRow = outlaysBody.querySelector('tr[data-row=\"bank-charges\"]');
    if (bankRow) outlaysBody.insertBefore(clone, bankRow);
    else outlaysBody.appendChild(clone);
  });
}

function initIconClickAnimationCompletion() {
  document.addEventListener('pointerdown', (event) => {
    if (!event.isTrusted) return;
    if (event.button !== 0) return;
    const target = event.target;
    const button = target && target.closest ? target.closest('button') : null;
    if (!button) return;
    triggerIconPressAnimation(button);
  }, true);

  document.addEventListener('keydown', (event) => {
    if (!event.isTrusted) return;
    if (event.key !== 'Enter' && event.key !== ' ') return;
    const target = event.target;
    const button = target && target.closest ? target.closest('button') : null;
    if (!button) return;
    triggerIconPressAnimation(button);
  }, true);
}

function updateImoToggleLabelColor(input) {
  if (!input) return;
  const toggleLabel = input.closest('label');
  if (!toggleLabel) return;
  const supportsCheckedColor =
    toggleLabel.classList.contains('global-imo-toggle') ||
    toggleLabel.classList.contains('imo-master-label');
  if (!supportsCheckedColor) return;
  toggleLabel.classList.toggle('is-checked', Boolean(input.checked));
}

function readFieldValue(field) {
  if (!field) return '';
  if (field.type === 'checkbox') return Boolean(field.checked);
  const raw = String(field.value || '').trim();
  const placeholder = String(field.placeholder || '').trim();
  if (raw && placeholder && raw === placeholder && field.dataset.userEdited !== '1') {
    return '';
  }
  return raw;
}

function applyFieldValue(field, value) {
  if (!field || value === undefined || value === null) return;
  if (field.type === 'checkbox') {
    field.checked = Boolean(value);
    return;
  }
  const nextValue = value == null ? '' : String(value);
  field.value = nextValue;
  if (field.dataset) {
    if (nextValue.trim()) field.dataset.userEdited = '1';
    else delete field.dataset.userEdited;
  }
}

function migrateLegacyIconMarkup(markup) {
  if (typeof markup !== 'string' || !markup) return markup;
  return markup
    .replace(/assets\/icons\/copy-plus\.(?:png|svg)/g, 'assets/icons/copy-plus.svg')
    .replace(/assets\/icons\/edit_48.\.png/g, 'assets/icons/square-pen.svg')
    .replace(/assets\/icons\/remove_48.\.png/g, 'assets/icons/trash.svg')
    .replace(/assets\/icons\/move_grabber_48.\.png/g, 'assets/icons/grip-vertical.svg');
}

function saveIndexState() {
  const outlaysBody = document.getElementById('outlaysBody');
  if (!outlaysBody) return;

  const fields = {};
  INDEX_FIELD_IDS.forEach((id) => {
    const field = document.getElementById(id);
    if (!field) return;
    fields[id] = readFieldValue(field);
  });

  const outlaysValues = Array.from(outlaysBody.querySelectorAll('tr')).map((row) => {
    return Array.from(row.querySelectorAll('textarea, input')).map((field) => field.value);
  });

  const state = {
    fields,
    outlaysHtml: getPersistableOutlaysHtml(outlaysBody),
    outlaysValues,
    density: getDensityMode()
  };
  safeStorageSet(STORAGE_KEYS.indexState, JSON.stringify(state));
}

function restoreIndexState() {
  const outlaysBody = document.getElementById('outlaysBody');
  if (!outlaysBody) return false;

  const raw = safeStorageGet(STORAGE_KEYS.indexState);
  if (!raw) return false;

  let state = null;
  try {
    state = JSON.parse(raw);
  } catch (error) {
    return false;
  }
  if (!state || typeof state !== 'object') return false;

  if (typeof state.outlaysHtml === 'string' && state.outlaysHtml.trim()) {
    outlaysBody.innerHTML = migrateLegacyIconMarkup(state.outlaysHtml);
    clearTransientIconAnimationState(outlaysBody);
  }

  if (Array.isArray(state.outlaysValues)) {
    const rows = outlaysBody.querySelectorAll('tr');
    state.outlaysValues.forEach((values, rowIndex) => {
      const row = rows[rowIndex];
      if (!row || !Array.isArray(values)) return;
      const fields = row.querySelectorAll('textarea, input');
      values.forEach((value, fieldIndex) => {
        const field = fields[fieldIndex];
        if (!field) return;
        field.value = value == null ? '' : String(value);
      });
    });
  }

  ensureBaselineOutlayRows(outlaysBody);

  if (state.fields && typeof state.fields === 'object') {
    INDEX_FIELD_IDS.forEach((id) => {
      const field = document.getElementById(id);
      if (!field) return;
      applyFieldValue(field, state.fields[id]);
    });
  }

  if (state.density === 'comfortable' || state.density === 'dense') {
    setDensity(state.density);
  }
  return true;
}

let isRestoringTugs = false;

function isSailingTugsPage() {
  return document.body && document.body.classList.contains('page-tugs-sailing');
}

function getTugStorageKeys(isSailing) {
  return isSailing ? TUG_STORAGE.sailing : TUG_STORAGE.standard;
}

function getTugsState(isSailing) {
  const keys = getTugStorageKeys(isSailing);
  const raw = safeStorageGet(keys.tugsState);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (error) {
    return null;
  }
}

function saveTugsState() {
  if (isRestoringTugs) return;
  const tugCards = document.getElementById('tugCards');
  if (!tugCards) return;

  const keys = getTugStorageKeys(isSailingTugsPage());
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

  safeStorageSet(keys.tugsState, JSON.stringify(state));
}

function restoreTugsState() {
  const tugCards = document.getElementById('tugCards');
  if (!tugCards) return false;

  const state = getTugsState(isSailingTugsPage());
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
  if (vesselName) {
    vesselName.value = state.vesselName || '';
    updateVesselNameFromStorage(vesselName);
  }

  const imoMasterInput = document.getElementById('imoMaster');
  if (imoMasterInput) {
    imoMasterInput.checked = Boolean(state.imoMaster);
  }
  const linesMasterInput = document.getElementById('linesMaster');
  if (linesMasterInput) {
    linesMasterInput.checked = Boolean(state.linesMaster);
  }
  const arrivalOtMasterInput = document.getElementById('arrivalOtMaster');
  if (arrivalOtMasterInput) {
    arrivalOtMasterInput.checked = Boolean(state.arrivalOtMaster);
  }
  const arrivalOtSundayInput = document.getElementById('arrivalOtSunday');
  if (arrivalOtSundayInput) {
    arrivalOtSundayInput.checked = Boolean(state.arrivalOtSunday);
  }
  const departureOtMasterInput = document.getElementById('departureOtMaster');
  if (departureOtMasterInput) {
    departureOtMasterInput.checked = Boolean(state.departureOtMaster);
  }
  const departureOtSundayInput = document.getElementById('departureOtSunday');
  if (departureOtSundayInput) {
    departureOtSundayInput.checked = Boolean(state.departureOtSunday);
  }
  const discountEnabledInput = document.getElementById('discountEnabled');
  if (discountEnabledInput) {
    discountEnabledInput.checked = Boolean(state.discountEnabled);
  }
  const discountPercentField = document.getElementById('discountPercent');
  if (discountPercentField) {
    discountPercentField.value = state.discountPercent == null ? '' : String(state.discountPercent);
  }
  const standbyEnabledInput = document.getElementById('standbyEnabled');
  if (standbyEnabledInput) {
    standbyEnabledInput.checked = Boolean(state.standbyEnabled);
  }
  const standbyHoursField = document.getElementById('standbyHours');
  if (standbyHoursField) {
    standbyHoursField.value = state.standbyHours == null ? '' : String(state.standbyHours);
  }

  const gtInput = document.getElementById('gt');
  if (gtInput) {
    const sharedGt = safeStorageGet(STORAGE_KEYS.gt);
    if (sharedGt !== null) gtInput.value = sharedGt;
    else if (state.gt) gtInput.value = state.gt;
    const gtValue = Number(gtInput.value);
    const tariff = getTariffFromGT(gtValue);
    const tariffInput = document.getElementById('tariff');
    if (tariffInput) tariffInput.value = tariff || '';
  }

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

function copyTugboatsFromPdaToSailing() {
  if (!isSailingTugsPage()) return false;

  const sourceState = getTugsState(false);
  if (!sourceState || !Array.isArray(sourceState.tugs)) return false;

  const vesselNameInput = document.getElementById('vesselName');
  const gtInput = document.getElementById('gt');
  const keys = getTugStorageKeys(true);
  const nextState = {
    vesselName: vesselNameInput ? String(vesselNameInput.value || '') : String(sourceState.vesselName || ''),
    gt: gtInput ? String(gtInput.value || '') : String(sourceState.gt || ''),
    imoMaster: Boolean(sourceState.imoMaster),
    linesMaster: Boolean(sourceState.linesMaster),
    arrivalOtMaster: Boolean(sourceState.arrivalOtMaster),
    arrivalOtSunday: Boolean(sourceState.arrivalOtSunday),
    departureOtMaster: Boolean(sourceState.departureOtMaster),
    departureOtSunday: Boolean(sourceState.departureOtSunday),
    discountEnabled: Boolean(sourceState.discountEnabled),
    discountPercent: sourceState.discountPercent == null ? '' : String(sourceState.discountPercent),
    tugs: sourceState.tugs.map((tug) => ({
      op: tug && tug.op === 'departure' ? 'departure' : 'arrival',
      voyage: tug && tug.voyage != null ? String(tug.voyage) : '1',
      assist: tug && tug.assist != null ? String(tug.assist) : '0.5',
      imo: Boolean(tug && tug.imo),
      lines: Boolean(tug && tug.lines),
      kw: Boolean(tug && tug.kw),
      voy_ot: tug && tug.voy_ot != null ? String(tug.voy_ot) : '0',
      assist_ot: tug && tug.assist_ot != null ? String(tug.assist_ot) : '0'
    }))
  };

  safeStorageSet(keys.tugsState, JSON.stringify(nextState));
  return restoreTugsState();
}

async function addOutlayRow(values = {}) {
  const tbody = document.getElementById('outlaysBody');
  const template = document.getElementById('outlayRowTemplate');
  if (!tbody || !template) return;

  const hasPresetValues = Boolean(values && (values.desc || values.pda || values.sailing));
  if (!hasPresetValues) {
    const hiddenProtectedRows = getProtectedOutlayRows({ hidden: true });
    if (hiddenProtectedRows.length > 0) {
      const selection = await promptProtectedOutlayRestore(hiddenProtectedRows);
      if (selection.type === 'restore') {
        restoreProtectedOutlayRow(selection.item);
        return;
      }
      if (selection.type === 'cancel') {
        return;
      }
    }
  }

  const row = template.content.firstElementChild.cloneNode(true);
  const desc = row.querySelector('[data-field="desc"]');
  const pda = row.querySelector('[data-field="pda"]');
  const sailing = row.querySelector('[data-field="sailing"]');

  if (desc) desc.value = values.desc || '';
  if (pda) pda.value = values.pda || '';
  if (sailing) sailing.value = values.sailing || '';

  const bankRow = tbody.querySelector('tr[data-row="bank-charges"]');
  if (bankRow) {
    tbody.insertBefore(row, bankRow);
  } else {
    tbody.appendChild(row);
  }
  if (desc && desc.tagName === 'TEXTAREA') autoResizeTextarea(desc);
  wrapMoneyFields();
  decorateMoneyEditCells();
  recalcOutlayTotals();
  saveIndexState();
}

function clearOutlayRow(row) {
  row.querySelectorAll('input, textarea').forEach((field) => {
    field.value = '';
    if (field.tagName === 'TEXTAREA') autoResizeTextarea(field);
  });
  recalcOutlayTotals();
  saveIndexState();
}

const PROTECTED_OUTLAY_KEY_ATTR = 'protectedOutlayKey';
let protectedOutlayKeyCounter = 0;
const PROTECTED_OUTLAY_REFRESH_DEFINITIONS = [
  {
    key: 'light-dues',
    match: (row, desc) => row.dataset.row === 'light-dues' || desc.startsWith('LIGHT DUES')
  },
  {
    key: 'port-dues',
    match: (row, desc) => row.dataset.row === 'port-dues' || desc.startsWith('PORT DUES')
  },
  {
    key: 'bunkering',
    match: (row, desc) => row.dataset.row === 'bunkering' || row.dataset.row === 'bunker' || desc.startsWith('BUNKER')
  },
  {
    key: 'berthage-quayage',
    match: (row, desc) =>
      row.dataset.row === 'berthage-quayage' ||
      row.dataset.row === 'berthage-anchorage-extra' ||
      desc.startsWith('BERHAGE/QUAYAGE') ||
      desc.startsWith('BERTHAGE/QUAYAGE')
  },
  {
    key: 'pilotage',
    match: (row, desc) => row.dataset.row === 'pilotage' || desc.startsWith('PILOTAGE')
  },
  {
    key: 'pilot-boat',
    match: (row, desc) => row.dataset.row === 'pilot-boat' || desc.startsWith('PILOT BOAT')
  },
  {
    key: 'mooring',
    match: (row, desc) => row.dataset.row === 'mooring' || desc.startsWith('MOORING/UNMOORING')
  },
  {
    key: 'towage',
    match: (row, desc) => row.dataset.row === 'towage' || desc.startsWith('TOWAGE')
  }
];

function createProtectedOutlayKey() {
  protectedOutlayKeyCounter += 1;
  return `outlay-${Date.now()}-${protectedOutlayKeyCounter}`;
}

function ensureProtectedOutlayKey(row) {
  if (!row) return '';
  const existing = String(row.dataset[PROTECTED_OUTLAY_KEY_ATTR] || '').trim();
  if (existing) return existing;
  const next = createProtectedOutlayKey();
  row.dataset[PROTECTED_OUTLAY_KEY_ATTR] = next;
  return next;
}

function getOutlayRowLabel(row) {
  if (!row) return 'Untitled row';
  const descField = row.querySelector('td.desc textarea, td.desc input');
  const text = String(descField?.value || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (text) return text;
  return 'Untitled row';
}

function getOutlayRefreshDefinition(row) {
  if (!row) return null;
  const desc = getOutlayDescription(row);
  return PROTECTED_OUTLAY_REFRESH_DEFINITIONS.find((item) => item.match(row, desc)) || null;
}

function getProtectedOutlayInfoForRow(row) {
  if (!row) return null;
  if (row.dataset.row === 'bank-charges') return null;
  const refreshDefinition = getOutlayRefreshDefinition(row);
  return {
    key: ensureProtectedOutlayKey(row),
    label: getOutlayRowLabel(row),
    refreshKey: refreshDefinition ? refreshDefinition.key : '',
    row
  };
}

function getProtectedOutlayRows(options = {}) {
  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return [];
  const requestedHidden = options.hidden;
  const allRows = Array.from(tbody.querySelectorAll('tr'))
    .map((row) => getProtectedOutlayInfoForRow(row))
    .filter(Boolean);

  const labelTotals = new Map();
  allRows.forEach((item) => {
    labelTotals.set(item.label, (labelTotals.get(item.label) || 0) + 1);
  });
  const labelSeen = new Map();
  allRows.forEach((item) => {
    const total = labelTotals.get(item.label) || 0;
    if (total <= 1) return;
    const next = (labelSeen.get(item.label) || 0) + 1;
    labelSeen.set(item.label, next);
    item.label = `${item.label} (${next})`;
  });

  return allRows.filter((item) => {
    if (requestedHidden === true) return Boolean(item.row.hidden);
    if (requestedHidden === false) return !item.row.hidden;
    return true;
  });
}

function refreshProtectedOutlayRow(key) {
  if (key === 'light-dues') {
    updateLightDuesFromStorage();
    return;
  }
  if (key === 'port-dues') {
    updatePortDuesFromStorage();
    return;
  }
  if (key === 'bunkering') {
    updateBunkeringFromStorage();
    return;
  }
  if (key === 'berthage-quayage') {
    updateBerthageFromStorage();
    return;
  }
  if (key === 'pilotage') {
    updatePilotageFromStorage();
    return;
  }
  if (key === 'pilot-boat') {
    updatePilotBoatFromStorage();
    return;
  }
  if (key === 'mooring') {
    updateMooringFromStorage();
    return;
  }
  if (key === 'towage') {
    updateTowageFromStorage();
  }
}

function restoreProtectedOutlayRow(item) {
  if (!item || !item.row) return;
  item.row.hidden = false;
  delete item.row.dataset.outlayHidden;
  if (item.refreshKey) {
    refreshProtectedOutlayRow(item.refreshKey);
  }
  decorateMoneyEditCells();
  recalcOutlayTotals();
  saveIndexState();
}

function showSelectDialog(config = {}) {
  const titleText = String(config.title || 'Select option');
  const labelText = String(config.label || 'Choose one option');
  const options = Array.isArray(config.options) ? config.options.filter(Boolean) : [];
  const confirmLabel = String(config.confirmLabel || 'Confirm');
  const cancelLabel = String(config.cancelLabel || 'Cancel');
  const alternateLabel = config.alternateLabel ? String(config.alternateLabel) : '';
  const invalidSelectionMessage = String(config.invalidSelectionMessage || 'Select an option.');

  if (options.length === 0) {
    return Promise.resolve({ type: 'cancel' });
  }

  return new Promise((resolve) => {
    const uniqueId = `selectDialog_${Date.now()}_${Math.floor(Math.random() * 10000)}`;
    const backdrop = document.createElement('div');
    backdrop.className = 'app-select-dialog-backdrop';

    const dialog = document.createElement('div');
    dialog.className = 'app-select-dialog';
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-modal', 'true');
    dialog.setAttribute('aria-labelledby', `${uniqueId}_title`);
    dialog.setAttribute('aria-describedby', `${uniqueId}_label`);

    const title = document.createElement('h3');
    title.id = `${uniqueId}_title`;
    title.textContent = titleText;
    dialog.appendChild(title);

    const label = document.createElement('label');
    label.id = `${uniqueId}_label`;
    label.setAttribute('for', `${uniqueId}_select`);
    label.textContent = labelText;
    dialog.appendChild(label);

    const select = document.createElement('select');
    select.id = `${uniqueId}_select`;
    select.className = 'app-select-dialog-select';
    options.forEach((option) => {
      const node = document.createElement('option');
      node.value = String(option.value ?? '');
      node.textContent = String(option.label ?? option.value ?? '');
      select.appendChild(node);
    });
    dialog.appendChild(select);

    const actions = document.createElement('div');
    actions.className = 'app-select-dialog-actions';

    const cancelButton = document.createElement('button');
    cancelButton.type = 'button';
    cancelButton.className = 'app-select-dialog-btn app-select-dialog-btn-neutral';
    cancelButton.dataset.action = 'cancel';
    cancelButton.textContent = cancelLabel;
    actions.appendChild(cancelButton);

    if (alternateLabel) {
      const alternateButton = document.createElement('button');
      alternateButton.type = 'button';
      alternateButton.className = 'app-select-dialog-btn app-select-dialog-btn-neutral';
      alternateButton.dataset.action = 'alternate';
      alternateButton.textContent = alternateLabel;
      actions.appendChild(alternateButton);
    }

    const confirmButton = document.createElement('button');
    confirmButton.type = 'button';
    confirmButton.className = 'app-select-dialog-btn app-select-dialog-btn-primary';
    confirmButton.dataset.action = 'confirm';
    confirmButton.textContent = confirmLabel;
    actions.appendChild(confirmButton);

    dialog.appendChild(actions);
    backdrop.appendChild(dialog);
    document.body.appendChild(backdrop);

    let isClosed = false;
    const cleanup = () => {
      if (isClosed) return;
      isClosed = true;
      document.removeEventListener('keydown', onKeyDown, true);
      backdrop.remove();
    };

    const finish = (result) => {
      cleanup();
      resolve(result);
    };

    const confirmSelection = () => {
      const value = String(select.value || '');
      if (!value) {
        window.alert(invalidSelectionMessage);
        return;
      }
      finish({ type: 'confirm', value });
    };

    const onKeyDown = (event) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        finish({ type: 'cancel' });
        return;
      }
      if (event.key === 'Enter' && event.target === select) {
        event.preventDefault();
        confirmSelection();
      }
    };

    backdrop.addEventListener('click', (event) => {
      if (event.target === backdrop) {
        finish({ type: 'cancel' });
      }
    });

    actions.addEventListener('click', (event) => {
      const button = event.target.closest('button[data-action]');
      if (!button) return;
      const action = button.dataset.action;
      if (action === 'cancel') {
        finish({ type: 'cancel' });
        return;
      }
      if (action === 'alternate') {
        finish({ type: 'alternate' });
        return;
      }
      if (action === 'confirm') {
        confirmSelection();
      }
    });

    document.addEventListener('keydown', onKeyDown, true);
    requestAnimationFrame(() => {
      select.focus();
    });
  });
}

function promptProtectedOutlayRestore(hiddenRows) {
  if (!Array.isArray(hiddenRows) || hiddenRows.length === 0) {
    return Promise.resolve({ type: 'blank' });
  }

  return showSelectDialog({
    title: 'Restore hidden row',
    label: 'Choose row to return',
    options: hiddenRows.map((item) => ({ value: item.key, label: item.label })),
    confirmLabel: 'Return row',
    alternateLabel: 'Add blank row',
    cancelLabel: 'Cancel',
    invalidSelectionMessage: 'Select a row to return.'
  }).then((result) => {
    if (result.type === 'alternate') {
      return { type: 'blank' };
    }
    if (result.type !== 'confirm') {
      return { type: 'cancel' };
    }
    const item = hiddenRows.find((entry) => entry.key === result.value);
    if (!item) {
      return { type: 'cancel' };
    }
    return { type: 'restore', item };
  });
}

function decorateOutlayRows() {
  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return;

  const normalizeOutlayActionIcons = (row) => {
    if (!row) return;
    const handle = row.querySelector('.row-handle');
    if (handle) {
      handle.setAttribute('draggable', 'true');
      const handleImg = handle.querySelector('img');
      if (handleImg) {
        handleImg.setAttribute('draggable', 'false');
        handleImg.setAttribute('aria-hidden', 'true');
      }
    }
    row.querySelectorAll('.row-remove img, .row-edit img').forEach((img) => {
      img.setAttribute('draggable', 'false');
    });
  };

  tbody.querySelectorAll('tr').forEach((row) => {
    const descCell = row.querySelector('td.desc');
    if (!descCell) return;
    if (row.dataset.row === 'bank-charges') return;
    if (descCell.querySelector('.row-item')) {
      normalizeOutlayActionIcons(row);
      return;
    }

    const input = descCell.querySelector('input, textarea');
    if (!input) return;

    const wrapper = document.createElement('div');
    wrapper.className = 'row-item';

    descCell.removeChild(input);
    const handle = document.createElement('button');
    handle.type = 'button';
    handle.className = 'row-handle';
    handle.setAttribute('draggable', 'true');
    handle.setAttribute('aria-label', 'Move row');
    const handleImg = document.createElement('img');
    handleImg.src = 'assets/icons/grip-vertical.svg';
    handleImg.alt = '';
    handleImg.setAttribute('draggable', 'false');
    handle.appendChild(handleImg);
    wrapper.appendChild(handle);

    wrapper.appendChild(input);

    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'row-remove';
    removeBtn.setAttribute('aria-label', 'Remove row');
    const removeImg = document.createElement('img');
    removeImg.src = 'assets/icons/trash.svg';
    removeImg.alt = '';
    removeImg.setAttribute('draggable', 'false');
    removeBtn.appendChild(removeImg);
    wrapper.appendChild(removeBtn);

    descCell.appendChild(wrapper);
    normalizeOutlayActionIcons(row);
  });
}

function autoResizeTextarea(textarea) {
  if (!textarea) return;
  textarea.style.height = 'auto';
  textarea.style.height = `${textarea.scrollHeight}px`;
}

function stripTowageCounts(text) {
  return String(text || '').replace(/\s*\(Arrival tugs:.*?Departure tugs:.*?\)\s*$/i, '').trim();
}

function getTowageCalculationFromState(state) {
  if (!state || typeof state !== 'object' || !Array.isArray(state.tugs)) return null;
  if (state.tugs.length === 0) {
    return {
      arrivalTotal: 0,
      departureTotal: 0,
      grandTotal: 0,
      arrivalCount: 0,
      departureCount: 0
    };
  }

  const storedGt = parseMoneyValue(state.gt);
  const sharedGt = parseMoneyValue(safeStorageGet(STORAGE_KEYS.gt));
  const gt = sharedGt > 0 ? sharedGt : storedGt;
  const tariff = getTariffFromGT(gt);
  const discountPercent = Boolean(state.discountEnabled)
    ? Math.min(Math.max(parsePercentValue(state.discountPercent), 0), 100)
    : 0;
  const discountFactor = discountPercent > 0 ? (1 - (discountPercent / 100)) : 1;
  const effectiveTariff = tariff * discountFactor;
  const standbyCharge = Boolean(state.standbyEnabled)
    ? getTugStandbyCharge(state.standbyHours, gt)
    : 0;
  if (!Number.isFinite(tariff) || tariff <= 0) {
    let arrivalCountOnly = 0;
    let departureCountOnly = 0;
    state.tugs.forEach((tug) => {
      const operation = tug?.op === 'departure' ? 'departure' : 'arrival';
      if (operation === 'departure') departureCountOnly += 1;
      else arrivalCountOnly += 1;
    });
    return {
      arrivalTotal: 0,
      departureTotal: 0,
      grandTotal: 0,
      arrivalCount: arrivalCountOnly,
      departureCount: departureCountOnly,
      standbyCharge: 0
    };
  }

  const globalImo = getGlobalImoTransportState();
  let arrivalTotal = 0;
  let departureTotal = 0;
  let arrivalCount = 0;
  let departureCount = 0;

  state.tugs.forEach((tug) => {
    const operation = tug?.op === 'departure' ? 'departure' : 'arrival';
    if (operation === 'departure') departureCount += 1;
    else arrivalCount += 1;
    const calculation = getTugChargeBreakdown({
      operation,
      voyage: tug?.voyage,
      assist: tug?.assist,
      voyageOT: tug?.voy_ot,
      assistOT: tug?.assist_ot,
      imo: globalImo === null ? tug?.imo : globalImo,
      lines: tug?.lines,
      kw: tug?.kw
    }, effectiveTariff);
    const total = calculation.total;

    if (operation === 'departure') departureTotal += total;
    else arrivalTotal += total;
  });

  return {
    arrivalTotal,
    departureTotal,
    grandTotal: arrivalTotal + departureTotal + standbyCharge,
    arrivalCount,
    departureCount,
    discountPercent,
    standbyCharge
  };
}

function updateTowageFromStorage() {
  const towageRow = document.querySelector('tr[data-row="towage"]');
  if (!towageRow) return;

  const descInput = towageRow.querySelector('textarea.cell-input.text');
  if (descInput && !descInput.dataset.baseText) {
    descInput.dataset.baseText = 'TOWAGE';
  }

  const pdaKeys = getTugStorageKeys(false);
  const sailingKeys = getTugStorageKeys(true);

  const pdaCalc = getTowageGrandTotalFromStoredState(false);
  const sailingCalc = getTowageGrandTotalFromStoredState(true);
  if (pdaCalc) {
    safeStorageSet(pdaKeys.towageTotal, pdaCalc.grandTotal.toFixed(2));
    safeStorageSet(pdaKeys.towageArrivalCount, pdaCalc.arrivalCount);
    safeStorageSet(pdaKeys.towageDepartureCount, pdaCalc.departureCount);
  }
  if (sailingCalc) {
    safeStorageSet(sailingKeys.towageTotal, sailingCalc.grandTotal.toFixed(2));
    safeStorageSet(sailingKeys.towageArrivalCount, sailingCalc.arrivalCount);
    safeStorageSet(sailingKeys.towageDepartureCount, sailingCalc.departureCount);
  }

  const totalRaw = safeStorageGet(pdaKeys.towageTotal);
  const sailingTotalRaw = safeStorageGet(sailingKeys.towageTotal);
  const arrivalCountRaw = safeStorageGet(pdaKeys.towageArrivalCount);
  const departureCountRaw = safeStorageGet(pdaKeys.towageDepartureCount);

  let arrivalCount = Number(arrivalCountRaw);
  let departureCount = Number(departureCountRaw);
  let totalValue = Number(totalRaw);
  let sailingTotalValue = Number(sailingTotalRaw);

  const hasStoredCounts = Number.isFinite(arrivalCount) || Number.isFinite(departureCount);
  const hasStoredTotals = Number.isFinite(totalValue) || Number.isFinite(sailingTotalValue);
  if ((!Number.isFinite(arrivalCount) || !Number.isFinite(departureCount)) && pdaCalc) {
    arrivalCount = pdaCalc.arrivalCount;
    departureCount = pdaCalc.departureCount;
    safeStorageSet(pdaKeys.towageArrivalCount, arrivalCount);
    safeStorageSet(pdaKeys.towageDepartureCount, departureCount);
  }
  if (!hasStoredCounts && !hasStoredTotals) {
    if (descInput) {
      const baseText = descInput.dataset.baseText || 'TOWAGE';
      descInput.value = baseText;
      autoResizeTextarea(descInput);
    }
    const pdaInput = towageRow.querySelector('td:nth-child(2) input');
    const sailingInput = towageRow.querySelector('td:nth-child(3) input');
    if (pdaInput) {
      pdaInput.value = formatMoneyValue(0);
      delete pdaInput.dataset.rawValue;
    }
    if (sailingInput) sailingInput.value = formatMoneyValue(0);
    recalcOutlayTotals();
    saveIndexState();
    return;
  }

  if (descInput) {
    descInput.dataset.baseText = 'TOWAGE';
    const safeArrival = Number.isFinite(arrivalCount) ? arrivalCount : 0;
    const safeDeparture = Number.isFinite(departureCount) ? departureCount : 0;
    descInput.value = `TOWAGE (Arrival tugs: ${safeArrival}, Departure tugs: ${safeDeparture})`;
    autoResizeTextarea(descInput);
  }

  if (!Number.isFinite(totalValue) && pdaCalc) {
    totalValue = pdaCalc.grandTotal;
    safeStorageSet(pdaKeys.towageTotal, totalValue.toFixed(2));
  }
  const pdaInput = towageRow.querySelector('td:nth-child(2) input');
  const sailingInput = towageRow.querySelector('td:nth-child(3) input');

  if (Number.isFinite(totalValue) && pdaInput) {
    const formattedTowage = formatMoneyValue(totalValue);
    pdaInput.value = formattedTowage;
    if (isPdaRoundingEnabled()) {
      pdaInput.dataset.rawValue = formattedTowage;
    } else {
      delete pdaInput.dataset.rawValue;
    }
  }
  if (!Number.isFinite(sailingTotalValue) && sailingCalc) {
    sailingTotalValue = sailingCalc.grandTotal;
    safeStorageSet(sailingKeys.towageTotal, sailingTotalValue.toFixed(2));
  }
  if (Number.isFinite(sailingTotalValue) && sailingInput) {
    sailingInput.value = formatMoneyValue(sailingTotalValue);
  }
  if (Number.isFinite(totalValue) || Number.isFinite(sailingTotalValue)) {
    recalcOutlayTotals();
    saveIndexState();
  }
}

function findLightDuesRow() {
  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return null;
  return Array.from(tbody.querySelectorAll('tr')).find((row) => {
    const descField = row.querySelector('td.desc textarea, td.desc input');
    const desc = String(descField?.value || '').trim().toUpperCase();
    return desc.startsWith('LIGHT DUES');
  }) || null;
}

function getCurrentGtValue() {
  const gtInput = document.getElementById('grossTonnage');
  const raw = gtInput ? gtInput.value : safeStorageGet(STORAGE_KEYS.gt);
  return parseMoneyValue(raw);
}

function getCurrentQuantityValue() {
  const quantityInput = document.getElementById('quantityInput');
  const raw = quantityInput ? quantityInput.value : safeStorageGet(STORAGE_KEYS.quantity);
  return parseMoneyValue(raw);
}

function getCurrentLengthOverallRawValue() {
  const lengthInput = document.getElementById('lengthOverall');
  if (lengthInput) return String(lengthInput.value || '').trim();
  const state = parseStoredJson(safeStorageGet(STORAGE_KEYS.indexState));
  if (state && state.fields && typeof state.fields.lengthOverall === 'string') {
    return String(state.fields.lengthOverall || '').trim();
  }
  return '';
}

function getCurrentLengthOverallValue() {
  return parseMoneyValue(getCurrentLengthOverallRawValue());
}

function parseStoredJson(raw) {
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch (error) {
    return null;
  }
}

function getSelectedPeriodFromState(state) {
  if (!state || typeof state !== 'object') return '30';
  if (state.period12 === true) return '12';
  return '30';
}

function getLightDuesCalculationFromState(state, gt) {
  if (!state || typeof state !== 'object') return null;
  const type = typeof state.type === 'string' ? state.type : LIGHT_DUES_DEFAULT_TYPE;
  const tierBand = state.tierBand || state.bulkBand || '';
  const tariff = getLightDuesTariff(type, tierBand, gt);
  const selectedPeriod = getSelectedPeriodFromState(state);
  if (!tariff) return null;

  if (state.validLightDues === true) {
    return {
      tariffLabel: selectedPeriod === '12' ? tariff.label12 : tariff.label30,
      amount: 0
    };
  }

  if (!Number.isFinite(gt) || gt <= 0) return null;

  if (selectedPeriod === '12') {
    return {
      tariffLabel: tariff.label12,
      amount: gt * tariff.rate12
    };
  }
  return {
    tariffLabel: tariff.label30,
    amount: gt * tariff.rate30
  };
}

function getCoefficientFromDescription(descValue) {
  const match = String(descValue || '').match(/EUR\s*([0-9.,]+)\s*x/i);
  if (!match || !match[1]) return null;
  const coefficient = parseMoneyValue(match[1]);
  if (!Number.isFinite(coefficient) || coefficient <= 0) return null;
  return {
    value: coefficient,
    label: match[1].replace(/\s+/g, '')
  };
}

function updateLightDuesFromStorage() {
  const lightDuesRow = findLightDuesRow();
  if (!lightDuesRow) return;

  let changed = false;
  const gt = getCurrentGtValue();
  const hasValidGt = Number.isFinite(gt) && gt > 0;
  const descInput = lightDuesRow.querySelector('td.desc textarea, td.desc input');
  const pdaState = parseStoredJson(safeStorageGet(STORAGE_KEYS.lightDuesState));
  const sailingState = parseStoredJson(safeStorageGet(STORAGE_KEYS.lightDuesStateSailing));
  const pdaCalc = getLightDuesCalculationFromState(pdaState, gt);
  const sailingCalc = getLightDuesCalculationFromState(sailingState, gt);
  const descCoefficient = getCoefficientFromDescription(descInput ? descInput.value : '');

  const pdaInput = lightDuesRow.querySelector('td:nth-child(2) input.cell-input.money');
  let pdaTariffLabel = safeStorageGet(STORAGE_KEYS.lightDuesTariffPda);
  let pdaAmount = Number(safeStorageGet(STORAGE_KEYS.lightDuesAmountPda));
  if (pdaCalc) {
    pdaTariffLabel = pdaCalc.tariffLabel;
    pdaAmount = pdaCalc.amount;
    safeStorageSet(STORAGE_KEYS.lightDuesTariffPda, pdaTariffLabel);
    safeStorageSet(STORAGE_KEYS.lightDuesAmountPda, pdaAmount.toFixed(2));
  } else if (hasValidGt && descCoefficient) {
    pdaTariffLabel = descCoefficient.label;
    pdaAmount = gt * descCoefficient.value;
    safeStorageSet(STORAGE_KEYS.lightDuesTariffPda, pdaTariffLabel);
    safeStorageSet(STORAGE_KEYS.lightDuesAmountPda, pdaAmount.toFixed(2));
  }
  if (!hasValidGt && !(pdaState && pdaState.validLightDues === true)) {
    pdaAmount = 0;
    safeStorageRemove(STORAGE_KEYS.lightDuesAmountPda);
  }

  if (pdaTariffLabel && descInput) {
    const nextDesc = `LIGHT DUES (EUR ${pdaTariffLabel} x vsl's GRT)`;
    if (descInput.value !== nextDesc) {
      descInput.value = nextDesc;
      if (descInput.tagName === 'TEXTAREA') autoResizeTextarea(descInput);
      changed = true;
    }
  }

  if (pdaInput && Number.isFinite(pdaAmount)) {
    const formatted = formatMoneyValue(pdaAmount);
    if (pdaInput.value !== formatted) {
      pdaInput.value = formatted;
      if (isPdaRoundingEnabled()) pdaInput.dataset.rawValue = formatted;
      else delete pdaInput.dataset.rawValue;
      changed = true;
    }
  }

  const sailingInput = lightDuesRow.querySelector('td:nth-child(3) input.cell-input.money');
  let sailingAmount = Number(safeStorageGet(STORAGE_KEYS.lightDuesAmountSailing));
  if (sailingCalc) {
    sailingAmount = sailingCalc.amount;
    safeStorageSet(STORAGE_KEYS.lightDuesTariffSailing, sailingCalc.tariffLabel);
    safeStorageSet(STORAGE_KEYS.lightDuesAmountSailing, sailingAmount.toFixed(2));
  } else if (hasValidGt) {
    if (sailingState && sailingCalc === null) {
      // keep existing value if sailing state is invalid.
    } else if (pdaCalc && (!pdaState || pdaState.validLightDues !== true)) {
      sailingAmount = pdaCalc.amount;
      safeStorageSet(STORAGE_KEYS.lightDuesTariffSailing, pdaCalc.tariffLabel);
      safeStorageSet(STORAGE_KEYS.lightDuesAmountSailing, sailingAmount.toFixed(2));
    } else if (descCoefficient) {
      sailingAmount = gt * descCoefficient.value;
      safeStorageSet(STORAGE_KEYS.lightDuesAmountSailing, sailingAmount.toFixed(2));
    }
  }
  if (!hasValidGt && !(sailingState && sailingState.validLightDues === true)) {
    sailingAmount = 0;
    safeStorageRemove(STORAGE_KEYS.lightDuesAmountSailing);
  }

  if (sailingInput && Number.isFinite(sailingAmount)) {
    const formatted = formatMoneyValue(sailingAmount);
    if (sailingInput.value !== formatted) {
      sailingInput.value = formatted;
      changed = true;
    }
  }

  if (changed) {
    recalcOutlayTotals();
    saveIndexState();
  }
}

function findPortDuesRow() {
  const byDataRow = document.querySelector('tr[data-row="port-dues"]');
  if (byDataRow) return byDataRow;
  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return null;
  return Array.from(tbody.querySelectorAll('tr')).find((row) => {
    const descField = row.querySelector('td.desc textarea, td.desc input');
    const desc = String(descField?.value || '').trim().toUpperCase();
    return desc.startsWith('PORT DUES');
  }) || null;
}

function findBunkeringRow() {
  const byDataRow = document.querySelector('tr[data-row="bunkering"], tr[data-row="bunker"]');
  if (byDataRow) {
    if (byDataRow.dataset.row === 'bunker') byDataRow.dataset.row = 'bunkering';
    return byDataRow;
  }
  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return null;
  return Array.from(tbody.querySelectorAll('tr')).find((row) => {
    const descField = row.querySelector('td.desc textarea, td.desc input');
    const desc = String(descField?.value || '').trim().toUpperCase();
    return desc.startsWith('BUNKER');
  }) || null;
}

function createBunkeringRow() {
  const row = document.createElement('tr');
  row.dataset.row = 'bunkering';
  row.innerHTML = `
    <td class="desc"><textarea class="cell-input text" rows="1">${BUNKER_ROW_DESCRIPTION}</textarea></td>
    <td><input class="cell-input money" placeholder="0,00" /></td>
    <td><input class="cell-input money" placeholder="0,00" /></td>
  `;
  return row;
}

function findPilotageRow() {
  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return null;
  return Array.from(tbody.querySelectorAll('tr')).find((row) => {
    const descField = row.querySelector('td.desc textarea, td.desc input');
    const desc = String(descField?.value || '').trim().toUpperCase();
    return row.dataset.row === 'pilotage' || desc.startsWith('PILOTAGE');
  }) || null;
}

function ensureBunkeringRow() {
  const existing = findBunkeringRow();
  if (existing) return existing;

  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return null;

  const row = createBunkeringRow();
  const pilotageRow = findPilotageRow();
  if (pilotageRow && pilotageRow.parentElement === tbody) {
    tbody.insertBefore(row, pilotageRow);
  } else {
    const portDuesRow = findPortDuesRow();
    if (portDuesRow && portDuesRow.parentElement === tbody) {
      tbody.insertBefore(row, portDuesRow.nextElementSibling);
    } else {
      const bankRow = tbody.querySelector('tr[data-row="bank-charges"]');
      if (bankRow) tbody.insertBefore(row, bankRow);
      else tbody.appendChild(row);
    }
  }

  decorateOutlayRows();
  wrapMoneyFields();
  decorateMoneyEditCells();
  const descInput = row.querySelector('textarea.cell-input.text');
  if (descInput) autoResizeTextarea(descInput);
  return row;
}

function updateBunkeringFromStorage() {
  const bunkeringRow = ensureBunkeringRow();
  if (!bunkeringRow) return;

  let changed = false;
  const descInput = bunkeringRow.querySelector('td.desc textarea, td.desc input');
  if (descInput && descInput.value !== BUNKER_ROW_DESCRIPTION) {
    descInput.value = BUNKER_ROW_DESCRIPTION;
    if (descInput.tagName === 'TEXTAREA') autoResizeTextarea(descInput);
    changed = true;
  }

  const pdaInput = bunkeringRow.querySelector('td:nth-child(2) input.cell-input.money');
  const pdaBunkeringRaw = Number(safeStorageGet(STORAGE_KEYS.portDuesBunkeringAmountPda));
  const pdaBunkering = Number.isFinite(pdaBunkeringRaw) ? pdaBunkeringRaw : 0;
  if (pdaInput) {
    const formatted = formatMoneyValue(pdaBunkering);
    if (pdaInput.value !== formatted) {
      pdaInput.value = formatted;
      if (isPdaRoundingEnabled()) pdaInput.dataset.rawValue = formatted;
      else delete pdaInput.dataset.rawValue;
      changed = true;
    }
  }

  const sailingInput = bunkeringRow.querySelector('td:nth-child(3) input.cell-input.money');
  const sailingBunkeringRaw = Number(safeStorageGet(STORAGE_KEYS.portDuesBunkeringAmountSailing));
  const sailingBunkering = Number.isFinite(sailingBunkeringRaw) ? sailingBunkeringRaw : 0;
  if (sailingInput) {
    const formatted = formatMoneyValue(sailingBunkering);
    if (sailingInput.value !== formatted) {
      sailingInput.value = formatted;
      changed = true;
    }
  }

  if (changed) {
    recalcOutlayTotals();
    saveIndexState();
  }
}

function findBerthageQuayageRow() {
  const byDataRow = document.querySelector('tr[data-row="berthage-quayage"]');
  if (byDataRow) return byDataRow;
  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return null;
  return Array.from(tbody.querySelectorAll('tr')).find((row) => {
    const descField = row.querySelector('td.desc textarea, td.desc input');
    const desc = String(descField?.value || '').trim().toUpperCase();
    return desc.startsWith('BERHAGE/QUAYAGE') || desc.startsWith('BERTHAGE/QUAYAGE');
  }) || null;
}

function createBerthageQuayageRow() {
  const row = document.createElement('tr');
  row.dataset.row = 'berthage-quayage';
  row.innerHTML = `
    <td class="desc"><textarea class="cell-input text" rows="1">${BERTHAGE_QUAYAGE_ROW_DESCRIPTION}</textarea></td>
    <td><input class="cell-input money" placeholder="0,00" /></td>
    <td><input class="cell-input money" placeholder="0,00" /></td>
  `;
  return row;
}

function ensureBerthageQuayageRow() {
  const existing = findBerthageQuayageRow();
  if (existing) return existing;

  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return null;

  const row = createBerthageQuayageRow();
  const pilotageRow = findPilotageRow();
  if (pilotageRow && pilotageRow.parentElement === tbody) {
    tbody.insertBefore(row, pilotageRow);
  } else {
    const bankRow = tbody.querySelector('tr[data-row="bank-charges"]');
    if (bankRow) tbody.insertBefore(row, bankRow);
    else tbody.appendChild(row);
  }

  decorateOutlayRows();
  wrapMoneyFields();
  decorateMoneyEditCells();
  const descInput = row.querySelector('textarea.cell-input.text');
  if (descInput) autoResizeTextarea(descInput);
  return row;
}

function findBerthageAnchorageExtraRow() {
  return document.querySelector('tr[data-row="berthage-anchorage-extra"]');
}

function createBerthageAnchorageExtraRow() {
  const row = document.createElement('tr');
  row.dataset.row = 'berthage-anchorage-extra';
  row.innerHTML = `
    <td class="desc"><textarea class="cell-input text" rows="1">ANCHORAGE</textarea></td>
    <td><input class="cell-input money" placeholder="0,00" /></td>
    <td><input class="cell-input money" placeholder="0,00" /></td>
  `;
  return row;
}

function ensureBerthageAnchorageExtraRow(baseRow) {
  if (!baseRow || !baseRow.parentElement) return null;
  const tbody = baseRow.parentElement;
  let extraRow = findBerthageAnchorageExtraRow();
  if (!extraRow) {
    extraRow = createBerthageAnchorageExtraRow();
  }

  const targetNext = baseRow.nextElementSibling;
  if (targetNext !== extraRow) {
    tbody.insertBefore(extraRow, targetNext);
  }

  extraRow.hidden = false;
  delete extraRow.dataset.outlayHidden;
  decorateOutlayRows();
  wrapMoneyFields();
  decorateMoneyEditCells();
  const descInput = extraRow.querySelector('textarea.cell-input.text');
  if (descInput) autoResizeTextarea(descInput);
  return extraRow;
}

function removeBerthageAnchorageExtraRow() {
  const extraRow = findBerthageAnchorageExtraRow();
  if (!extraRow) return false;
  extraRow.remove();
  return true;
}

function formatDaysLabel(daysValue) {
  if (!Number.isFinite(daysValue) || daysValue <= 0) return '0';
  const rounded = Math.round(daysValue);
  if (Math.abs(daysValue - rounded) < 1e-9) return String(rounded);
  return String(daysValue).replace('.', ',');
}

function setMoneyInputAmount(input, amount, allowRoundingDataset) {
  if (!input) return false;
  const formatted = formatMoneyValue(amount);
  if (input.value === formatted) {
    if (allowRoundingDataset) {
      if (isPdaRoundingEnabled()) input.dataset.rawValue = formatted;
      else delete input.dataset.rawValue;
    }
    return false;
  }
  input.value = formatted;
  if (allowRoundingDataset) {
    if (isPdaRoundingEnabled()) input.dataset.rawValue = formatted;
    else delete input.dataset.rawValue;
  }
  return true;
}

function getBerthageAnchorageCalculationFromState(state, lengthOverall) {
  if (!state || typeof state !== 'object') {
    return {
      berthageAmount: 0,
      anchorageAmount: 0,
      totalAmount: 0
    };
  }
  const lengthValue = Number.isFinite(lengthOverall) && lengthOverall > 0 ? lengthOverall : 0;
  const berthageDays = parseMoneyValue(state.berthageDays);
  const anchorageDays = parseMoneyValue(state.anchorageDays);
  const berthageAmount = state.berthageEnabled && lengthValue > 0 && berthageDays > 0
    ? lengthValue * BERTHAGE_TARIFF_RATE * berthageDays
    : 0;
  const anchorageAmount = state.anchorageEnabled && lengthValue > 0 && anchorageDays > 0
    ? lengthValue * ANCHORAGE_TARIFF_RATE * anchorageDays
    : 0;

  return {
    berthageAmount,
    anchorageAmount,
    totalAmount: berthageAmount + anchorageAmount
  };
}

function updateBerthageFromStorage() {
  const row = ensureBerthageQuayageRow();
  if (!row) return;

  let changed = false;
  const lengthOverall = getCurrentLengthOverallValue();
  const pdaState = parseStoredJson(safeStorageGet(STORAGE_KEYS.berthageState));
  const sailingState = parseStoredJson(safeStorageGet(STORAGE_KEYS.berthageStateSailing));
  const pdaCalc = getBerthageAnchorageCalculationFromState(pdaState, lengthOverall);
  const sailingCalc = getBerthageAnchorageCalculationFromState(sailingState, lengthOverall);
  safeStorageSet(STORAGE_KEYS.berthageAmountPda, pdaCalc.totalAmount.toFixed(2));
  safeStorageSet(STORAGE_KEYS.berthageAmountSailing, sailingCalc.totalAmount.toFixed(2));

  const pdaBerthageEnabled = Boolean(pdaState && pdaState.berthageEnabled);
  const pdaAnchorageEnabled = Boolean(pdaState && pdaState.anchorageEnabled);
  const pdaBerthageDays = parseMoneyValue(pdaState && pdaState.berthageDays);
  const pdaAnchorageDays = parseMoneyValue(pdaState && pdaState.anchorageDays);

  const descInput = row.querySelector('td.desc textarea, td.desc input');
  const primaryPdaInput = row.querySelector('td:nth-child(2) input.cell-input.money');
  const primarySailingInput = row.querySelector('td:nth-child(3) input.cell-input.money');

  let primaryDesc = BERTHAGE_QUAYAGE_ROW_DESCRIPTION;
  let primaryPdaAmount = 0;
  let primarySailingAmount = 0;
  const showAnchorageExtraRow = pdaBerthageEnabled && pdaAnchorageEnabled;

  if (pdaBerthageEnabled) {
    primaryDesc = `BERTHAGE/QUAYAGE (2 EUR x ${formatDaysLabel(pdaBerthageDays)} DAYS)`;
    primaryPdaAmount = pdaCalc.berthageAmount;
    primarySailingAmount = sailingCalc.berthageAmount;
  } else if (pdaAnchorageEnabled) {
    primaryDesc = `ANCHORAGE (1 EUR x ${formatDaysLabel(pdaAnchorageDays)} DAYS)`;
    primaryPdaAmount = pdaCalc.anchorageAmount;
    primarySailingAmount = sailingCalc.anchorageAmount;
  }

  if (descInput && descInput.value !== primaryDesc) {
    descInput.value = primaryDesc;
    if (descInput.tagName === 'TEXTAREA') autoResizeTextarea(descInput);
    changed = true;
  }
  changed = setMoneyInputAmount(primaryPdaInput, primaryPdaAmount, true) || changed;
  changed = setMoneyInputAmount(primarySailingInput, primarySailingAmount, false) || changed;

  if (showAnchorageExtraRow) {
    const extraRow = ensureBerthageAnchorageExtraRow(row);
    if (extraRow) {
      const extraDescInput = extraRow.querySelector('td.desc textarea, td.desc input');
      const extraDesc = `ANCHORAGE (1 EUR x ${formatDaysLabel(pdaAnchorageDays)} DAYS)`;
      if (extraDescInput && extraDescInput.value !== extraDesc) {
        extraDescInput.value = extraDesc;
        if (extraDescInput.tagName === 'TEXTAREA') autoResizeTextarea(extraDescInput);
        changed = true;
      }
      const extraPdaInput = extraRow.querySelector('td:nth-child(2) input.cell-input.money');
      const extraSailingInput = extraRow.querySelector('td:nth-child(3) input.cell-input.money');
      changed = setMoneyInputAmount(extraPdaInput, pdaCalc.anchorageAmount, true) || changed;
      changed = setMoneyInputAmount(extraSailingInput, sailingCalc.anchorageAmount, false) || changed;
    }
  } else if (removeBerthageAnchorageExtraRow()) {
    changed = true;
  }

  if (changed) {
    recalcOutlayTotals();
    saveIndexState();
  }
}

function getPortDuesCalculationFromState(state, quantityValue, options) {
  if (!state || typeof state !== 'object') return null;
  const opts = options && typeof options === 'object' ? options : {};
  const cargoType = getPortDuesCargoTypeConfig(state.cargoType);
  const terminalDischargedEnabled = Boolean(opts.useTerminalDischarged && state.terminalDischargedEnabled);
  const cargoQuantityFromState = terminalDischargedEnabled
    ? parseMoneyValue(state.terminalDischargedQuantity)
    : parseMoneyValue(state.cargoQuantity);
  const cargoQuantity = terminalDischargedEnabled
    ? cargoQuantityFromState
    : (quantityValue > 0 ? quantityValue : cargoQuantityFromState);
  const cargoAmount = cargoQuantity > 0 ? cargoQuantity * cargoType.rate : 0;

  return {
    coefficientLabel: cargoType.label,
    cargoAmount,
    totalAmount: cargoAmount
  };
}

function updatePortDuesFromStorage() {
  const portDuesRow = findPortDuesRow();
  if (!portDuesRow) return;

  let changed = false;
  const quantity = getCurrentQuantityValue();
  const hasValidQuantity = Number.isFinite(quantity) && quantity > 0;
  const pdaState = parseStoredJson(safeStorageGet(STORAGE_KEYS.portDuesState));
  const sailingState = parseStoredJson(safeStorageGet(STORAGE_KEYS.portDuesStateSailing));
  const sailingTerminalDischargedEnabled = Boolean(sailingState && sailingState.terminalDischargedEnabled);
  const pdaCalc = hasValidQuantity ? getPortDuesCalculationFromState(pdaState, quantity) : null;
  const sailingCalc = (hasValidQuantity || sailingTerminalDischargedEnabled)
    ? getPortDuesCalculationFromState(sailingState, quantity, { useTerminalDischarged: true })
    : null;

  const descInput = portDuesRow.querySelector('td.desc textarea, td.desc input');
  const descCoefficient = getCoefficientFromDescription(descInput ? descInput.value : '');
  if (descInput) {
    let coefficientLabel = '';
    if (pdaCalc) {
      coefficientLabel = pdaCalc.coefficientLabel;
    } else if (pdaState && typeof pdaState.cargoType === 'string') {
      coefficientLabel = getPortDuesCargoTypeConfig(pdaState.cargoType).label;
    } else if (descCoefficient) {
      coefficientLabel = descCoefficient.label;
    }
    if (coefficientLabel) {
      const nextDesc = `PORT DUES (EUR ${coefficientLabel} x cargo / MT)`;
      if (descInput.value !== nextDesc) {
        descInput.value = nextDesc;
        if (descInput.tagName === 'TEXTAREA') autoResizeTextarea(descInput);
        changed = true;
      }
    }
  }

  const pdaInput = portDuesRow.querySelector('td:nth-child(2) input.cell-input.money');
  let pdaAmount = Number(safeStorageGet(STORAGE_KEYS.portDuesCargoAmountPda) ?? safeStorageGet(STORAGE_KEYS.portDuesAmountPda));
  if (pdaCalc) {
    pdaAmount = pdaCalc.totalAmount;
    safeStorageSet(STORAGE_KEYS.portDuesAmountPda, pdaCalc.totalAmount.toFixed(2));
    safeStorageSet(STORAGE_KEYS.portDuesCargoAmountPda, pdaCalc.cargoAmount.toFixed(2));
  } else if (hasValidQuantity && descCoefficient) {
    pdaAmount = quantity * descCoefficient.value;
    safeStorageSet(STORAGE_KEYS.portDuesAmountPda, pdaAmount.toFixed(2));
    safeStorageSet(STORAGE_KEYS.portDuesCargoAmountPda, pdaAmount.toFixed(2));
  }
  if (!hasValidQuantity) {
    pdaAmount = 0;
    safeStorageRemove(STORAGE_KEYS.portDuesAmountPda);
    safeStorageRemove(STORAGE_KEYS.portDuesCargoAmountPda);
  }

  if (pdaInput && Number.isFinite(pdaAmount)) {
    const formatted = formatMoneyValue(pdaAmount);
    if (pdaInput.value !== formatted) {
      pdaInput.value = formatted;
      if (isPdaRoundingEnabled()) pdaInput.dataset.rawValue = formatted;
      else delete pdaInput.dataset.rawValue;
      changed = true;
    }
  }

  const sailingInput = portDuesRow.querySelector('td:nth-child(3) input.cell-input.money');
  let sailingAmount = Number(safeStorageGet(STORAGE_KEYS.portDuesCargoAmountSailing) ?? safeStorageGet(STORAGE_KEYS.portDuesAmountSailing));
  if (sailingCalc) {
    sailingAmount = sailingCalc.totalAmount;
    safeStorageSet(STORAGE_KEYS.portDuesAmountSailing, sailingCalc.totalAmount.toFixed(2));
    safeStorageSet(STORAGE_KEYS.portDuesCargoAmountSailing, sailingCalc.cargoAmount.toFixed(2));
  } else if (hasValidQuantity && !sailingTerminalDischargedEnabled) {
    if (sailingState && sailingCalc === null) {
      // keep existing value if sailing state exists but cannot be calculated.
    } else if (pdaCalc) {
      sailingAmount = pdaCalc.totalAmount;
      safeStorageSet(STORAGE_KEYS.portDuesAmountSailing, sailingAmount.toFixed(2));
      safeStorageSet(STORAGE_KEYS.portDuesCargoAmountSailing, sailingAmount.toFixed(2));
    } else if (descCoefficient) {
      sailingAmount = quantity * descCoefficient.value;
      safeStorageSet(STORAGE_KEYS.portDuesAmountSailing, sailingAmount.toFixed(2));
      safeStorageSet(STORAGE_KEYS.portDuesCargoAmountSailing, sailingAmount.toFixed(2));
    }
  }
  if (!hasValidQuantity && !sailingTerminalDischargedEnabled) {
    sailingAmount = 0;
    safeStorageRemove(STORAGE_KEYS.portDuesAmountSailing);
    safeStorageRemove(STORAGE_KEYS.portDuesCargoAmountSailing);
  }

  if (sailingInput && Number.isFinite(sailingAmount)) {
    const formatted = formatMoneyValue(sailingAmount);
    if (sailingInput.value !== formatted) {
      sailingInput.value = formatted;
      changed = true;
    }
  }

  if (changed) {
    recalcOutlayTotals();
    saveIndexState();
  }
}

function findMooringRow() {
  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return null;
  return Array.from(tbody.querySelectorAll('tr')).find((row) => {
    const descField = row.querySelector('td.desc textarea, td.desc input');
    const desc = String(descField?.value || '').trim().toUpperCase();
    return desc.startsWith('MOORING/UNMOORING');
  }) || null;
}

function findPilotBoatRow() {
  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return null;
  return Array.from(tbody.querySelectorAll('tr')).find((row) => {
    const descField = row.querySelector('td.desc textarea, td.desc input');
    const desc = String(descField?.value || '').trim().toUpperCase();
    return desc.startsWith('PILOT BOAT');
  }) || null;
}

function getPilotBoatGrandTotalFromStoredState(stateKey) {
  if (!stateKey) return null;
  const state = parseStoredJson(safeStorageGet(stateKey));
  const cards = getPilotBoatCardsFromState(state);
  if (!cards.length) return null;
  const grandTotal = cards.reduce((sum, card) => {
    const calculation = getPilotBoatCardCalculation(card);
    return sum + calculation.totalAmount;
  }, 0);
  return Number.isFinite(grandTotal) ? grandTotal : null;
}

function getTowageGrandTotalFromStoredState(isSailing) {
  const calc = getTowageCalculationFromState(getTugsState(Boolean(isSailing)));
  if (!calc || !Number.isFinite(calc.grandTotal)) return null;
  return calc;
}

function getPilotageGrandTotalFromStoredState(stateKey, gtValue, options = {}) {
  if (!stateKey) return null;
  const state = parseStoredJson(safeStorageGet(stateKey));
  const calc = getPilotageCalculationFromState(state, gtValue, options);
  if (!calc || !Number.isFinite(calc.totalAmount)) return null;
  return calc.totalAmount;
}

function getMooringGrandTotalFromStoredState(stateKey, gtValue) {
  if (!stateKey) return null;
  const state = parseStoredJson(safeStorageGet(stateKey));
  const calc = getMooringCalculationFromState(state, gtValue);
  if (!calc || !Number.isFinite(calc.totalAmount)) return null;
  return calc.totalAmount;
}

function updatePilotBoatFromStorage() {
  const pilotBoatRow = findPilotBoatRow();
  if (!pilotBoatRow) return;

  let changed = false;

  const pdaInput = pilotBoatRow.querySelector('td:nth-child(2) input.cell-input.money');
  const pdaStoredRaw = safeStorageGet(STORAGE_KEYS.pilotBoatAmountPda);
  let pdaAmount = Number(pdaStoredRaw);
  if (!Number.isFinite(pdaAmount)) {
    pdaAmount = getPilotBoatGrandTotalFromStoredState(STORAGE_KEYS.pilotBoatState);
    if (Number.isFinite(pdaAmount)) {
      safeStorageSet(STORAGE_KEYS.pilotBoatAmountPda, pdaAmount.toFixed(2));
    }
  }
  if (pdaInput && Number.isFinite(pdaAmount)) {
    const formatted = formatMoneyValue(pdaAmount);
    if (pdaInput.value !== formatted) {
      pdaInput.value = formatted;
      if (isPdaRoundingEnabled()) pdaInput.dataset.rawValue = formatted;
      else delete pdaInput.dataset.rawValue;
      changed = true;
    }
  }

  const sailingInput = pilotBoatRow.querySelector('td:nth-child(3) input.cell-input.money');
  const sailingStoredRaw = safeStorageGet(STORAGE_KEYS.pilotBoatAmountSailing);
  let sailingAmount = Number(sailingStoredRaw);
  if (!Number.isFinite(sailingAmount)) {
    sailingAmount = getPilotBoatGrandTotalFromStoredState(STORAGE_KEYS.pilotBoatStateSailing);
    if (Number.isFinite(sailingAmount)) {
      safeStorageSet(STORAGE_KEYS.pilotBoatAmountSailing, sailingAmount.toFixed(2));
    }
  }
  if (sailingInput && Number.isFinite(sailingAmount)) {
    const formatted = formatMoneyValue(sailingAmount);
    if (sailingInput.value !== formatted) {
      sailingInput.value = formatted;
      changed = true;
    }
  }

  if (changed) {
    recalcOutlayTotals();
    saveIndexState();
  }
}

function updatePilotageFromStorage() {
  const pilotageRow = findPilotageRow();
  if (!pilotageRow) return;

  let changed = false;
  const gt = getCurrentGtValue();
  const hasValidGt = Number.isFinite(gt) && gt > 0;
  const globalImo = getGlobalImoTransportState();
  const pilotageCalcOptions = {};
  if (typeof globalImo === 'boolean') pilotageCalcOptions.forceImo = globalImo;
  const pdaState = parseStoredJson(safeStorageGet(STORAGE_KEYS.pilotageState));
  const sailingState = parseStoredJson(safeStorageGet(STORAGE_KEYS.pilotageStateSailing));
  const pdaCalc = getPilotageCalculationFromState(pdaState, gt, pilotageCalcOptions);
  const sailingCalc = getPilotageCalculationFromState(sailingState, gt, pilotageCalcOptions);

  const pdaInput = pilotageRow.querySelector('td:nth-child(2) input.cell-input.money');
  const pdaStoredRaw = safeStorageGet(STORAGE_KEYS.pilotageAmountPda);
  let pdaAmount = Number(pdaStoredRaw);
  if (pdaCalc) {
    pdaAmount = pdaCalc.totalAmount;
    safeStorageSet(STORAGE_KEYS.pilotageAmountPda, pdaAmount.toFixed(2));
  } else if (!Number.isFinite(pdaAmount)) {
    pdaAmount = getPilotageGrandTotalFromStoredState(STORAGE_KEYS.pilotageState, gt, pilotageCalcOptions);
    if (Number.isFinite(pdaAmount)) {
      safeStorageSet(STORAGE_KEYS.pilotageAmountPda, pdaAmount.toFixed(2));
    }
  }
  if (!hasValidGt) {
    pdaAmount = 0;
    safeStorageRemove(STORAGE_KEYS.pilotageAmountPda);
  }
  if (pdaInput && Number.isFinite(pdaAmount)) {
    const formatted = formatMoneyValue(pdaAmount);
    if (pdaInput.value !== formatted) {
      pdaInput.value = formatted;
      if (isPdaRoundingEnabled()) pdaInput.dataset.rawValue = formatted;
      else delete pdaInput.dataset.rawValue;
      changed = true;
    }
  }

  const sailingInput = pilotageRow.querySelector('td:nth-child(3) input.cell-input.money');
  const sailingStoredRaw = safeStorageGet(STORAGE_KEYS.pilotageAmountSailing);
  let sailingAmount = Number(sailingStoredRaw);
  if (sailingCalc) {
    sailingAmount = sailingCalc.totalAmount;
    safeStorageSet(STORAGE_KEYS.pilotageAmountSailing, sailingAmount.toFixed(2));
  } else if (!Number.isFinite(sailingAmount)) {
    sailingAmount = getPilotageGrandTotalFromStoredState(STORAGE_KEYS.pilotageStateSailing, gt, pilotageCalcOptions);
    if (Number.isFinite(sailingAmount)) {
      safeStorageSet(STORAGE_KEYS.pilotageAmountSailing, sailingAmount.toFixed(2));
    }
  }
  if ((!Number.isFinite(sailingAmount) || sailingAmount <= 0) && pdaCalc) {
    sailingAmount = pdaCalc.totalAmount;
    safeStorageSet(STORAGE_KEYS.pilotageAmountSailing, sailingAmount.toFixed(2));
  }
  if (!hasValidGt) {
    sailingAmount = 0;
    safeStorageRemove(STORAGE_KEYS.pilotageAmountSailing);
  }
  if (sailingInput && Number.isFinite(sailingAmount)) {
    const formatted = formatMoneyValue(sailingAmount);
    if (sailingInput.value !== formatted) {
      sailingInput.value = formatted;
      changed = true;
    }
  }

  if (changed) {
    recalcOutlayTotals();
    saveIndexState();
  }
}

function updateMooringFromStorage() {
  const mooringRow = findMooringRow();
  if (!mooringRow) return;

  let changed = false;
  const gt = getCurrentGtValue();
  const hasValidGt = Number.isFinite(gt) && gt > 0;
  const pdaState = parseStoredJson(safeStorageGet(STORAGE_KEYS.mooringState));
  const sailingState = parseStoredJson(safeStorageGet(STORAGE_KEYS.mooringStateSailing));
  const pdaCalc = getMooringCalculationFromState(pdaState, gt);
  const sailingCalc = getMooringCalculationFromState(sailingState, gt);

  const pdaInput = mooringRow.querySelector('td:nth-child(2) input.cell-input.money');
  let pdaAmount = Number(safeStorageGet(STORAGE_KEYS.mooringAmountPda));
  if (pdaCalc) {
    pdaAmount = pdaCalc.totalAmount;
    safeStorageSet(STORAGE_KEYS.mooringAmountPda, pdaAmount.toFixed(2));
  } else if (!Number.isFinite(pdaAmount)) {
    pdaAmount = getMooringGrandTotalFromStoredState(STORAGE_KEYS.mooringState, gt);
    if (Number.isFinite(pdaAmount)) {
      safeStorageSet(STORAGE_KEYS.mooringAmountPda, pdaAmount.toFixed(2));
    }
  }
  if (!hasValidGt) {
    pdaAmount = 0;
    safeStorageRemove(STORAGE_KEYS.mooringAmountPda);
  }
  if (pdaInput && Number.isFinite(pdaAmount)) {
    const formatted = formatMoneyValue(pdaAmount);
    if (pdaInput.value !== formatted) {
      pdaInput.value = formatted;
      if (isPdaRoundingEnabled()) pdaInput.dataset.rawValue = formatted;
      else delete pdaInput.dataset.rawValue;
      changed = true;
    }
  }

  const sailingInput = mooringRow.querySelector('td:nth-child(3) input.cell-input.money');
  let sailingAmount = Number(safeStorageGet(STORAGE_KEYS.mooringAmountSailing));
  if (sailingCalc) {
    sailingAmount = sailingCalc.totalAmount;
    safeStorageSet(STORAGE_KEYS.mooringAmountSailing, sailingAmount.toFixed(2));
  } else if (!Number.isFinite(sailingAmount)) {
    sailingAmount = getMooringGrandTotalFromStoredState(STORAGE_KEYS.mooringStateSailing, gt);
    if (Number.isFinite(sailingAmount)) {
      safeStorageSet(STORAGE_KEYS.mooringAmountSailing, sailingAmount.toFixed(2));
    }
  }
  if ((!Number.isFinite(sailingAmount) || sailingAmount <= 0) && pdaCalc) {
    sailingAmount = pdaCalc.totalAmount;
    safeStorageSet(STORAGE_KEYS.mooringAmountSailing, sailingAmount.toFixed(2));
  }
  if (!hasValidGt) {
    sailingAmount = 0;
    safeStorageRemove(STORAGE_KEYS.mooringAmountSailing);
  }
  if (sailingInput && Number.isFinite(sailingAmount)) {
    const formatted = formatMoneyValue(sailingAmount);
    if (sailingInput.value !== formatted) {
      sailingInput.value = formatted;
      changed = true;
    }
  }

  if (changed) {
    recalcOutlayTotals();
    saveIndexState();
  }
}

function updateIndexGtFromStorage() {
  const gtInputIndex = document.getElementById('grossTonnage');
  if (!gtInputIndex) return;
  const storedGt = safeStorageGet(STORAGE_KEYS.gt);
  if (!storedGt) return;
  if (gtInputIndex.value !== storedGt) {
    gtInputIndex.value = storedGt;
  }
}

function updateIndexQuantityFromStorage() {
  const quantityInput = document.getElementById('quantityInput');
  if (!quantityInput) return;
  const storedQuantity = safeStorageGet(STORAGE_KEYS.quantity);
  if (storedQuantity === null) return;
  if (quantityInput.value !== storedQuantity) {
    quantityInput.value = storedQuantity;
  }
}

function updateVesselNameFromStorage(targetInput) {
  if (!targetInput) return;
  const storedName = safeStorageGet(STORAGE_KEYS.vesselName);
  if (!storedName) return;
  if (targetInput.value !== storedName) {
    targetInput.value = storedName;
  }
}

function refreshOutlayLayout() {
  document.querySelectorAll('#outlaysBody textarea.cell-input.text').forEach(autoResizeTextarea);
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

function getCurrencySymbol(code) {
  const map = {
    EUR: '€',
    USD: '$',
    GBP: '£',
    HRK: 'kn'
  };
  return map[code] || code;
}

function updateCurrencySymbol(code) {
  const symbol = getCurrencySymbol(code);
  document.querySelectorAll('[data-currency-symbol]').forEach((el) => {
    el.textContent = symbol;
  });
}

function wrapMoneyFields() {
  document.querySelectorAll('input.cell-input.money').forEach((input) => {
    if (input.closest('.money-field')) return;
    const wrapper = document.createElement('div');
    wrapper.className = 'money-field';
    const symbol = document.createElement('span');
    symbol.className = 'currency-symbol';
    symbol.setAttribute('data-currency-symbol', '');
    symbol.textContent = getCurrencySymbol(outlaysCurrency?.value || 'EUR');
    const parent = input.parentNode;
    parent.insertBefore(wrapper, input);
    wrapper.appendChild(symbol);
    wrapper.appendChild(input);
  });
}

function isPdaRoundingEnabled() {
  return Boolean(roundPdaPrices && roundPdaPrices.checked);
}

function isPdaMoneyInput(input) {
  if (!input) return false;
  const cell = input.closest('td');
  return Boolean(cell && cell.cellIndex === 1);
}

function getMoneyInputSourceValue(input, preferredRaw) {
  if (!input) return '';
  const raw = preferredRaw === undefined || preferredRaw === null ? '' : String(preferredRaw).trim();
  if (raw) return raw;
  const current = String(input.value || '').trim();
  if (current) return current;
  return String(input.placeholder || '').trim();
}

function getMoneyInputEnteredValue(input, preferredRaw) {
  if (!input) return '';
  const raw = preferredRaw === undefined || preferredRaw === null ? '' : String(preferredRaw).trim();
  if (raw) return raw;
  const current = String(input.value || '').trim();
  if (current) return current;
  return '';
}

function getRawOrCurrentPdaValue(input) {
  if (!input) return 0;
  const source = getMoneyInputEnteredValue(
    input,
    input.dataset.rawValue !== undefined ? input.dataset.rawValue : input.value
  );
  return parseMoneyValue(source);
}

function focusMoneyFromEdit(button) {
  if (!button) return;
  const cell = button.closest('td');
  if (!cell) return;
  const input = cell.querySelector('input.cell-input.money');
  if (!input) return;
  input.focus();
  if (typeof input.select === 'function') input.select();
}

function getOutlayDescription(row) {
  const descField = row.querySelector('td.desc textarea, td.desc input');
  return String(descField?.value || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

function shouldShowInlineMoneyEdit(row) {
  const desc = getOutlayDescription(row);
  const editablePrefixes = [
    'LIGHT DUES',
    'PORT DUES',
    'BUNKER',
    'BERHAGE/QUAYAGE',
    'BERTHAGE/QUAYAGE',
    'ANCHORAGE',
    'PILOTAGE',
    'PILOT BOAT',
    'MOORING/UNMOORING'
  ];
  return editablePrefixes.some((prefix) => desc.startsWith(prefix));
}

function removeInlineMoneyEdit(cell) {
  const wrapper = cell.querySelector('.money-edit');
  if (!wrapper) return;
  const moneyField = wrapper.querySelector('.money-field');
  if (!moneyField) return;
  wrapper.replaceWith(moneyField);
}

function getInlineMoneyEditConfig(row, columnIndex) {
  const desc = getOutlayDescription(row);
  if (desc.startsWith('LIGHT DUES')) {
    if (columnIndex === 2) {
      return {
        onClick: 'openLightDuesPda()',
        ariaLabel: 'Edit light dues PDA'
      };
    }
    return {
      onClick: 'openLightDuesSailingPda()',
      ariaLabel: 'Edit light dues sailing PDA'
    };
  }
  if (desc.startsWith('PORT DUES')) {
    if (columnIndex === 2) {
      return {
        onClick: 'openPortDuesPda()',
        ariaLabel: 'Edit port dues PDA'
      };
    }
    return {
      onClick: 'openPortDuesSailingPda()',
      ariaLabel: 'Edit port dues sailing PDA'
    };
  }
  if (desc.startsWith('BUNKER')) {
    if (columnIndex === 2) {
      return {
        onClick: 'openBunkeringPda()',
        ariaLabel: 'Edit bunkering PDA'
      };
    }
    return {
      onClick: 'openBunkeringSailingPda()',
      ariaLabel: 'Edit bunkering sailing PDA'
    };
  }
  if (desc.startsWith('BERHAGE/QUAYAGE') || desc.startsWith('BERTHAGE/QUAYAGE')) {
    if (columnIndex === 2) {
      return {
        onClick: 'openBerthageAnchoragePda()',
        ariaLabel: 'Edit berthage/quayage PDA'
      };
    }
    return {
      onClick: 'openBerthageAnchorageSailingPda()',
      ariaLabel: 'Edit berthage/quayage sailing PDA'
    };
  }
  if (desc.startsWith('ANCHORAGE')) {
    if (columnIndex === 2) {
      return {
        onClick: 'openBerthageAnchoragePda()',
        ariaLabel: 'Edit anchorage PDA'
      };
    }
    return {
      onClick: 'openBerthageAnchorageSailingPda()',
      ariaLabel: 'Edit anchorage sailing PDA'
    };
  }
  if (desc.startsWith('PILOTAGE')) {
    if (columnIndex === 2) {
      return {
        onClick: 'openPilotagePda()',
        ariaLabel: 'Edit pilotage PDA'
      };
    }
    return {
      onClick: 'openPilotageSailingPda()',
      ariaLabel: 'Edit pilotage sailing PDA'
    };
  }
  if (desc.startsWith('PILOT BOAT')) {
    if (columnIndex === 2) {
      return {
        onClick: 'openPilotBoatPda()',
        ariaLabel: 'Edit pilot boat PDA'
      };
    }
    return {
      onClick: 'openPilotBoatSailingPda()',
      ariaLabel: 'Edit pilot boat sailing PDA'
    };
  }
  if (desc.startsWith('MOORING/UNMOORING')) {
    if (columnIndex === 2) {
      return {
        onClick: 'openMooringPda()',
        ariaLabel: 'Edit mooring PDA'
      };
    }
    return {
      onClick: 'openMooringSailingPda()',
      ariaLabel: 'Edit mooring sailing PDA'
    };
  }
  return {
    onClick: 'focusMoneyFromEdit(this)',
    ariaLabel: columnIndex === 2 ? 'Edit PDA amount' : 'Edit Sailing PDA amount'
  };
}

function decorateMoneyEditCells() {
  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return;

  tbody.querySelectorAll('tr').forEach((row) => {
    if (row.dataset.row === 'towage' || row.dataset.row === 'bank-charges') return;
    const shouldShowEdit = shouldShowInlineMoneyEdit(row);

    [2, 3].forEach((columnIndex) => {
      const cell = row.querySelector(`td:nth-child(${columnIndex})`);
      if (!cell) return;

      if (!shouldShowEdit) {
        removeInlineMoneyEdit(cell);
        return;
      }

      let button = cell.querySelector('.row-edit');
      if (button) {
        const config = getInlineMoneyEditConfig(row, columnIndex);
        button.setAttribute('aria-label', config.ariaLabel);
        button.setAttribute('onclick', config.onClick);
        return;
      }

      const moneyField = cell.querySelector('.money-field');
      if (!moneyField) return;

      const wrapper = document.createElement('div');
      wrapper.className = 'money-edit';

      button = document.createElement('button');
      button.type = 'button';
      button.className = 'row-edit money-edit-btn';
      const config = getInlineMoneyEditConfig(row, columnIndex);
      button.setAttribute('aria-label', config.ariaLabel);
      button.setAttribute('onclick', config.onClick);

      const icon = document.createElement('img');
      icon.src = 'assets/icons/square-pen.svg';
      icon.alt = '';
      button.appendChild(icon);

      moneyField.remove();
      wrapper.appendChild(button);
      wrapper.appendChild(moneyField);
      cell.appendChild(wrapper);
    });
  });
}

function recalcOutlayTotals() {
  const tbody = document.getElementById('outlaysBody');
  if (!tbody) return;
  const roundPda = isPdaRoundingEnabled();

  let totalPda = 0;
  let totalSailing = 0;
  let runningPda = 0;
  let runningSailing = 0;
  const bankRow = tbody.querySelector('tr[data-row="bank-charges"]');
  const bankRateInput = document.getElementById('bankRate');
  const bankRate = parsePercentValue(bankRateInput?.value) / 100;

  tbody.querySelectorAll('tr').forEach((row) => {
    if (row.hidden) return;
    if (row === bankRow) {
      const bankPdaRaw = runningPda * bankRate;
      const bankPda = roundPda ? Math.ceil(bankPdaRaw) : bankPdaRaw;
      const bankSailing = runningSailing * bankRate;
      const bankPdaInput = document.getElementById('bankPda');
      const bankSailingInput = document.getElementById('bankSailing');
      if (bankPdaInput) bankPdaInput.value = formatMoneyValue(bankPda);
      if (bankSailingInput) bankSailingInput.value = formatMoneyValue(bankSailing);
      totalPda += bankPda;
      totalSailing += bankSailing;
      return;
    }

    const inputs = row.querySelectorAll('input.cell-input.money');
    const pdaInput = inputs.length >= 1 ? inputs[0] : null;
    let pdaValue = pdaInput ? getRawOrCurrentPdaValue(pdaInput) : 0;
    if (roundPda) {
      if (pdaInput && !pdaInput.readOnly && pdaInput.dataset.rawValue === undefined) {
        pdaInput.dataset.rawValue = pdaInput.value;
      }
      pdaValue = Math.ceil(pdaValue);
      if (pdaInput && !pdaInput.readOnly) {
        pdaInput.value = formatMoneyValue(pdaValue);
      }
    } else if (pdaInput && pdaInput.dataset.rawValue !== undefined && !pdaInput.readOnly) {
      pdaInput.value = pdaInput.dataset.rawValue;
      delete pdaInput.dataset.rawValue;
      pdaValue = parseMoneyValue(getMoneyInputEnteredValue(pdaInput, pdaInput.value));
    }
    const sailingValue = inputs.length >= 2
      ? parseMoneyValue(getMoneyInputEnteredValue(inputs[1], inputs[1].value))
      : 0;
    runningPda += pdaValue;
    runningSailing += sailingValue;
    totalPda += pdaValue;
    totalSailing += sailingValue;
  });

  const totalPdaInput = document.getElementById('totalPda');
  const totalSailingInput = document.getElementById('totalSailing');
  if (totalPdaInput) totalPdaInput.value = formatMoneyValue(totalPda);
  if (totalSailingInput) totalSailingInput.value = formatMoneyValue(totalSailing);
}

function removeOutlayRow(row) {
  const tbody = document.getElementById('outlaysBody');
  if (!tbody || !row) return;

  const rowInfo = getProtectedOutlayInfoForRow(row);
  if (!rowInfo) return;

  const descField = row.querySelector('td.desc textarea, td.desc input');
  const descText = String(descField?.value || '').trim();
  const label = descText ? `\n\nRow: ${descText}` : '';
  const message = `Hide this row from table and totals?${label}`;
  const confirmed = window.confirm(message);
  if (!confirmed) return;

  ensureProtectedOutlayKey(rowInfo.row);
  row.hidden = true;
  row.dataset.outlayHidden = '1';
  recalcOutlayTotals();
  saveIndexState();
}

function setSailingVisible(visible) {
  if (!outlaysTable) return;
  outlaysTable.classList.toggle('hide-sailing', !visible);
  requestAnimationFrame(() => {
    void outlaysTable.offsetWidth;
    requestAnimationFrame(refreshOutlayLayout);
  });
}

function setPdaVisible(visible) {
  if (!outlaysTable) return;
  outlaysTable.classList.toggle('hide-pda', !visible);
  requestAnimationFrame(() => {
    void outlaysTable.offsetWidth;
    requestAnimationFrame(refreshOutlayLayout);
  });
}

function setLogoNoteVisible(visible) {
  const logoNote = document.getElementById('logoLeftNote');
  if (!logoNote) return;
  logoNote.hidden = !visible;
}

function autoSizeInput(input) {
  if (!input) return;
  const value = input.value || input.placeholder || '';
  const length = Math.max(value.length, 4);
  input.style.width = `${length + 1}ch`;
}

function setDensity(mode) {
  document.body.classList.toggle('density-comfortable', mode === 'comfortable');
  document.body.classList.toggle('density-dense', mode === 'dense');
  if (densityComfortable) densityComfortable.classList.toggle('active', mode === 'comfortable');
  if (densityDense) densityDense.classList.toggle('active', mode === 'dense');
}

function getDensityMode() {
  if (document.body.classList.contains('density-dense')) return 'dense';
  if (document.body.classList.contains('density-comfortable')) return 'comfortable';
  return 'none';
}

function applyPrintDensity() {
  const current = getDensityMode();
  if (current !== 'dense') {
    printRestoreDensity = current;
    setDensity('dense');
  }
}

function shouldSuppressMobilePrintFooters() {
  return isLikelyMobileViewport();
}

function enableMobilePrintFooterSuppression() {
  if (mobilePrintPageStyle) return;
  const style = document.createElement('style');
  style.id = 'mobilePrintPageStyle';
  style.textContent = [
    '@page { margin: 0 !important; size: A4; }',
    'body.page-index { padding: 10mm 6mm 10mm 6mm !important; }'
  ].join('\n');
  document.head.appendChild(style);
  mobilePrintPageStyle = style;
}

function disableMobilePrintFooterSuppression() {
  if (!mobilePrintPageStyle) return;
  mobilePrintPageStyle.remove();
  mobilePrintPageStyle = null;
}

function printFit() {
  applyPrintDensity();
  document.body.classList.add('print-fit');
  if (document.body.classList.contains('page-index')) {
    document.body.classList.add('print-desktop-lock');
    if (shouldSuppressMobilePrintFooters()) {
      enableMobilePrintFooterSuppression();
    } else {
      disableMobilePrintFooterSuppression();
    }
    if (isLikelyMobileViewport()) {
      document.body.classList.add('print-mobile');
    } else {
      document.body.classList.remove('print-mobile');
    }
  }
  updatePrintHidden();
  setTimeout(() => {
    printWithTitle();
  }, 160);
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getExcelColumnLetter(index) {
  let result = '';
  let n = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

async function fetchImageAsDataUrl(url) {
  try {
    const response = await fetch(url);
    if (!response.ok) return '';
    const blob = await response.blob();
    return await new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result || ''));
      reader.onerror = () => resolve('');
      reader.readAsDataURL(blob);
    });
  } catch (error) {
    return '';
  }
}

async function exportToExcelWithExcelJs(options) {
  const {
    headerRows,
    detailRows,
    outlayRows,
    outlayColumns,
    totalRowValues,
    dateValue,
    currencyValue
  } = options;

  const templatePath = 'templates/PFMA DA.xlsx';
  const templateUrl = encodeURI(templatePath);
  try {
    const templateResponse = await fetch(templateUrl, { cache: 'no-store' });
    if (!templateResponse.ok) {
      throw new Error('Template not found');
    }
    const templateBuffer = await templateResponse.arrayBuffer();
    const templateWorkbook = new ExcelJS.Workbook();
    await templateWorkbook.xlsx.load(templateBuffer);
    const sheet = templateWorkbook.worksheets[0] || templateWorkbook.getWorksheet('Sheet1');
    if (sheet) {
      const getInputValue = (id) => {
        const field = document.getElementById(id);
        return field ? String(field.value || '').trim() : '';
      };

        sheet.getCell('J6').value = dateValue || '';
        sheet.getCell('C8').value = getInputValue('portInput');
        sheet.getCell('H8').value = getInputValue('berthTerminal');
        sheet.getCell('C9').value = getInputValue('agentInput');
        sheet.getCell('H9').value = currencyValue || 'EUR';
        sheet.getCell('C11').value = getInputValue('vesselNameIndex') || getVesselNameForPrint();
        sheet.getCell('C12').value = getInputValue('lengthOverall');
        sheet.getCell('C13').value = getInputValue('grossTonnage');
        sheet.getCell('H13').value = getInputValue('operationsInput');
        const cargoValueCell = sheet.getCell('H11');
        cargoValueCell.value = getInputValue('cargoInput');
        cargoValueCell.font = { ...(cargoValueCell.font || {}), bold: true };
        sheet.getCell('H12').value = getInputValue('quantityInput');

        const startRow = 16;
        const endRow = 35;

        const pdaIndex = outlayColumns.findIndex((column) => column.key === 'pda');
        outlayRows.forEach((row, index) => {
          const rowNumber = startRow + index;
          if (rowNumber > endRow) return;
          const descValue = row[0] != null ? String(row[0]) : '';
          const descCell = sheet.getCell(`A${rowNumber}`);
          descCell.value = descValue;
          descCell.font = { ...(descCell.font || {}), bold: true };
          const rawAmount = pdaIndex >= 0 ? row[pdaIndex] : '';
          const amountValue = parseMoneyValue(rawAmount);
          if (Number.isFinite(amountValue) && rawAmount !== '') {
            const amountCell = sheet.getCell(`H${rowNumber}`);
            amountCell.value = amountValue;
            amountCell.font = { ...(amountCell.font || {}), bold: true };
          } else {
            const amountCell = sheet.getCell(`H${rowNumber}`);
            amountCell.value = rawAmount || '';
            amountCell.font = { ...(amountCell.font || {}), bold: true };
          }
        });

        const totalPdaValue = pdaIndex >= 0 ? totalRowValues[pdaIndex] : '';
        const totalAmount = parseMoneyValue(totalPdaValue);
        if (Number.isFinite(totalAmount) && totalPdaValue !== '') {
          sheet.getCell('H36').value = totalAmount;
        } else {
          sheet.getCell('H36').value = totalPdaValue || '';
        }



      const existingMedia = Array.isArray(templateWorkbook.model?.media)
        ? templateWorkbook.model.media.length
        : 0;
      if (!existingMedia) {
        const logoDataUrl = await fetchImageAsDataUrl('images/logo.png');
        if (logoDataUrl) {
          const logoId = templateWorkbook.addImage({ base64: logoDataUrl, extension: 'png' });
          sheet.addImage(logoId, { tl: { col: 0, row: 0 }, ext: { width: 160, height: 60 } });
        }
      }

      const buffer = await templateWorkbook.xlsx.writeBuffer({
        useStyles: true,
        useSharedStrings: true
      });
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      const safeDate = dateValue.replace(/[^\d-]/g, '') || 'pro-forma';
      link.href = url;
      link.download = `pro-forma-da-${safeDate}.xlsx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
      return;
    }
  } catch (error) {
    window.alert('Excel template could not be loaded. Make sure templates/PFMA DA.xlsx exists.');
    return;
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'PDA';
  const dataSheet = workbook.addWorksheet('Data');
  const pfmaSheet = workbook.addWorksheet('PFMA DA');

  const headerRowMap = {};
  headerRows.forEach((row) => {
    const excelRow = dataSheet.addRow(row);
    if (row[0]) headerRowMap[String(row[0])] = excelRow.number;
  });
  dataSheet.addRow([]);
  const detailsHeader = dataSheet.addRow(['Details']);
  detailsHeader.font = { bold: true };

  const detailRowMap = {};
  detailRows.forEach((row) => {
    const excelRow = dataSheet.addRow(row);
    if (row[0]) detailRowMap[String(row[0])] = excelRow.number;
  });

  dataSheet.addRow([]);
  const outlayHeaderRow = dataSheet.addRow(outlayColumns.map((column) => column.label));
  outlayHeaderRow.font = { bold: true };
  const outlaysStartRow = outlayHeaderRow.number + 1;
  outlayRows.forEach((row) => dataSheet.addRow(row));
  const totalRow = dataSheet.addRow(totalRowValues);
  totalRow.font = { bold: true };

  dataSheet.getColumn(1).width = 50;
  dataSheet.getColumn(2).width = 22;
  dataSheet.getColumn(3).width = 22;

  pfmaSheet.getColumn(1).width = 42;
  pfmaSheet.getColumn(2).width = 28;
  pfmaSheet.getColumn(3).width = 18;

  const logoDataUrl = await fetchImageAsDataUrl('images/logo.png');
  if (logoDataUrl) {
    const logoId = workbook.addImage({ base64: logoDataUrl, extension: 'png' });
    pfmaSheet.addImage(logoId, { tl: { col: 0, row: 0 }, ext: { width: 160, height: 60 } });
  }

  pfmaSheet.mergeCells('A1:C1');
  const titleCell = pfmaSheet.getCell('A1');
  titleCell.value = 'PRO-FORMA D/A';
  titleCell.font = { bold: true, size: 16 };

  pfmaSheet.getCell('A2').value = 'Split';
  const dateCell = pfmaSheet.getCell('B2');
  dateCell.value = { formula: `Data!B${headerRowMap.Split || 3}` };

  const detailStartRow = 4;
  const detailLabels = [
    { label: 'Vessel Name', keys: ['Vessel Name', 'Vessel name'] },
    { label: 'Gross Tonnage (GT)', keys: ['Gross Tonnage (GT)', 'Gross Tonnage'] },
    { label: 'Port', keys: ['Port'] },
    { label: 'Terminal', keys: ['Terminal', 'Berth/Terminal', 'Berth/terminal'] },
    { label: 'Operation', keys: ['Operation', 'Operations'] },
    { label: 'Cargo', keys: ['Cargo'] },
    { label: 'Quantity', keys: ['Quantity'] },
    { label: 'Agent', keys: ['Agent'] }
  ];

  let rowCursor = detailStartRow;
  detailLabels.forEach((item) => {
    pfmaSheet.getCell(rowCursor, 1).value = item.label;
    const matchKey = item.keys.find((key) => detailRowMap[key]);
    if (matchKey) {
      pfmaSheet.getCell(rowCursor, 2).value = { formula: `Data!B${detailRowMap[matchKey]}` };
    }
    rowCursor += 1;
  });

  rowCursor += 1;
  const pfmaCurrency = currencyValue || 'EUR';
  pfmaSheet.getCell(rowCursor, 1).value = `Outlays & Charges Expressed In ${pfmaCurrency}`.trim();
  pfmaSheet.getCell(rowCursor, 1).font = { bold: true };
  rowCursor += 1;

  outlayColumns.forEach((column, index) => {
    const cell = pfmaSheet.getCell(rowCursor, index + 1);
    cell.value = column.label;
    cell.font = { bold: true };
  });

  const dataSheetRef = 'Data';
  outlayRows.forEach((row, index) => {
    const dataRow = outlaysStartRow + index;
    outlayColumns.forEach((column, colIndex) => {
      const dataCol = getExcelColumnLetter(colIndex);
      pfmaSheet.getCell(rowCursor + 1 + index, colIndex + 1).value = {
        formula: `${dataSheetRef}!${dataCol}${dataRow}`
      };
    });
  });

  const totalRowIndex = outlaysStartRow + outlayRows.length;
  const totalsRow = rowCursor + 1 + outlayRows.length;
  outlayColumns.forEach((column, colIndex) => {
    const dataCol = getExcelColumnLetter(colIndex);
    const cell = pfmaSheet.getCell(totalsRow, colIndex + 1);
    cell.value = { formula: `${dataSheetRef}!${dataCol}${totalRowIndex}` };
    cell.font = { bold: true };
  });

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  const safeDate = dateValue.replace(/[^\d-]/g, '') || 'pro-forma';
  link.href = url;
  link.download = `pro-forma-da-${safeDate}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

async function exportToExcel() {
  const outlaysBody = document.getElementById('outlaysBody');
  if (!outlaysBody) return;

  const dateValue = document.getElementById('dateInput')?.value || '';
  const noteValue = document.getElementById('titleNote')?.value || '';
  const currencyValue = outlaysCurrency ? outlaysCurrency.value : '';
  const includePda = outlaysTable ? !outlaysTable.classList.contains('hide-pda') : true;
  const includeSailing = outlaysTable ? !outlaysTable.classList.contains('hide-sailing') : true;

  const headerRows = [];
  headerRows.push(['Title', 'PRO-FORMA D/A']);
  if (noteValue.trim()) headerRows.push(['Note', noteValue]);
  headerRows.push(['Split', dateValue]);
  if (currencyValue) headerRows.push(['Outlays Currency', currencyValue]);

  const detailRows = [];
  document.querySelectorAll('.card-row .card .row > div').forEach((group) => {
    const label = group.querySelector('label')?.textContent?.trim() || '';
    const value = group.querySelector('input')?.value || '';
    if (label) detailRows.push([label, value]);
  });

  const outlayRows = [];
  outlaysBody.querySelectorAll('tr').forEach((row) => {
    if (row.hidden) return;
    const descInput = row.querySelector('textarea, input');
    let descValue = descInput ? descInput.value : '';
    if (row.dataset.row === 'bank-charges') {
      const rate = document.getElementById('bankRate')?.value || '';
      if (rate) descValue = `${descValue} (${rate}%)`;
    }
    const pdaInput = row.querySelector('td:nth-child(2) input');
    const sailingInput = row.querySelector('td:nth-child(3) input');
    const pdaValue = pdaInput ? getMoneyInputSourceValue(pdaInput, pdaInput.value) : '';
    const sailingValue = sailingInput ? getMoneyInputSourceValue(sailingInput, sailingInput.value) : '';
    outlayRows.push([descValue, pdaValue, sailingValue]);
  });

  const totalPda = document.getElementById('totalPda')?.value || '';
  const totalSailing = document.getElementById('totalSailing')?.value || '';

  const outlayColumns = [{ key: 'desc', label: 'Outlays' }];
  if (includePda) outlayColumns.push({ key: 'pda', label: `PDA (${currencyValue})` });
  if (includeSailing) outlayColumns.push({ key: 'sailing', label: `Sailing PDA (${currencyValue})` });

  const outlayRowsNormalized = outlayRows.map((row) => {
    const rowMap = { desc: row[0], pda: row[1], sailing: row[2] };
    return outlayColumns.map((column) => rowMap[column.key]);
  });

  const columnCount = outlayColumns.length;
  const outlayHeaderCells = outlayColumns
    .map((column) => `<th>${escapeHtml(column.label)}</th>`)
    .join('');

  const outlayBodyRows = outlayRowsNormalized
    .map((row) => {
      const cells = row
        .map((value) => `<td>${escapeHtml(value)}</td>`)
        .join('');
      return `<tr>${cells}</tr>`;
    })
    .join('');

  const totalMap = {
    desc: `Total (${currencyValue})`,
    pda: totalPda,
    sailing: totalSailing
  };
  const totalRowValues = outlayColumns.map((column) => totalMap[column.key]);
  const totalCells = totalRowValues
    .map((value) => `<td>${escapeHtml(value)}</td>`)
    .join('');

  if (window.ExcelJS && typeof ExcelJS.Workbook === 'function') {
    await exportToExcelWithExcelJs({
      headerRows,
      detailRows,
      outlayRows: outlayRowsNormalized,
      outlayColumns,
      totalRowValues,
      dateValue,
      currencyValue
    });
    return;
  }

  const headerTable = `
    <table>
      ${headerRows
        .map((row) => `<tr><th>${escapeHtml(row[0])}</th><td>${escapeHtml(row[1])}</td></tr>`)
        .join('')}
    </table>
  `;

  const detailTable = `
    <table>
      <tr><th colspan="2">Details</th></tr>
      ${detailRows
        .map((row) => `<tr><th>${escapeHtml(row[0])}</th><td>${escapeHtml(row[1])}</td></tr>`)
        .join('')}
    </table>
  `;

  const outlayTable = `
    <table>
      <tr><th colspan="${columnCount}">Outlays &amp; Charges Expressed In ${escapeHtml(currencyValue)}</th></tr>
      <tr>${outlayHeaderCells}</tr>
      ${outlayBodyRows}
      <tr>${totalCells}</tr>
    </table>
  `;

  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        table { border-collapse: collapse; margin-bottom: 12px; }
        th, td { border: 1px solid #9ca3af; padding: 4px 6px; text-align: left; }
        th { background: #e5e7eb; font-weight: 700; }
      </style>
    </head>
    <body>
      ${headerTable}
      ${detailTable}
      ${outlayTable}
    </body>
    </html>
  `;

  const blob = new Blob([html], { type: 'application/vnd.ms-excel' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  const safeDate = dateValue.replace(/[^\d-]/g, '') || 'pro-forma';
  link.href = url;
  link.download = `pro-forma-da-${safeDate}.xls`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

async function autosaveCurrentPdaBeforeNavigation() {
  if (typeof window.requestPdaAutosaveNow !== 'function') return;
  try {
    await window.requestPdaAutosaveNow({ silent: true });
  } catch (error) {
    // ignore autosave errors before navigation
  }
}

async function openTugsCalculator() {
  saveIndexState();
  const gtInputIndex = document.getElementById('grossTonnage');
  if (gtInputIndex) {
    const gtValue = gtInputIndex.value.trim();
    if (gtValue) safeStorageSet(STORAGE_KEYS.gt, gtValue);
    else safeStorageRemove(STORAGE_KEYS.gt);
  }
  await autosaveCurrentPdaBeforeNavigation();
  window.location.href = 'tugs-pda.html';
}

async function openTugsCalculatorSailing() {
  saveIndexState();
  const gtInputIndex = document.getElementById('grossTonnage');
  if (gtInputIndex) {
    const gtValue = gtInputIndex.value.trim();
    if (gtValue) safeStorageSet(STORAGE_KEYS.gt, gtValue);
    else safeStorageRemove(STORAGE_KEYS.gt);
  }
  await autosaveCurrentPdaBeforeNavigation();
  window.location.href = 'tugs-sailing-pda.html';
}

async function openLightDuesPda() {
  saveIndexState();
  const gtInputIndex = document.getElementById('grossTonnage');
  if (gtInputIndex) {
    const gtValue = gtInputIndex.value.trim();
    if (gtValue) safeStorageSet(STORAGE_KEYS.gt, gtValue);
    else safeStorageRemove(STORAGE_KEYS.gt);
  }
  await autosaveCurrentPdaBeforeNavigation();
  window.location.href = 'light-dues-pda.html';
}

async function openLightDuesSailingPda() {
  saveIndexState();
  const gtInputIndex = document.getElementById('grossTonnage');
  if (gtInputIndex) {
    const gtValue = gtInputIndex.value.trim();
    if (gtValue) safeStorageSet(STORAGE_KEYS.gt, gtValue);
    else safeStorageRemove(STORAGE_KEYS.gt);
  }
  await autosaveCurrentPdaBeforeNavigation();
  window.location.href = 'light-dues-sailing-pda.html';
}

async function openPortDuesPda() {
  saveIndexState();
  const quantityInput = document.getElementById('quantityInput');
  if (quantityInput) {
    const quantityValue = quantityInput.value.trim();
    if (quantityValue) safeStorageSet(STORAGE_KEYS.quantity, quantityValue);
    else safeStorageRemove(STORAGE_KEYS.quantity);
  }
  await autosaveCurrentPdaBeforeNavigation();
  window.location.href = 'port-dues-pda.html';
}

async function openPortDuesSailingPda() {
  saveIndexState();
  const quantityInput = document.getElementById('quantityInput');
  if (quantityInput) {
    const quantityValue = quantityInput.value.trim();
    if (quantityValue) safeStorageSet(STORAGE_KEYS.quantity, quantityValue);
    else safeStorageRemove(STORAGE_KEYS.quantity);
  }
  await autosaveCurrentPdaBeforeNavigation();
  window.location.href = 'port-dues-sailing-pda.html';
}

async function openBunkeringPda() {
  saveIndexState();
  await autosaveCurrentPdaBeforeNavigation();
  window.location.href = 'bunkering-pda.html';
}

async function openBunkeringSailingPda() {
  saveIndexState();
  await autosaveCurrentPdaBeforeNavigation();
  window.location.href = 'bunkering-sailing-pda.html';
}

async function openBerthageAnchoragePda() {
  saveIndexState();
  await autosaveCurrentPdaBeforeNavigation();
  window.location.href = 'berthage-anchorage-pda.html';
}

async function openBerthageAnchorageSailingPda() {
  saveIndexState();
  await autosaveCurrentPdaBeforeNavigation();
  window.location.href = 'berthage-anchorage-sailing-pda.html';
}

async function openPilotagePda() {
  saveIndexState();
  const gtInputIndex = document.getElementById('grossTonnage');
  if (gtInputIndex) {
    const gtValue = gtInputIndex.value.trim();
    if (gtValue) safeStorageSet(STORAGE_KEYS.gt, gtValue);
    else safeStorageRemove(STORAGE_KEYS.gt);
  }
  await autosaveCurrentPdaBeforeNavigation();
  window.location.href = 'pilot-pda.html';
}

async function openPilotageSailingPda() {
  saveIndexState();
  const gtInputIndex = document.getElementById('grossTonnage');
  if (gtInputIndex) {
    const gtValue = gtInputIndex.value.trim();
    if (gtValue) safeStorageSet(STORAGE_KEYS.gt, gtValue);
    else safeStorageRemove(STORAGE_KEYS.gt);
  }
  await autosaveCurrentPdaBeforeNavigation();
  window.location.href = 'pilot-sailing-pda.html';
}

async function openPilotBoatPda() {
  saveIndexState();
  await autosaveCurrentPdaBeforeNavigation();
  window.location.href = 'pilot-boat-pda.html';
}

async function openPilotBoatSailingPda() {
  saveIndexState();
  await autosaveCurrentPdaBeforeNavigation();
  window.location.href = 'pilot-boat-sailing-pda.html';
}

async function openMooringPda() {
  saveIndexState();
  const gtInputIndex = document.getElementById('grossTonnage');
  if (gtInputIndex) {
    const gtValue = gtInputIndex.value.trim();
    if (gtValue) safeStorageSet(STORAGE_KEYS.gt, gtValue);
    else safeStorageRemove(STORAGE_KEYS.gt);
  }
  await autosaveCurrentPdaBeforeNavigation();
  window.location.href = 'mooring-pda.html';
}

async function openMooringSailingPda() {
  saveIndexState();
  const gtInputIndex = document.getElementById('grossTonnage');
  if (gtInputIndex) {
    const gtValue = gtInputIndex.value.trim();
    if (gtValue) safeStorageSet(STORAGE_KEYS.gt, gtValue);
    else safeStorageRemove(STORAGE_KEYS.gt);
  }
  await autosaveCurrentPdaBeforeNavigation();
  window.location.href = 'mooring-sailing-pda.html';
}

async function goHome() {
  savePilotBoatAmount();
  saveTugsState();
  window.location.href = 'index.html';
}

function updatePrintHidden() {
  document.querySelectorAll('[data-print-hide-empty]').forEach((input) => {
    const empty = !(input.value || '').trim();
    input.classList.toggle('print-hide', empty);
    const group = input.closest('[data-print-hide-group]');
    if (group) group.classList.toggle('print-hide', empty);
  });
}

function clearPrintHidden() {
  document.querySelectorAll('.print-hide').forEach((el) => {
    el.classList.remove('print-hide');
  });
}

function getDragAfterElement(container, y) {
  const rows = [...container.querySelectorAll('tr:not(.dragging)')];
  let closest = { offset: Number.NEGATIVE_INFINITY, element: null };
  rows.forEach((row) => {
    const box = row.getBoundingClientRect();
    const offset = y - box.top - box.height / 2;
    if (offset < 0 && offset > closest.offset) {
      closest = { offset, element: row };
    }
  });
  return closest.element;
}

function initIndex() {
  const outlaysBody = document.getElementById('outlaysBody');
  if (!outlaysBody) return;

  if (isNewDraftMode()) {
    clearCalculatorStorageForNewDraft();
    safeStorageRemove(STORAGE_KEYS.indexState);
  }

  if (!defaultOutlaysHtml) {
    defaultOutlaysHtml = outlaysBody.innerHTML;
  }

  document.body.classList.add('density-comfortable');
  restoreIndexState();
  decorateOutlayRows();
  recalcOutlayTotals();
  refreshOutlayLayout();

  outlaysBody.addEventListener('click', (event) => {
    const button = event.target.closest('.row-remove');
    if (!button) return;
    const row = button.closest('tr');
    removeOutlayRow(row);
  });

  outlaysBody.addEventListener('dragstart', (event) => {
    const target = event.target instanceof Element ? event.target : null;
    if (!target) return;
    if (target.matches('.row-handle img')) {
      event.preventDefault();
      return;
    }
    const handle = target.closest('.row-handle');
    if (!handle) return;
    const row = handle.closest('tr');
    if (!row) return;
    draggingRow = row;
    row.classList.add('dragging');
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', '');
  });

  outlaysBody.addEventListener('dragend', () => {
    if (draggingRow) draggingRow.classList.remove('dragging');
    draggingRow = null;
    saveIndexState();
  });

  outlaysBody.addEventListener('dragover', (event) => {
    if (!draggingRow) return;
    event.preventDefault();
    const afterElement = getDragAfterElement(outlaysBody, event.clientY);
    const bankRow = outlaysBody.querySelector('tr[data-row="bank-charges"]');
    if (!afterElement) {
      if (bankRow) {
        outlaysBody.insertBefore(draggingRow, bankRow);
      } else {
        outlaysBody.appendChild(draggingRow);
      }
    } else if (afterElement !== draggingRow) {
      if (bankRow && afterElement === bankRow) {
        outlaysBody.insertBefore(draggingRow, bankRow);
      } else {
        outlaysBody.insertBefore(draggingRow, afterElement);
      }
    }
  });

  outlaysBody.addEventListener('input', (event) => {
    const target = event.target;
    if (!target) return;
    if (target.tagName === 'TEXTAREA') {
      autoResizeTextarea(target);
      if (target.classList.contains('text')) {
        decorateMoneyEditCells();
      }
    }
    if (target.classList.contains('money') && isPdaRoundingEnabled() && isPdaMoneyInput(target) && !target.readOnly) {
      target.dataset.rawValue = target.value;
    }
    if (target.classList.contains('money') || target.classList.contains('percent-input')) {
      recalcOutlayTotals();
    }
    saveIndexState();
  });

  outlaysTable = document.querySelector('.outlays-table');
  const togglePda = document.getElementById('togglePda');
  toggleSailing = document.getElementById('toggleSailing');
  const toggleLogoNote = document.getElementById('toggleLogoNote');
  outlaysCurrency = document.getElementById('outlaysCurrency');
  roundPdaPrices = document.getElementById('roundPdaPrices');
  const globalImoTransport = document.getElementById('globalImoTransport');
  const optionsCard = document.getElementById('optionsCard');
  const optionsBody = document.getElementById('optionsBody');
  const optionsToggleBtn = document.getElementById('optionsToggleBtn');
  const optionsToggleLabel = optionsToggleBtn ? optionsToggleBtn.querySelector('.options-toggle-label') : null;
  const setOptionsExpanded = (expanded) => {
    if (!optionsCard || !optionsBody || !optionsToggleBtn) return;
    optionsCard.classList.toggle('is-expanded', Boolean(expanded));
    optionsCard.classList.toggle('is-collapsed', !expanded);
    optionsBody.hidden = !expanded;
    optionsToggleBtn.setAttribute('aria-expanded', expanded ? 'true' : 'false');
    if (optionsToggleLabel) optionsToggleLabel.textContent = expanded ? 'Less' : 'More';
  };
  if (optionsCard && optionsBody && optionsToggleBtn) {
    setOptionsExpanded(false);
    optionsToggleBtn.addEventListener('click', () => {
      const isExpanded = optionsToggleBtn.getAttribute('aria-expanded') === 'true';
      setOptionsExpanded(!isExpanded);
    });
  }
  if (togglePda) {
    togglePda.addEventListener('change', () => {
      setPdaVisible(togglePda.checked);
      saveIndexState();
    });
    setPdaVisible(togglePda.checked);
  }
  if (toggleSailing) {
    toggleSailing.addEventListener('change', () => {
      setSailingVisible(toggleSailing.checked);
      saveIndexState();
    });
    setSailingVisible(toggleSailing.checked);
  }
  if (toggleLogoNote) {
    setLogoNoteVisible(toggleLogoNote.checked);
    toggleLogoNote.addEventListener('change', () => {
      setLogoNoteVisible(toggleLogoNote.checked);
      saveIndexState();
    });
  }
  if (roundPdaPrices) {
    roundPdaPrices.addEventListener('change', () => {
      recalcOutlayTotals();
      saveIndexState();
    });
  }
  if (globalImoTransport) {
    const globalImoState = getGlobalImoTransportState();
    if (globalImoState !== null && globalImoTransport.checked !== globalImoState) {
      globalImoTransport.checked = globalImoState;
    }
    if (shouldApplyGlobalImoStateOnInit(globalImoState, globalImoTransport.checked)) {
      setGlobalImoTransportState(true);
    }

    globalImoTransport.addEventListener('change', () => {
      setGlobalImoTransportState(Boolean(globalImoTransport.checked));
      updateTowageFromStorage();
      updateLightDuesFromStorage();
      updatePilotageFromStorage();
      updatePortDuesFromStorage();
      updateImoToggleLabelColor(globalImoTransport);
      saveIndexState();
    });
    updateImoToggleLabelColor(globalImoTransport);
  }
  wrapMoneyFields();
  decorateMoneyEditCells();
  recalcOutlayTotals();
  if (outlaysCurrency) {
    updateCurrencySymbol(outlaysCurrency.value);
    outlaysCurrency.addEventListener('change', () => {
      updateCurrencySymbol(outlaysCurrency.value);
      saveIndexState();
    });
  }

  const vesselNameIndex = document.getElementById('vesselNameIndex');
  if (vesselNameIndex) {
    updateVesselNameFromStorage(vesselNameIndex);
    vesselNameIndex.addEventListener('input', () => {
      const value = vesselNameIndex.value.trim();
      if (value) safeStorageSet(STORAGE_KEYS.vesselName, value);
      else safeStorageRemove(STORAGE_KEYS.vesselName);
      saveIndexState();
    });
  }

  const gtInputIndex = document.getElementById('grossTonnage');
  if (gtInputIndex) {
    updateIndexGtFromStorage();
    gtInputIndex.addEventListener('input', () => {
      const value = gtInputIndex.value.trim();
      if (value) safeStorageSet(STORAGE_KEYS.gt, value);
      else safeStorageRemove(STORAGE_KEYS.gt);
      updateTowageFromStorage();
      updateLightDuesFromStorage();
      updatePilotageFromStorage();
      updateMooringFromStorage();
      saveIndexState();
    });
  }

  const lengthOverallInput = document.getElementById('lengthOverall');
  if (lengthOverallInput) {
    lengthOverallInput.addEventListener('input', () => {
      updateBerthageFromStorage();
      saveIndexState();
    });
  }

  const quantityInputIndex = document.getElementById('quantityInput');
  if (quantityInputIndex) {
    const storedQuantity = safeStorageGet(STORAGE_KEYS.quantity);
    if (storedQuantity === null) {
      const currentQuantity = quantityInputIndex.value.trim();
      if (currentQuantity) safeStorageSet(STORAGE_KEYS.quantity, currentQuantity);
    }
    updateIndexQuantityFromStorage();
    quantityInputIndex.addEventListener('input', () => {
      const value = quantityInputIndex.value.trim();
      if (value) safeStorageSet(STORAGE_KEYS.quantity, value);
      else safeStorageRemove(STORAGE_KEYS.quantity);
      updatePortDuesFromStorage();
      saveIndexState();
    });
  }

  const towageDesc = document.querySelector('tr[data-row="towage"] textarea.cell-input.text');
  if (towageDesc) {
    towageDesc.dataset.baseText = 'TOWAGE';
    towageDesc.addEventListener('input', () => {
      towageDesc.dataset.baseText = 'TOWAGE';
    });
  }

  const titleNote = document.getElementById('titleNote');
  if (titleNote) {
    autoSizeInput(titleNote);
    titleNote.addEventListener('input', () => {
      autoSizeInput(titleNote);
      saveIndexState();
    });
  }

  const logoLeftNote = document.getElementById('logoLeftNote');
  if (logoLeftNote) {
    autoResizeTextarea(logoLeftNote);
    logoLeftNote.addEventListener('input', () => {
      autoResizeTextarea(logoLeftNote);
      saveIndexState();
    });
  }

  const dateInput = document.getElementById('dateInput');
  if (dateInput && !dateInput.value) {
    const now = new Date();
    const day = String(now.getDate());
    const month = String(now.getMonth() + 1);
    const year = String(now.getFullYear()).slice(-2);
    dateInput.value = `${day}/${month}/${year}`;
  }
  if (dateInput) {
    dateInput.addEventListener('input', saveIndexState);
  }

  densityComfortable = document.getElementById('densityComfortable');
  densityDense = document.getElementById('densityDense');
  if (densityComfortable) {
    densityComfortable.addEventListener('click', () => {
      setDensity('comfortable');
      saveIndexState();
    });
  }
  if (densityDense) {
    densityDense.addEventListener('click', () => {
      setDensity('dense');
      saveIndexState();
    });
  }

  updateTowageFromStorage();
  updateLightDuesFromStorage();
  updatePortDuesFromStorage();
  updateBunkeringFromStorage();
  updateBerthageFromStorage();
  updatePilotageFromStorage();
  updatePilotBoatFromStorage();
  updateMooringFromStorage();
  updateIndexGtFromStorage();
  updateIndexQuantityFromStorage();
  window.addEventListener('pageshow', () => {
    updateTowageFromStorage();
    updateLightDuesFromStorage();
    updatePortDuesFromStorage();
    updateBunkeringFromStorage();
    updateBerthageFromStorage();
    updatePilotageFromStorage();
    updatePilotBoatFromStorage();
    updateMooringFromStorage();
    updateIndexQuantityFromStorage();
    if (globalImoTransport) {
      const globalImoState = getGlobalImoTransportState();
      if (globalImoState !== null && globalImoTransport.checked !== globalImoState) {
        globalImoTransport.checked = globalImoState;
      }
      updateImoToggleLabelColor(globalImoTransport);
    }
  });
  window.addEventListener('storage', () => {
    updateTowageFromStorage();
    updateLightDuesFromStorage();
    updatePortDuesFromStorage();
    updateBunkeringFromStorage();
    updateBerthageFromStorage();
    updatePilotageFromStorage();
    updatePilotBoatFromStorage();
    updateMooringFromStorage();
    updateIndexGtFromStorage();
    updateIndexQuantityFromStorage();
    updateVesselNameFromStorage(vesselNameIndex);
    if (globalImoTransport) {
      const globalImoState = getGlobalImoTransportState();
      if (globalImoState !== null && globalImoTransport.checked !== globalImoState) {
        globalImoTransport.checked = globalImoState;
      }
      updateImoToggleLabelColor(globalImoTransport);
    }
  });

  window.addEventListener('beforeunload', saveIndexState);
  window.addEventListener('pagehide', saveIndexState);

  window.addEventListener('afterprint', clearPrintHidden);
}

function isLightDuesPdaPage() {
  return Boolean(document.body && document.body.classList.contains('page-light-dues-pda'));
}

function isLightDuesSailingPage() {
  return Boolean(document.body && document.body.classList.contains('page-light-dues-sailing'));
}

function getLightDuesStateKey() {
  if (isLightDuesSailingPage()) return STORAGE_KEYS.lightDuesStateSailing;
  if (isLightDuesPdaPage()) return STORAGE_KEYS.lightDuesState;
  return null;
}

function getTierBandFromGt(type, gt) {
  const tierOptions = LIGHT_DUES_TIER_OPTIONS[type];
  if (!Array.isArray(tierOptions) || tierOptions.length === 0) return '';
  if (!Number.isFinite(gt) || gt <= 0) return tierOptions[0].value;

  const matching = tierOptions.find((option) => {
    const minOk = option.min === undefined || gt >= option.min;
    const maxOk = option.max === undefined || gt <= option.max;
    return minOk && maxOk;
  });
  return matching ? matching.value : tierOptions[tierOptions.length - 1].value;
}

function rebuildTierOptions(typeInput, tierBandInput, preferredBand, gt) {
  if (!typeInput || !tierBandInput) return;
  const options = LIGHT_DUES_TIER_OPTIONS[typeInput.value] || [];
  tierBandInput.innerHTML = '';
  options.forEach((option) => {
    const node = document.createElement('option');
    node.value = option.value;
    node.textContent = option.label;
    tierBandInput.appendChild(node);
  });
  if (options.length === 0) return;

  const hasPreferred = preferredBand && options.some((option) => option.value === preferredBand);
  tierBandInput.value = hasPreferred ? preferredBand : getTierBandFromGt(typeInput.value, gt);
}

function getLightDuesTariff(type, tierBand, gt) {
  const tierOptions = LIGHT_DUES_TIER_OPTIONS[type];
  if (Array.isArray(tierOptions) && tierOptions.length > 0) {
    const resolvedBand = tierBand || getTierBandFromGt(type, gt);
    return LIGHT_DUES_TARIFFS[type][resolvedBand] || LIGHT_DUES_TARIFFS[type][tierOptions[0].value];
  }
  return LIGHT_DUES_TARIFFS[type] || LIGHT_DUES_TARIFFS[LIGHT_DUES_DEFAULT_TYPE];
}

function setTierVisibility(typeInput, tierWrap) {
  if (!typeInput || !tierWrap) return;
  const hasTierOptions = Boolean(LIGHT_DUES_TIER_OPTIONS[typeInput.value]);
  tierWrap.style.display = hasTierOptions ? '' : 'none';
}

function enforceLightDuesTypeFromImo(typeInput, tierBandInput, gtInput, options = {}) {
  if (!typeInput || !isLightDuesPdaPage() && !isLightDuesSailingPage()) return false;
  const globalImo = getGlobalImoTransportState();
  if (globalImo) {
    const targetType = 'tanker';
    if (typeInput.value === targetType) return false;
    typeInput.value = targetType;
    rebuildTierOptions(typeInput, tierBandInput, '', Number(gtInput?.value));
    setTierVisibility(typeInput, document.getElementById('lightDuesTierWrap'));
    return true;
  }

  if (options.resetForced && typeInput.value === 'tanker') {
    typeInput.value = LIGHT_DUES_DEFAULT_TYPE;
    rebuildTierOptions(typeInput, tierBandInput, '', Number(gtInput?.value));
    setTierVisibility(typeInput, document.getElementById('lightDuesTierWrap'));
    return true;
  }

  return false;
}

function saveLightDuesState() {
  const stateKey = getLightDuesStateKey();
  if (!stateKey) return;
  const typeInput = document.getElementById('lightDuesType');
  const tierBandInput = document.getElementById('lightDuesTierBand');
  const gtInput = document.getElementById('lightDuesGt');
  const period30Input = document.getElementById('lightDuesPeriod30');
  const period12Input = document.getElementById('lightDuesPeriod12');
  const validInput = document.getElementById('lightDuesValid');
  if (!typeInput || !tierBandInput || !gtInput || !period30Input || !period12Input) return;

  const state = {
    type: typeInput.value,
    tierBand: tierBandInput.value,
    gt: gtInput.value,
    period30: Boolean(period30Input.checked),
    period12: Boolean(period12Input.checked)
  };

  if (validInput && (isLightDuesPdaPage() || isLightDuesSailingPage())) {
    state.validLightDues = Boolean(validInput.checked);
  }

  safeStorageSet(stateKey, JSON.stringify(state));
}

function restoreLightDuesState(typeInput, tierBandInput, gtInput, period30Input, period12Input) {
  const stateKey = getLightDuesStateKey();
  const validInput = document.getElementById('lightDuesValid');
  let preferredTierBand = '';

  if (stateKey) {
    const raw = safeStorageGet(stateKey);
    if (raw) {
      let state = null;
      try {
        state = JSON.parse(raw);
      } catch (error) {
        state = null;
      }
      if (state && typeof state === 'object') {
        if (typeof state.type === 'string' && state.type) typeInput.value = state.type;
        if (typeof state.gt === 'string') gtInput.value = state.gt;
        preferredTierBand = state.tierBand || state.bulkBand || '';
        if (typeof state.period30 === 'boolean') period30Input.checked = state.period30;
        if (typeof state.period12 === 'boolean') period12Input.checked = state.period12;
        if (validInput && typeof state.validLightDues === 'boolean') validInput.checked = state.validLightDues;
      }
    }
  }

  const sharedGt = safeStorageGet(STORAGE_KEYS.gt);
  if (sharedGt !== null) {
    gtInput.value = sharedGt;
  }

  if (!period30Input.checked && !period12Input.checked) {
    period30Input.checked = true;
  }
  if (period30Input.checked && period12Input.checked) {
    period12Input.checked = false;
  }
  rebuildTierOptions(typeInput, tierBandInput, preferredTierBand, Number(gtInput.value));
}

function enforceSingleLightDuesPeriod(changedInput, otherInput) {
  if (!changedInput || !otherInput) return;
  if (changedInput.checked) {
    otherInput.checked = false;
    return;
  }
  if (!otherInput.checked) {
    changedInput.checked = true;
  }
}

function calculateLightDues() {
  const typeInput = document.getElementById('lightDuesType');
  const tierBandInput = document.getElementById('lightDuesTierBand');
  const gtInput = document.getElementById('lightDuesGt');
  const period30Input = document.getElementById('lightDuesPeriod30');
  const period12Input = document.getElementById('lightDuesPeriod12');
  const validInput = document.getElementById('lightDuesValid');
  const tariff30 = document.getElementById('lightDuesTariff30');
  const tariff12 = document.getElementById('lightDuesTariff12');
  const amount30 = document.getElementById('lightDuesAmount30');
  const amount12 = document.getElementById('lightDuesAmount12');

  if (
    !typeInput || !tierBandInput || !gtInput || !period30Input || !period12Input ||
    !tariff30 || !tariff12 || !amount30 || !amount12
  ) return;

  if (!period30Input.checked && !period12Input.checked) {
    period30Input.checked = true;
  }
  if (period30Input.checked && period12Input.checked) {
    period12Input.checked = false;
  }

  const gt = Number(gtInput.value);
  const tariff = getLightDuesTariff(typeInput.value, tierBandInput.value, gt);
  tariff30.value = tariff.label30;
  tariff12.value = tariff.label12;

  const selectedPeriod = period12Input.checked ? '12' : '30';
  const selectedTariffLabel = selectedPeriod === '12' ? tariff.label12 : tariff.label30;
  const amountKey = isLightDuesSailingPage() ? STORAGE_KEYS.lightDuesAmountSailing : STORAGE_KEYS.lightDuesAmountPda;
  const tariffKey = isLightDuesSailingPage() ? STORAGE_KEYS.lightDuesTariffSailing : STORAGE_KEYS.lightDuesTariffPda;

  safeStorageSet(tariffKey, selectedTariffLabel);

  const validLightDues = Boolean(validInput && validInput.checked && (isLightDuesPdaPage() || isLightDuesSailingPage()));
  if (validLightDues) {
    const zeroFormatted = formatMoneyValue(0);
    amount30.value = zeroFormatted;
    amount12.value = zeroFormatted;
    safeStorageSet(amountKey, '0.00');
    saveLightDuesState();
    return;
  }

  if (!Number.isFinite(gt) || gt <= 0) {
    amount30.value = '';
    amount12.value = '';
    safeStorageRemove(amountKey);
    saveLightDuesState();
    return;
  }

  const calculated30 = gt * tariff.rate30;
  const calculated12 = gt * tariff.rate12;
  amount30.value = formatMoneyValue(calculated30);
  amount12.value = formatMoneyValue(calculated12);

  const selectedAmount = selectedPeriod === '12' ? calculated12 : calculated30;
  safeStorageSet(amountKey, selectedAmount.toFixed(2));
  saveLightDuesState();
}

function initLightDues() {
  const typeInput = document.getElementById('lightDuesType');
  const tierBandInput = document.getElementById('lightDuesTierBand');
  const tierWrap = document.getElementById('lightDuesTierWrap');
  const gtInput = document.getElementById('lightDuesGt');
  const period30Input = document.getElementById('lightDuesPeriod30');
  const period12Input = document.getElementById('lightDuesPeriod12');
  const validInput = document.getElementById('lightDuesValid');
  if (!typeInput || !tierBandInput || !tierWrap || !gtInput || !period30Input || !period12Input) return;

  restoreLightDuesState(typeInput, tierBandInput, gtInput, period30Input, period12Input);
  enforceLightDuesTypeFromImo(typeInput, tierBandInput, gtInput);
  setTierVisibility(typeInput, tierWrap);
  calculateLightDues();

  typeInput.addEventListener('change', () => {
    rebuildTierOptions(typeInput, tierBandInput, '', Number(gtInput.value));
    setTierVisibility(typeInput, tierWrap);
    calculateLightDues();
  });

  tierBandInput.addEventListener('change', calculateLightDues);

  period30Input.addEventListener('change', () => {
    enforceSingleLightDuesPeriod(period30Input, period12Input);
    calculateLightDues();
  });

  period12Input.addEventListener('change', () => {
    enforceSingleLightDuesPeriod(period12Input, period30Input);
    calculateLightDues();
  });

  if (validInput) {
    validInput.addEventListener('change', calculateLightDues);
  }

  gtInput.addEventListener('input', () => {
    const raw = gtInput.value.trim();
    if (raw) safeStorageSet(STORAGE_KEYS.gt, raw);
    else safeStorageRemove(STORAGE_KEYS.gt);

    if (LIGHT_DUES_TIER_OPTIONS[typeInput.value]) {
      rebuildTierOptions(typeInput, tierBandInput, '', Number(gtInput.value));
    }
    calculateLightDues();
  });

  const syncGtFromShared = () => {
    const sharedGt = safeStorageGet(STORAGE_KEYS.gt);
    const nextGt = sharedGt === null ? '' : sharedGt;
    if (gtInput.value !== nextGt) {
      gtInput.value = nextGt;
      if (LIGHT_DUES_TIER_OPTIONS[typeInput.value]) {
        rebuildTierOptions(typeInput, tierBandInput, '', Number(gtInput.value));
      }
    }
    calculateLightDues();
  };

  window.addEventListener('storage', (event) => {
    if (event.key === STORAGE_KEYS.gt) {
      syncGtFromShared();
    }
    if (event.key === STORAGE_KEYS.globalImoTransport) {
      if (enforceLightDuesTypeFromImo(typeInput, tierBandInput, gtInput, { resetForced: true })) {
        calculateLightDues();
      }
    }
  });

  window.addEventListener('pageshow', () => {
    syncGtFromShared();
  });
}

function isPortDuesPdaPage() {
  return Boolean(document.body && document.body.classList.contains('page-port-dues-pda'));
}

function isPortDuesSailingPage() {
  return Boolean(document.body && document.body.classList.contains('page-port-dues-sailing'));
}

function getPortDuesStateKey() {
  if (isPortDuesSailingPage()) return STORAGE_KEYS.portDuesStateSailing;
  if (isPortDuesPdaPage()) return STORAGE_KEYS.portDuesState;
  return null;
}

function getPortDuesAmountKey() {
  if (isPortDuesSailingPage()) return STORAGE_KEYS.portDuesAmountSailing;
  if (isPortDuesPdaPage()) return STORAGE_KEYS.portDuesAmountPda;
  return null;
}

function getPortDuesCargoAmountKey() {
  if (isPortDuesSailingPage()) return STORAGE_KEYS.portDuesCargoAmountSailing;
  if (isPortDuesPdaPage()) return STORAGE_KEYS.portDuesCargoAmountPda;
  return null;
}

function getPortDuesCargoTypeConfig(type) {
  return PORT_DUES_CARGO_TYPES[type] || PORT_DUES_CARGO_TYPES[PORT_DUES_DEFAULT_CARGO_TYPE];
}

function enforcePortDuesCargoTypeFromImo(cargoTypeInput, options = {}) {
  if (!cargoTypeInput || (!isPortDuesPdaPage() && !isPortDuesSailingPage())) return false;
  const globalImo = getGlobalImoTransportState();
  const stateKey = getPortDuesStateKey();
  const state = stateKey ? parseStoredJson(safeStorageGet(stateKey)) : null;
  const lastManualCargoType = (
    state
    && typeof state.lastManualCargoType === 'string'
    && PORT_DUES_CARGO_TYPES[state.lastManualCargoType]
  )
    ? state.lastManualCargoType
    : PORT_DUES_DEFAULT_CARGO_TYPE;
  const wasForcedByImo = Boolean(state && state.cargoTypeForcedByImo);
  if (globalImo) {
    const targetType = 'liquidCargo';
    if (cargoTypeInput.value === targetType) return false;
    cargoTypeInput.value = targetType;
    return true;
  }

  const nextType = resolvePortDuesCargoTypeForImo(cargoTypeInput.value, false, {
    resetForced: Boolean(options.resetForced),
    wasForcedByImo,
    lastManualCargoType
  });
  if (options.resetForced && cargoTypeInput.value !== nextType) {
    cargoTypeInput.value = nextType;
    return true;
  }

  return false;
}

function setPortDuesTerminalDischargedState(toggleInput, wrapInput, quantityInput) {
  if (!toggleInput || !wrapInput || !quantityInput) return;
  const enabled = Boolean(toggleInput.checked);
  wrapInput.style.display = enabled ? '' : 'none';
  quantityInput.disabled = !enabled;
}

function savePortDuesState() {
  const stateKey = getPortDuesStateKey();
  if (!stateKey) return;

  const cargoQuantityInput = document.getElementById('portDuesCargoQuantity');
  const cargoTypeInput = document.getElementById('portDuesCargoType');
  if (!cargoQuantityInput || !cargoTypeInput) return;
  const terminalDischargedToggleInput = document.getElementById('portDuesTerminalDischargedToggle');
  const terminalDischargedQuantityInput = document.getElementById('portDuesTerminalDischargedQuantity');

  const existingState = parseStoredJson(safeStorageGet(stateKey));
  const currentCargoType = PORT_DUES_CARGO_TYPES[cargoTypeInput.value]
    ? cargoTypeInput.value
    : PORT_DUES_DEFAULT_CARGO_TYPE;
  const state = {
    ...(existingState && typeof existingState === 'object' ? existingState : {}),
    cargoQuantity: cargoQuantityInput.value,
    cargoType: currentCargoType
  };
  const globalImo = getGlobalImoTransportState();
  if (globalImo && currentCargoType === 'liquidCargo') {
    state.cargoTypeForcedByImo = true;
    if (
      !state.lastManualCargoType
      || !PORT_DUES_CARGO_TYPES[state.lastManualCargoType]
      || state.lastManualCargoType === 'liquidCargo'
    ) {
      state.lastManualCargoType = (
        existingState
        && typeof existingState.cargoType === 'string'
        && PORT_DUES_CARGO_TYPES[existingState.cargoType]
        && existingState.cargoType !== 'liquidCargo'
      )
        ? existingState.cargoType
        : PORT_DUES_DEFAULT_CARGO_TYPE;
    }
  } else {
    state.cargoTypeForcedByImo = false;
    state.lastManualCargoType = currentCargoType;
  }
  if (terminalDischargedToggleInput && terminalDischargedQuantityInput) {
    state.terminalDischargedEnabled = Boolean(terminalDischargedToggleInput.checked);
    state.terminalDischargedQuantity = terminalDischargedQuantityInput.value;
  }
  safeStorageSet(stateKey, JSON.stringify(state));
}

function restorePortDuesState(
  cargoQuantityInput,
  cargoTypeInput,
  terminalDischargedToggleInput,
  terminalDischargedWrapInput,
  terminalDischargedQuantityInput
) {
  const stateKey = getPortDuesStateKey();
  if (stateKey) {
    const raw = safeStorageGet(stateKey);
    if (raw) {
      let state = null;
      try {
        state = JSON.parse(raw);
      } catch (error) {
        state = null;
      }
      if (state && typeof state === 'object') {
        if (typeof state.cargoQuantity === 'string') cargoQuantityInput.value = state.cargoQuantity;
        if (typeof state.cargoType === 'string' && PORT_DUES_CARGO_TYPES[state.cargoType]) {
          cargoTypeInput.value = state.cargoType;
        }
        if (terminalDischargedToggleInput && typeof state.terminalDischargedEnabled === 'boolean') {
          terminalDischargedToggleInput.checked = state.terminalDischargedEnabled;
        }
        if (terminalDischargedQuantityInput && typeof state.terminalDischargedQuantity === 'string') {
          terminalDischargedQuantityInput.value = state.terminalDischargedQuantity;
        }
      }
    }
  }

  const sharedQuantity = safeStorageGet(STORAGE_KEYS.quantity);
  if (sharedQuantity !== null) {
    cargoQuantityInput.value = sharedQuantity;
  }

  if (!PORT_DUES_CARGO_TYPES[cargoTypeInput.value]) {
    cargoTypeInput.value = 'bulkCargo';
  }
  setPortDuesTerminalDischargedState(
    terminalDischargedToggleInput,
    terminalDischargedWrapInput,
    terminalDischargedQuantityInput
  );
}

function calculatePortDues() {
  const cargoQuantityInput = document.getElementById('portDuesCargoQuantity');
  const cargoTypeInput = document.getElementById('portDuesCargoType');
  const cargoTariffInput = document.getElementById('portDuesCargoTariff');
  const cargoAmountInput = document.getElementById('portDuesCargoAmount');
  const totalAmountInput = document.getElementById('portDuesTotalAmount');
  const terminalDischargedToggleInput = document.getElementById('portDuesTerminalDischargedToggle');
  const terminalDischargedQuantityInput = document.getElementById('portDuesTerminalDischargedQuantity');
  if (
    !cargoQuantityInput || !cargoTypeInput || !cargoTariffInput || !cargoAmountInput || !totalAmountInput
  ) return;

  const cargoType = getPortDuesCargoTypeConfig(cargoTypeInput.value);
  const cargoQuantity = parseMoneyValue(cargoQuantityInput.value);
  const terminalDischargedQuantity = terminalDischargedQuantityInput
    ? parseMoneyValue(terminalDischargedQuantityInput.value)
    : 0;
  const useTerminalDischargedQuantity = Boolean(
    isPortDuesSailingPage() && terminalDischargedToggleInput && terminalDischargedToggleInput.checked
  );
  const effectiveCargoQuantity = useTerminalDischargedQuantity ? terminalDischargedQuantity : cargoQuantity;
  const cargoAmount = effectiveCargoQuantity > 0 ? effectiveCargoQuantity * cargoType.rate : 0;

  const totalAmount = cargoAmount;
  const amountKey = getPortDuesAmountKey();
  const cargoAmountKey = getPortDuesCargoAmountKey();
  if (amountKey) {
    safeStorageSet(amountKey, totalAmount.toFixed(2));
  }
  if (cargoAmountKey) {
    safeStorageSet(cargoAmountKey, cargoAmount.toFixed(2));
  }

  const quantityRaw = cargoQuantityInput.value.trim();
  if (quantityRaw) safeStorageSet(STORAGE_KEYS.quantity, quantityRaw);
  else safeStorageRemove(STORAGE_KEYS.quantity);

  cargoTariffInput.value = cargoType.label;
  cargoAmountInput.value = formatMoneyValue(cargoAmount);
  totalAmountInput.value = formatMoneyValue(totalAmount);
  savePortDuesState();
}

function initPortDues() {
  const cargoQuantityInput = document.getElementById('portDuesCargoQuantity');
  const cargoTypeInput = document.getElementById('portDuesCargoType');
  const terminalDischargedToggleInput = document.getElementById('portDuesTerminalDischargedToggle');
  const terminalDischargedWrapInput = document.getElementById('portDuesTerminalDischargedWrap');
  const terminalDischargedQuantityInput = document.getElementById('portDuesTerminalDischargedQuantity');
  if (!cargoQuantityInput || !cargoTypeInput) return;

  restorePortDuesState(
    cargoQuantityInput,
    cargoTypeInput,
    terminalDischargedToggleInput,
    terminalDischargedWrapInput,
    terminalDischargedQuantityInput
  );
  enforcePortDuesCargoTypeFromImo(cargoTypeInput, { resetForced: true });
  calculatePortDues();

  cargoQuantityInput.addEventListener('input', calculatePortDues);
  cargoTypeInput.addEventListener('change', calculatePortDues);
  if (terminalDischargedToggleInput && terminalDischargedWrapInput && terminalDischargedQuantityInput) {
    terminalDischargedToggleInput.addEventListener('change', () => {
      setPortDuesTerminalDischargedState(
        terminalDischargedToggleInput,
        terminalDischargedWrapInput,
        terminalDischargedQuantityInput
      );
      calculatePortDues();
    });
    terminalDischargedQuantityInput.addEventListener('input', calculatePortDues);
  }

  window.addEventListener('storage', (event) => {
    if (event.key === STORAGE_KEYS.quantity) {
      const nextQuantity = safeStorageGet(STORAGE_KEYS.quantity) || '';
      if (cargoQuantityInput.value !== nextQuantity) {
        cargoQuantityInput.value = nextQuantity;
        calculatePortDues();
      }
    }
    if (event.key === STORAGE_KEYS.globalImoTransport) {
      if (enforcePortDuesCargoTypeFromImo(cargoTypeInput, { resetForced: true })) {
        calculatePortDues();
      }
    }
  });
}

function isBunkeringPdaPage() {
  return Boolean(document.body && document.body.classList.contains('page-bunkering-pda'));
}

function isBunkeringSailingPage() {
  return Boolean(document.body && document.body.classList.contains('page-bunkering-sailing'));
}

function getBunkeringStateKey() {
  if (isBunkeringSailingPage()) return STORAGE_KEYS.bunkeringStateSailing;
  if (isBunkeringPdaPage()) return STORAGE_KEYS.bunkeringState;
  return null;
}

function getBunkeringAmountKey() {
  if (isBunkeringSailingPage()) return STORAGE_KEYS.portDuesBunkeringAmountSailing;
  if (isBunkeringPdaPage()) return STORAGE_KEYS.portDuesBunkeringAmountPda;
  return null;
}

function saveBunkeringState() {
  const stateKey = getBunkeringStateKey();
  if (!stateKey) return;
  const quantityInput = document.getElementById('bunkeringQuantity');
  if (!quantityInput) return;
  safeStorageSet(stateKey, JSON.stringify({
    quantity: quantityInput.value
  }));
}

function restoreBunkeringState(quantityInput) {
  if (!quantityInput) return;
  const stateKey = getBunkeringStateKey();
  if (!stateKey) return;
  const raw = safeStorageGet(stateKey);
  if (!raw) return;
  try {
    const state = JSON.parse(raw);
    if (state && typeof state === 'object' && typeof state.quantity === 'string') {
      quantityInput.value = state.quantity;
    }
  } catch (error) {
    // ignore invalid saved state
  }
}

function calculateBunkering() {
  const quantityInput = document.getElementById('bunkeringQuantity');
  const tariffInput = document.getElementById('bunkeringTariff');
  const amountInput = document.getElementById('bunkeringAmount');
  if (!quantityInput || !tariffInput || !amountInput) return;

  const quantity = parseMoneyValue(quantityInput.value);
  const amount = quantity > 0 ? quantity * 0.3 : 0;

  const amountKey = getBunkeringAmountKey();
  if (amountKey) safeStorageSet(amountKey, amount.toFixed(2));

  tariffInput.value = '0,30';
  amountInput.value = formatMoneyValue(amount);
  saveBunkeringState();
}

function initBunkering() {
  const quantityInput = document.getElementById('bunkeringQuantity');
  if (!quantityInput) return;

  restoreBunkeringState(quantityInput);
  calculateBunkering();

  quantityInput.addEventListener('input', calculateBunkering);

  window.addEventListener('storage', (event) => {
    const stateKey = getBunkeringStateKey();
    if (!stateKey || event.key !== stateKey) return;
    restoreBunkeringState(quantityInput);
    calculateBunkering();
  });
}

function isBerthageAnchoragePdaPage() {
  return Boolean(document.body && document.body.classList.contains('page-berthage-anchorage-pda'));
}

function isBerthageAnchorageSailingPage() {
  return Boolean(document.body && document.body.classList.contains('page-berthage-anchorage-sailing'));
}

function getBerthageStateKey() {
  if (isBerthageAnchorageSailingPage()) return STORAGE_KEYS.berthageStateSailing;
  if (isBerthageAnchoragePdaPage()) return STORAGE_KEYS.berthageState;
  return null;
}

function getBerthageAmountKey() {
  if (isBerthageAnchorageSailingPage()) return STORAGE_KEYS.berthageAmountSailing;
  if (isBerthageAnchoragePdaPage()) return STORAGE_KEYS.berthageAmountPda;
  return null;
}

function setBerthageDaysInputState(toggleInput, daysInput) {
  if (!toggleInput || !daysInput) return;
  daysInput.disabled = !toggleInput.checked;
}

function saveBerthageState() {
  const stateKey = getBerthageStateKey();
  if (!stateKey) return;
  const berthageEnabled = document.getElementById('berthageEnabled');
  const berthageDays = document.getElementById('berthageDays');
  const anchorageEnabled = document.getElementById('anchorageEnabled');
  const anchorageDays = document.getElementById('anchorageDays');
  if (!berthageEnabled || !berthageDays || !anchorageEnabled || !anchorageDays) return;
  safeStorageSet(stateKey, JSON.stringify({
    berthageEnabled: Boolean(berthageEnabled.checked),
    berthageDays: berthageDays.value,
    anchorageEnabled: Boolean(anchorageEnabled.checked),
    anchorageDays: anchorageDays.value
  }));
}

function restoreBerthageState() {
  const stateKey = getBerthageStateKey();
  if (!stateKey) return;
  const berthageEnabled = document.getElementById('berthageEnabled');
  const berthageDays = document.getElementById('berthageDays');
  const anchorageEnabled = document.getElementById('anchorageEnabled');
  const anchorageDays = document.getElementById('anchorageDays');
  if (!berthageEnabled || !berthageDays || !anchorageEnabled || !anchorageDays) return;
  const raw = safeStorageGet(stateKey);
  if (!raw) return;
  try {
    const state = JSON.parse(raw);
    if (!state || typeof state !== 'object') return;
    if (typeof state.berthageEnabled === 'boolean') berthageEnabled.checked = state.berthageEnabled;
    if (typeof state.berthageDays === 'string') berthageDays.value = state.berthageDays;
    if (typeof state.anchorageEnabled === 'boolean') anchorageEnabled.checked = state.anchorageEnabled;
    if (typeof state.anchorageDays === 'string') anchorageDays.value = state.anchorageDays;
  } catch (error) {
    // ignore invalid saved state
  }
}

function calculateBerthageAnchorage() {
  const lengthInput = document.getElementById('berthageLengthOverall');
  const berthageEnabled = document.getElementById('berthageEnabled');
  const berthageDays = document.getElementById('berthageDays');
  const anchorageEnabled = document.getElementById('anchorageEnabled');
  const anchorageDays = document.getElementById('anchorageDays');
  const berthageTariff = document.getElementById('berthageTariff');
  const berthageDaysValue = document.getElementById('berthageDaysValue');
  const berthageAmount = document.getElementById('berthageAmount');
  const anchorageTariff = document.getElementById('anchorageTariff');
  const anchorageDaysValue = document.getElementById('anchorageDaysValue');
  const anchorageAmount = document.getElementById('anchorageAmount');
  if (
    !lengthInput ||
    !berthageEnabled || !berthageDays || !anchorageEnabled || !anchorageDays ||
    !berthageTariff || !berthageDaysValue || !berthageAmount ||
    !anchorageTariff || !anchorageDaysValue || !anchorageAmount
  ) return;

  setBerthageDaysInputState(berthageEnabled, berthageDays);
  setBerthageDaysInputState(anchorageEnabled, anchorageDays);

  const rawLength = getCurrentLengthOverallRawValue();
  if (lengthInput.value !== rawLength) lengthInput.value = rawLength;
  const lengthValue = parseMoneyValue(rawLength);

  const berthageDaysValueRaw = parseMoneyValue(berthageDays.value);
  const anchorageDaysValueRaw = parseMoneyValue(anchorageDays.value);

  const berthageCalcAmount = berthageEnabled.checked && lengthValue > 0 && berthageDaysValueRaw > 0
    ? lengthValue * BERTHAGE_TARIFF_RATE * berthageDaysValueRaw
    : 0;
  const anchorageCalcAmount = anchorageEnabled.checked && lengthValue > 0 && anchorageDaysValueRaw > 0
    ? lengthValue * ANCHORAGE_TARIFF_RATE * anchorageDaysValueRaw
    : 0;
  const totalAmount = berthageCalcAmount + anchorageCalcAmount;

  berthageTariff.value = formatMoneyValue(BERTHAGE_TARIFF_RATE);
  berthageDaysValue.value = berthageEnabled.checked ? (berthageDays.value.trim() || '0') : '0';
  berthageAmount.value = formatMoneyValue(berthageCalcAmount);
  anchorageTariff.value = formatMoneyValue(ANCHORAGE_TARIFF_RATE);
  anchorageDaysValue.value = anchorageEnabled.checked ? (anchorageDays.value.trim() || '0') : '0';
  anchorageAmount.value = formatMoneyValue(anchorageCalcAmount);

  const amountKey = getBerthageAmountKey();
  if (amountKey) safeStorageSet(amountKey, totalAmount.toFixed(2));
  saveBerthageState();
}

function initBerthageAnchorage() {
  const lengthInput = document.getElementById('berthageLengthOverall');
  const berthageEnabled = document.getElementById('berthageEnabled');
  const berthageDays = document.getElementById('berthageDays');
  const anchorageEnabled = document.getElementById('anchorageEnabled');
  const anchorageDays = document.getElementById('anchorageDays');
  if (!lengthInput || !berthageEnabled || !berthageDays || !anchorageEnabled || !anchorageDays) return;

  restoreBerthageState();
  calculateBerthageAnchorage();

  berthageEnabled.addEventListener('change', calculateBerthageAnchorage);
  berthageDays.addEventListener('input', calculateBerthageAnchorage);
  anchorageEnabled.addEventListener('change', calculateBerthageAnchorage);
  anchorageDays.addEventListener('input', calculateBerthageAnchorage);

  window.addEventListener('storage', (event) => {
    const stateKey = getBerthageStateKey();
    if (stateKey && event.key === stateKey) {
      restoreBerthageState();
      calculateBerthageAnchorage();
      return;
    }
    if (event.key === STORAGE_KEYS.indexState) {
      calculateBerthageAnchorage();
    }
  });
}

function isPilotBoatPdaPage() {
  return Boolean(document.body && document.body.classList.contains('page-pilot-boat-pda'));
}

function isPilotBoatSailingPage() {
  return Boolean(document.body && document.body.classList.contains('page-pilot-boat-sailing'));
}

function getPilotBoatStateKey() {
  if (isPilotBoatSailingPage()) return STORAGE_KEYS.pilotBoatStateSailing;
  if (isPilotBoatPdaPage()) return STORAGE_KEYS.pilotBoatState;
  return null;
}

function getPilotBoatAmountKey() {
  if (isPilotBoatSailingPage()) return STORAGE_KEYS.pilotBoatAmountSailing;
  if (isPilotBoatPdaPage()) return STORAGE_KEYS.pilotBoatAmountPda;
  return null;
}

function normalizePilotBoatOvertimeRate(value) {
  const raw = String(value == null ? '' : value).trim().toLowerCase();
  if (raw === '100' || raw === '1' || raw === '100%') return '100';
  if (raw === '50' || raw === '0.5' || raw === '50%') return '50';
  return 'normal';
}

function resolvePilotBoatOvertimeRate(rawValue, nightWeekend50, holiday100) {
  if (rawValue !== undefined && rawValue !== null && String(rawValue).trim() !== '') {
    return normalizePilotBoatOvertimeRate(rawValue);
  }
  if (Boolean(holiday100)) return '100';
  if (Boolean(nightWeekend50)) return '50';
  return 'normal';
}

function getPilotBoatOvertimeMultiplier(overtimeRate) {
  if (overtimeRate === '100') return 2;
  if (overtimeRate === '50') return 1.5;
  return 1;
}

function getDefaultPilotBoatCard() {
  return {
    extraNm: '',
    waitingEnabled: false,
    waitingMinutes: '',
    overtimeRate: 'normal'
  };
}

function getPilotBoatCardsFromState(state) {
  if (!state || typeof state !== 'object' || !Array.isArray(state.cards)) return [];
  return state.cards.map((card) => {
    const defaults = getDefaultPilotBoatCard();
    return {
      ...defaults,
      extraNm: String(card?.extraNm ?? defaults.extraNm),
      waitingEnabled: Boolean(card?.waitingEnabled),
      waitingMinutes: String(card?.waitingMinutes ?? defaults.waitingMinutes),
      overtimeRate: resolvePilotBoatOvertimeRate(card?.overtimeRate, card?.nightWeekend50, card?.holiday100)
    };
  });
}

let pilotBoatCardCount = 0;
let isRestoringPilotBoat = false;

function getPilotBoatCardState(id) {
  return {
    extraNm: document.getElementById(`pilotBoatExtraNm_${id}`)?.value || '',
    waitingEnabled: Boolean(document.getElementById(`pilotBoatWaitingEnabled_${id}`)?.checked),
    waitingMinutes: document.getElementById(`pilotBoatWaitingMinutes_${id}`)?.value || '',
    overtimeRate: normalizePilotBoatOvertimeRate(document.getElementById(`pilotBoatOvertime_${id}`)?.value)
  };
}

function getPilotBoatCardCalculation(card) {
  const extraNm = Math.max(parseMoneyValue(card.extraNm), 0);
  const waitingMinutes = Math.max(parseMoneyValue(card.waitingMinutes), 0);
  const waitingUnits = card.waitingEnabled ? Math.ceil(waitingMinutes / 15) : 0;

  const baseAmount = PILOT_BOAT_BASE_CHARGE;
  const extraNmAmount = extraNm * PILOT_BOAT_EXTRA_NM_RATE;
  const waitingAmount = waitingUnits * PILOT_BOAT_WAITING_15MIN_RATE;
  const subtotal = baseAmount + extraNmAmount + waitingAmount;

  const overtimeRate = normalizePilotBoatOvertimeRate(card?.overtimeRate);
  const multiplier = getPilotBoatOvertimeMultiplier(overtimeRate);

  return {
    baseAmount,
    extraNmAmount,
    waitingAmount,
    subtotal,
    multiplier,
    totalAmount: subtotal * multiplier
  };
}

function setPilotBoatWaitingState(id) {
  const enabledInput = document.getElementById(`pilotBoatWaitingEnabled_${id}`);
  const waitingWrap = document.getElementById(`pilotBoatWaitingWrap_${id}`);
  const waitingInput = document.getElementById(`pilotBoatWaitingMinutes_${id}`);
  if (!enabledInput || !waitingWrap || !waitingInput) return;

  const enabled = Boolean(enabledInput.checked);
  waitingWrap.style.display = enabled ? '' : 'none';
  waitingInput.disabled = !enabled;
}

function updatePilotBoatCardTitles() {
  const cards = document.querySelectorAll('#pilotBoatCards .card[data-pilot-boat-card]');
  cards.forEach((card, index) => {
    const title = card.querySelector('.pilot-boat-title');
    if (!title) return;
    if (cards.length === 1) {
      title.textContent = 'Pilot boat service';
      return;
    }
    const order = index + 1;
    title.textContent = `${order}${getOrdinal(order)} pilot boat service`;
  });
}

function addPilotBoatCard(initialState = {}) {
  const cardsWrap = document.getElementById('pilotBoatCards');
  if (!cardsWrap) return;

  const state = {
    ...getDefaultPilotBoatCard(),
    ...initialState
  };
  state.overtimeRate = resolvePilotBoatOvertimeRate(state.overtimeRate, state.nightWeekend50, state.holiday100);

  pilotBoatCardCount += 1;
  const id = pilotBoatCardCount;

  cardsWrap.insertAdjacentHTML('beforeend', `
    <div class="card" id="pilotBoat_${id}" data-pilot-boat-card="1" data-card-id="${id}">
      <div class="tug-header">
        <button class="icon-btn" onclick="removePilotBoatCard(${id})" aria-label="Remove pilot boat card">
          <img src="assets/icons/trash.svg" alt="Remove" />
        </button>
        <h3 class="pilot-boat-title">Pilot boat service</h3>
        <button class="icon-btn" onclick="duplicatePilotBoatCard(${id})" aria-label="Duplicate pilot boat card">
          <img src="assets/icons/copy-plus.svg" alt="Duplicate" />
        </button>
      </div>

      <div class="section-title">Base Charge</div>
      <div class="note">Up to 2 NM from the Port of Split breakwater: €${PILOT_BOAT_BASE_CHARGE.toFixed(2)} per departure.</div>

      <label>Additional NM beyond 2 NM boundary</label>
      <input id="pilotBoatExtraNm_${id}" type="number" min="0" step="0.1" value="${escapeHtml(state.extraNm)}" placeholder="0" />
      <div class="note">Additional charge: €${PILOT_BOAT_EXTRA_NM_RATE.toFixed(2)} per NM (to and from vessel).</div>

      <div class="section-title">Overtime</div>
      <label>Overtime Rate</label>
      <select id="pilotBoatOvertime_${id}">
        <option value="normal"${state.overtimeRate === 'normal' ? ' selected' : ''}>Normal working hours</option>
        <option value="50"${state.overtimeRate === '50' ? ' selected' : ''}>50% - 20:00 - 8:00 / Sat - Sun</option>
        <option value="100"${state.overtimeRate === '100' ? ' selected' : ''}>100% - Holidays</option>
      </select>

      <div class="section-title">Waiting Time</div>
      <label class="checkbox">
        <input id="pilotBoatWaitingEnabled_${id}" type="checkbox"${state.waitingEnabled ? ' checked' : ''} />
        Waiting time exceeds 15 minutes
      </label>
      <div id="pilotBoatWaitingWrap_${id}">
        <label>Minutes beyond first 15 minutes</label>
        <input id="pilotBoatWaitingMinutes_${id}" type="number" min="0" step="1" value="${escapeHtml(state.waitingMinutes)}" placeholder="0" />
        <div class="note">Charge: €${PILOT_BOAT_WAITING_15MIN_RATE.toFixed(2)} for each additional 15 minutes.</div>
      </div>

      <div class="tug-total" id="pilotBoatCardTotal_${id}">Card total: €0.00</div>
    </div>
  `);

  setPilotBoatWaitingState(id);
  updatePilotBoatCardTitles();
  if (!isRestoringPilotBoat) calculatePilotBoat();
}

function removePilotBoatCard(id) {
  const card = document.getElementById(`pilotBoat_${id}`);
  if (!card) return;
  card.remove();
  updatePilotBoatCardTitles();
  calculatePilotBoat();
}

function duplicatePilotBoatCard(id) {
  const card = document.getElementById(`pilotBoat_${id}`);
  if (!card) return;
  const state = getPilotBoatCardState(id);
  addPilotBoatCard(state);
  calculatePilotBoat();
}

function duplicateLastPilotBoatCard() {
  const cardsWrap = document.getElementById('pilotBoatCards');
  if (!cardsWrap) return;
  const cards = Array.from(cardsWrap.querySelectorAll('.card[data-pilot-boat-card]'));
  const last = cards[cards.length - 1];
  if (!last) {
    addPilotBoatCard();
    return;
  }
  duplicatePilotBoatCard(Number(last.dataset.cardId));
}

function savePilotBoatState() {
  const stateKey = getPilotBoatStateKey();
  if (!stateKey) return;

  const cardsWrap = document.getElementById('pilotBoatCards');
  if (!cardsWrap) return;

  const cards = Array.from(cardsWrap.querySelectorAll('.card[data-pilot-boat-card]')).map((card) => {
    const id = card.dataset.cardId;
    return getPilotBoatCardState(id);
  });

  safeStorageSet(stateKey, JSON.stringify({ cards }));
}

function restorePilotBoatState() {
  const cardsWrap = document.getElementById('pilotBoatCards');
  if (!cardsWrap) return;

  const stateKey = getPilotBoatStateKey();
  let state = null;
  if (stateKey) {
    state = parseStoredJson(safeStorageGet(stateKey));
  }

  const cards = getPilotBoatCardsFromState(state || {});
  isRestoringPilotBoat = true;
  cardsWrap.innerHTML = '';
  pilotBoatCardCount = 0;
  if (cards.length === 0) {
    addPilotBoatCard();
  } else {
    cards.forEach((cardState) => addPilotBoatCard(cardState));
  }
  isRestoringPilotBoat = false;
  updatePilotBoatCardTitles();
}

function calculatePilotBoat() {
  const cardsWrap = document.getElementById('pilotBoatCards');
  const final = document.getElementById('finalPilotBoatTotal');
  if (!cardsWrap || !final) return;

  const cards = Array.from(cardsWrap.querySelectorAll('.card[data-pilot-boat-card]'));
  let grandTotal = 0;
  const cardStates = [];

  cards.forEach((card) => {
    const id = card.dataset.cardId;
    const state = getPilotBoatCardState(id);
    cardStates.push(state);
    const calculation = getPilotBoatCardCalculation(state);
    grandTotal += calculation.totalAmount;

    const cardTotal = document.getElementById(`pilotBoatCardTotal_${id}`);
    if (cardTotal) {
      cardTotal.textContent = `Card total: €${calculation.totalAmount.toFixed(2)}`;
    }
  });

  final.style.display = 'block';
  final.innerHTML = `
    <div class="summary">
      <div><strong>Pilot boat services</strong><br>${cards.length}</div>
      <div><strong>Grand total</strong><br>€${grandTotal.toFixed(2)}</div>
    </div>
  `;

  const amountKey = getPilotBoatAmountKey();
  if (amountKey) {
    if (cards.length === 0) safeStorageRemove(amountKey);
    else safeStorageSet(amountKey, grandTotal.toFixed(2));
  }

  const stateKey = getPilotBoatStateKey();
  if (stateKey) {
    safeStorageSet(stateKey, JSON.stringify({ cards: cardStates }));
  }
}

function savePilotBoatAmount() {
  calculatePilotBoat();
}

function initPilotBoat() {
  const cardsWrap = document.getElementById('pilotBoatCards');
  const final = document.getElementById('finalPilotBoatTotal');
  if (!cardsWrap || !final) return;

  restorePilotBoatState();
  calculatePilotBoat();

  cardsWrap.addEventListener('input', () => {
    calculatePilotBoat();
  });

  cardsWrap.addEventListener('change', (event) => {
    const target = event.target;
    if (!target) return;
    if (target.id && target.id.startsWith('pilotBoatWaitingEnabled_')) {
      const id = target.id.replace('pilotBoatWaitingEnabled_', '');
      setPilotBoatWaitingState(id);
    }
    calculatePilotBoat();
  });

  window.addEventListener('storage', (event) => {
    const amountKey = getPilotBoatAmountKey();
    const stateKey = getPilotBoatStateKey();
    if (!amountKey || !stateKey) return;
    if (event.key !== amountKey && event.key !== stateKey) return;
    restorePilotBoatState();
    calculatePilotBoat();
  });

  window.addEventListener('beforeunload', savePilotBoatState);
  window.addEventListener('pagehide', savePilotBoatState);
}

function isPilotagePdaPage() {
  return Boolean(document.body && document.body.classList.contains('page-pilotage-pda'));
}

function isPilotageSailingPage() {
  return Boolean(document.body && document.body.classList.contains('page-pilotage-sailing'));
}

function normalizePilotageOperation(value) {
  if (value === 'departure') return 'departure';
  if (value === 'shifting') return 'shifting';
  return 'arrival';
}

function getPilotageStateKey() {
  if (isPilotageSailingPage()) return STORAGE_KEYS.pilotageStateSailing;
  if (isPilotagePdaPage()) return STORAGE_KEYS.pilotageState;
  return null;
}

function getPilotageAmountKey() {
  if (isPilotageSailingPage()) return STORAGE_KEYS.pilotageAmountSailing;
  if (isPilotagePdaPage()) return STORAGE_KEYS.pilotageAmountPda;
  return null;
}

function getPilotageBaseByGt(gt) {
  if (!Number.isFinite(gt) || gt <= 0) return 0;
  const matched = PILOTAGE_GT_TARIFFS.find((item) => gt >= item.min && gt <= item.max);
  if (matched) return matched.amount;
  return PILOTAGE_GT_TARIFFS[PILOTAGE_GT_TARIFFS.length - 1].amount;
}

function getDefaultPilotageCard(operation = 'arrival') {
  const op = normalizePilotageOperation(operation);
  return {
    operation: op,
    overtimeRate: 0,
    outOfPassengerPort: true,
    imoClasses: false,
    extrasEnabled: false,
    additionalPilot50: false,
    shipRelocation70: op === 'shifting',
    withoutPropulsion70: false
  };
}

function getPilotageCardCalculation(card, gt, options = {}) {
  const operation = normalizePilotageOperation(card.operation);
  const overtimeRate = Number(card.overtimeRate);
  const normalizedOvertimeRate = Number.isFinite(overtimeRate) && overtimeRate > 0 ? overtimeRate : 0;
  const outOfPassengerPort = card.outOfPassengerPort !== false;
  const forcedImo = options && typeof options.forceImo === 'boolean' ? options.forceImo : null;
  const imoClasses = forcedImo === null ? Boolean(card.imoClasses) : forcedImo;
  const extrasEnabled = Boolean(card.extrasEnabled);
  const additionalPilot50 = extrasEnabled && Boolean(card.additionalPilot50);
  const shipRelocation70 = extrasEnabled && operation === 'shifting' && Boolean(card.shipRelocation70);
  const withoutPropulsion70 = extrasEnabled && Boolean(card.withoutPropulsion70);
  const baseAmount = getPilotageBaseByGt(gt);
  const multiplier =
    1 +
    normalizedOvertimeRate +
    (outOfPassengerPort ? 0.2 : 0) +
    (imoClasses ? 0.2 : 0) +
    (additionalPilot50 ? 0.5 : 0) +
    (withoutPropulsion70 ? 0.7 : 0) +
    (operation === 'shifting' && shipRelocation70 ? 0.7 : 0);
  const totalAmount = baseAmount * multiplier;

  return {
    operation,
    baseAmount,
    multiplier,
    totalAmount
  };
}

function getPilotageCardsFromState(state) {
  if (!state || typeof state !== 'object') return [];

  if (Array.isArray(state.cards)) {
    return state.cards.map((card) => {
      const operation = normalizePilotageOperation(card?.operation);
      const baseCard = getDefaultPilotageCard(operation);
      let overtimeRate = Number(card?.overtimeRate);
      if (!Number.isFinite(overtimeRate) || overtimeRate < 0) {
        if (card?.overtime100) overtimeRate = 1;
        else if (card?.overtime50) overtimeRate = 0.5;
        else overtimeRate = 0;
      }
      const extrasEnabled =
        typeof card?.extrasEnabled === 'boolean'
          ? card.extrasEnabled
          : Boolean(card?.additionalPilot50) || Boolean(card?.withoutPropulsion70) || (operation === 'shifting' && Boolean(card?.shipRelocation70));

      return {
        ...baseCard,
        operation,
        overtimeRate,
        outOfPassengerPort: card?.outOfPassengerPort !== false,
        imoClasses: Boolean(card?.imoClasses),
        extrasEnabled,
        additionalPilot50: Boolean(card?.additionalPilot50),
        shipRelocation70: operation === 'shifting' ? card?.shipRelocation70 !== false : false,
        withoutPropulsion70: Boolean(card?.withoutPropulsion70)
      };
    });
  }

  const hasLegacyFields =
    Object.prototype.hasOwnProperty.call(state, 'overtimeRate') ||
    Object.prototype.hasOwnProperty.call(state, 'overtime50') ||
    Object.prototype.hasOwnProperty.call(state, 'overtime100') ||
    Object.prototype.hasOwnProperty.call(state, 'outOfPassengerPort') ||
    Object.prototype.hasOwnProperty.call(state, 'imoClasses') ||
    Object.prototype.hasOwnProperty.call(state, 'additionalPilot50') ||
    Object.prototype.hasOwnProperty.call(state, 'extrasEnabled') ||
    Object.prototype.hasOwnProperty.call(state, 'shipRelocation70') ||
    Object.prototype.hasOwnProperty.call(state, 'withoutPropulsion70');
  if (!hasLegacyFields) return [];

  let overtimeRate = Number(state.overtimeRate);
  if (!Number.isFinite(overtimeRate) || overtimeRate < 0) {
    if (state.overtime100) overtimeRate = 1;
    else if (state.overtime50) overtimeRate = 0.5;
    else overtimeRate = 0;
  }
  const extrasEnabled =
    typeof state.extrasEnabled === 'boolean'
      ? state.extrasEnabled
      : Boolean(state.additionalPilot50) || Boolean(state.shipRelocation70) || Boolean(state.withoutPropulsion70);

  return [
    {
      ...getDefaultPilotageCard('arrival'),
      overtimeRate,
      outOfPassengerPort: state.outOfPassengerPort !== false,
      imoClasses: Boolean(state.imoClasses),
      extrasEnabled,
      additionalPilot50: Boolean(state.additionalPilot50),
      shipRelocation70: false,
      withoutPropulsion70: Boolean(state.withoutPropulsion70)
    }
  ];
}

function getPilotageCalculationFromState(state, gtValue, options = {}) {
  if (!state || typeof state !== 'object') return null;
  const gt = Number.isFinite(gtValue) && gtValue > 0 ? gtValue : parseMoneyValue(state.gt);
  if (!Number.isFinite(gt) || gt <= 0) return null;

  const cards = getPilotageCardsFromState(state);
  if (cards.length === 0) return null;

  let arrivalTotal = 0;
  let departureTotal = 0;
  let shiftingTotal = 0;
  const calculatedCards = cards.map((card) => getPilotageCardCalculation(card, gt, options));
  calculatedCards.forEach((card) => {
    if (card.operation === 'departure') departureTotal += card.totalAmount;
    else if (card.operation === 'shifting') shiftingTotal += card.totalAmount;
    else arrivalTotal += card.totalAmount;
  });

  return {
    gt,
    cards: calculatedCards,
    arrivalTotal,
    departureTotal,
    shiftingTotal,
    totalAmount: arrivalTotal + departureTotal + shiftingTotal
  };
}

let pilotageCardCount = 0;
let isRestoringPilotage = false;
let pilotageImoMaster = null;

function getPilotageCardState(id) {
  const operation = normalizePilotageOperation(document.getElementById(`pilotageOp_${id}`)?.value);
  const overtimeRate = Number(document.getElementById(`pilotageOvertime_${id}`)?.value);
  const baseCard = getDefaultPilotageCard(operation);
  return {
    ...baseCard,
    operation,
    overtimeRate: Number.isFinite(overtimeRate) && overtimeRate >= 0 ? overtimeRate : 0,
    outOfPassengerPort: Boolean(document.getElementById(`pilotageOutOfPassenger_${id}`)?.checked),
    imoClasses: Boolean(document.getElementById(`pilotageImo_${id}`)?.checked),
    extrasEnabled: Boolean(document.getElementById(`pilotageExtrasEnabled_${id}`)?.checked),
    additionalPilot50: Boolean(document.getElementById(`pilotageAdditionalPilot_${id}`)?.checked),
    shipRelocation70: Boolean(document.getElementById(`pilotageShipRelocation70_${id}`)?.checked),
    withoutPropulsion70: Boolean(document.getElementById(`pilotageWithoutPropulsion70_${id}`)?.checked)
  };
}

function applyPilotageImoMaster() {
  if (!pilotageImoMaster) return;
  const checked = Boolean(pilotageImoMaster.checked);
  pilotageImoMaster.indeterminate = false;
  document.querySelectorAll('#pilotageCards .card[data-pilotage-card] input[id^="pilotageImo_"]').forEach((input) => {
    input.checked = checked;
  });
}

function syncPilotageImoMaster() {
  if (!pilotageImoMaster) return;
  const inputs = document.querySelectorAll('#pilotageCards .card[data-pilotage-card] input[id^="pilotageImo_"]');
  const total = inputs.length;
  if (total === 0) {
    pilotageImoMaster.checked = false;
    pilotageImoMaster.indeterminate = false;
    updateImoToggleLabelColor(pilotageImoMaster);
    return;
  }

  let checkedCount = 0;
  inputs.forEach((input) => {
    if (input.checked) checkedCount += 1;
  });

  if (checkedCount === 0) {
    pilotageImoMaster.checked = false;
    pilotageImoMaster.indeterminate = false;
    updateImoToggleLabelColor(pilotageImoMaster);
    return;
  }
  if (checkedCount === total) {
    pilotageImoMaster.checked = true;
    pilotageImoMaster.indeterminate = false;
    updateImoToggleLabelColor(pilotageImoMaster);
    return;
  }
  pilotageImoMaster.checked = false;
  pilotageImoMaster.indeterminate = true;
  updateImoToggleLabelColor(pilotageImoMaster);
}

function syncPilotageImoWithGlobalState(forceInitialize = false) {
  if (!pilotageImoMaster) return;
  const globalState = getGlobalImoTransportState();
  if (globalState === null) {
    if (forceInitialize && !pilotageImoMaster.indeterminate) {
      setGlobalImoTransportState(Boolean(pilotageImoMaster.checked));
    }
    return;
  }

  if (pilotageImoMaster.checked === globalState && !pilotageImoMaster.indeterminate) return;
  pilotageImoMaster.checked = globalState;
  pilotageImoMaster.indeterminate = false;
  applyPilotageImoMaster();
  syncPilotageImoMaster();
  updateImoToggleLabelColor(pilotageImoMaster);
}

function setPilotageExtrasState(id, defaultToChecked = false) {
  const opInput = document.getElementById(`pilotageOp_${id}`);
  const extrasEnabledInput = document.getElementById(`pilotageExtrasEnabled_${id}`);
  const extrasWrap = document.getElementById(`pilotageExtrasWrap_${id}`);
  const additionalPilotInput = document.getElementById(`pilotageAdditionalPilot_${id}`);
  const withoutPropulsionInput = document.getElementById(`pilotageWithoutPropulsion70_${id}`);
  const relocationRow = document.getElementById(`pilotageShipRelocationRow_${id}`);
  const relocationInput = document.getElementById(`pilotageShipRelocation70_${id}`);
  const toggleBtn = document.getElementById(`pilotageExtrasToggle_${id}`);
  const toggleLabel = toggleBtn ? toggleBtn.querySelector('.pilotage-extras-label') : null;
  if (!opInput || !extrasEnabledInput || !extrasWrap || !additionalPilotInput || !withoutPropulsionInput || !relocationRow || !relocationInput) return;

  const isShifting = normalizePilotageOperation(opInput.value) === 'shifting';
  const extrasEnabled = Boolean(extrasEnabledInput.checked);
  extrasWrap.style.display = extrasEnabled ? '' : 'none';
  additionalPilotInput.disabled = !extrasEnabled;
  withoutPropulsionInput.disabled = !extrasEnabled;
  if (toggleBtn) toggleBtn.setAttribute('aria-expanded', extrasEnabled ? 'true' : 'false');
  if (toggleLabel) toggleLabel.textContent = extrasEnabled ? 'Less' : 'More';

  relocationRow.style.display = isShifting ? '' : 'none';
  relocationInput.disabled = !extrasEnabled || !isShifting;
  if (isShifting && defaultToChecked && relocationInput.dataset.autoDefaulted !== '1') {
    relocationInput.checked = true;
    relocationInput.dataset.autoDefaulted = '1';
  }
}

function updatePilotageCardTitles() {
  const cards = document.querySelectorAll('#pilotageCards .card[data-pilotage-card]');
  const operations = Array.from(cards).map((card) => {
    const id = card.dataset.cardId;
    return normalizePilotageOperation(document.getElementById(`pilotageOp_${id}`)?.value);
  });
  const operationCounts = operations.reduce((acc, operation) => {
    acc[operation] = (acc[operation] || 0) + 1;
    return acc;
  }, {});

  let arrivalIndex = 0;
  let departureIndex = 0;
  let shiftingIndex = 0;
  cards.forEach((card) => {
    const id = card.dataset.cardId;
    const operation = normalizePilotageOperation(document.getElementById(`pilotageOp_${id}`)?.value);
    const title = card.querySelector('.pilotage-title');
    if (!title) return;

    if (operation === 'departure') {
      departureIndex += 1;
      const prefix = operationCounts.departure > 1 ? `${departureIndex}${getOrdinal(departureIndex)} ` : '';
      title.textContent = `${prefix}Pilot service on departure`;
      return;
    }

    if (operation === 'shifting') {
      shiftingIndex += 1;
      const prefix = operationCounts.shifting > 1 ? `${shiftingIndex}${getOrdinal(shiftingIndex)} ` : '';
      title.textContent = `${prefix}Pilot service shifting`;
      return;
    }

    arrivalIndex += 1;
    const prefix = operationCounts.arrival > 1 ? `${arrivalIndex}${getOrdinal(arrivalIndex)} ` : '';
    title.textContent = `${prefix}Pilot service on arrival`;
  });
}

function addPilotageCard(initialState = {}) {
  const cardsWrap = document.getElementById('pilotageCards');
  if (!cardsWrap) return;

  pilotageCardCount += 1;
  const id = pilotageCardCount;
  const state = {
    ...getDefaultPilotageCard('arrival'),
    ...initialState
  };
  const hasExplicitRelocation = Object.prototype.hasOwnProperty.call(initialState, 'shipRelocation70');
  const operation = normalizePilotageOperation(state.operation);
  const overtimeRate = Number(state.overtimeRate);
  const normalizedOvertimeRate = Number.isFinite(overtimeRate) && overtimeRate >= 0 ? overtimeRate : 0;

  cardsWrap.insertAdjacentHTML('beforeend', `
    <div class="card" id="pilotage_${id}" data-pilotage-card="1" data-card-id="${id}">
      <div class="tug-header">
        <button class="icon-btn" onclick="removePilotageCard(${id})" aria-label="Remove pilotage card">
          <img src="assets/icons/trash.svg" alt="Remove" />
        </button>
        <h3 class="pilotage-title">Pilot service</h3>
        <span></span>
      </div>

      <label>Operation</label>
      <select id="pilotageOp_${id}">
        <option value="arrival"${operation === 'arrival' ? ' selected' : ''}>Pilot service on arrival</option>
        <option value="departure"${operation === 'departure' ? ' selected' : ''}>Pilot service on departure</option>
        <option value="shifting"${operation === 'shifting' ? ' selected' : ''}>Pilot service shifting</option>
      </select>

      <div class="section-title">Additions</div>
      <label class="checkbox"><input id="pilotageOutOfPassenger_${id}" type="checkbox"${state.outOfPassengerPort !== false ? ' checked' : ''} /> 20% extra fee other ports</label>
      <label class="checkbox"><input id="pilotageImo_${id}" type="checkbox"${state.imoClasses ? ' checked' : ''} /> 20% pilotage ship transporting (IMO Code Classes I-III)</label>

      <div class="section-title">Overtime</div>
      <label>Overtime Rate</label>
      <select id="pilotageOvertime_${id}">
        <option value="0"${normalizedOvertimeRate === 0 ? ' selected' : ''}>Normal working hours</option>
        <option value="0.5"${normalizedOvertimeRate === 0.5 ? ' selected' : ''}>50% - 22:00 - 6:00 / Sat - Sun</option>
        <option value="1"${normalizedOvertimeRate === 1 ? ' selected' : ''}>100% - Holidays</option>
      </select>

      <div class="pilotage-extras-header">
        <button type="button" class="pilotage-extras-toggle" id="pilotageExtrasToggle_${id}" aria-expanded="${state.extrasEnabled ? 'true' : 'false'}" aria-controls="pilotageExtrasWrap_${id}">
          <span class="pilotage-extras-label">${state.extrasEnabled ? 'Less' : 'More'}</span>
          <span class="pilotage-extras-chevron" aria-hidden="true"></span>
        </button>
      </div>
      <input id="pilotageExtrasEnabled_${id}" type="checkbox"${state.extrasEnabled ? ' checked' : ''} hidden />
      <div id="pilotageExtrasWrap_${id}">
        <div id="pilotageShipRelocationRow_${id}">
          <label class="checkbox"><input id="pilotageShipRelocation70_${id}" type="checkbox"${state.shipRelocation70 ? ' checked' : ''} /> 70% ship relocation (shifting)</label>
        </div>
        <label class="checkbox"><input id="pilotageWithoutPropulsion70_${id}" type="checkbox"${state.withoutPropulsion70 ? ' checked' : ''} /> 70% ship without propulsion</label>
        <label class="checkbox"><input id="pilotageAdditionalPilot_${id}" type="checkbox"${state.additionalPilot50 ? ' checked' : ''} /> 50% additional pilot</label>
      </div>

      <div class="tug-total" id="pilotageCardTotal_${id}">Card total: €0.00</div>
    </div>
  `);

  setPilotageExtrasState(id, !hasExplicitRelocation);
  if (!isRestoringPilotage && pilotageImoMaster && pilotageImoMaster.checked) {
    const imoInput = document.getElementById(`pilotageImo_${id}`);
    if (imoInput) imoInput.checked = true;
  }
  updatePilotageCardTitles();
  syncPilotageImoMaster();
  if (!isRestoringPilotage) calculatePilotage();
}

function addPilotageServiceCard() {
  addPilotageCard({ operation: 'arrival' });
}

function removePilotageCard(id) {
  const card = document.getElementById(`pilotage_${id}`);
  if (!card) return;
  card.remove();
  updatePilotageCardTitles();
  syncPilotageImoMaster();
  calculatePilotage();
}

function savePilotageState() {
  const stateKey = getPilotageStateKey();
  if (!stateKey) return;
  const gtInput = document.getElementById('pilotageGt');
  const cardsWrap = document.getElementById('pilotageCards');
  if (!gtInput || !cardsWrap) return;

  const state = {
    gt: gtInput.value,
    cards: Array.from(cardsWrap.querySelectorAll('.card[data-pilotage-card]')).map((card) => {
      const id = card.dataset.cardId;
      return getPilotageCardState(id);
    })
  };
  safeStorageSet(stateKey, JSON.stringify(state));
}

function restorePilotageState() {
  const gtInput = document.getElementById('pilotageGt');
  const cardsWrap = document.getElementById('pilotageCards');
  if (!gtInput || !cardsWrap) return;

  const stateKey = getPilotageStateKey();
  let state = null;
  if (stateKey) {
    state = parseStoredJson(safeStorageGet(stateKey));
    if (state && typeof state.gt === 'string') gtInput.value = state.gt;
  }

  const sharedGt = safeStorageGet(STORAGE_KEYS.gt);
  if (sharedGt !== null) {
    gtInput.value = sharedGt;
  }

  const cards = getPilotageCardsFromState(state || {});
  isRestoringPilotage = true;
  cardsWrap.innerHTML = '';
  pilotageCardCount = 0;
  if (cards.length === 0) {
    addPilotageCard({ operation: 'arrival' });
    addPilotageCard({ operation: 'departure' });
  } else {
    cards.forEach((cardState) => addPilotageCard(cardState));
  }
  isRestoringPilotage = false;
  updatePilotageCardTitles();
  syncPilotageImoMaster();
}

function calculatePilotage() {
  const gtInput = document.getElementById('pilotageGt');
  const tariffInput = document.getElementById('pilotageTariff');
  const cardsWrap = document.getElementById('pilotageCards');
  const final = document.getElementById('finalPilotageTotal');
  if (!gtInput || !cardsWrap || !final) return;

  const gtRaw = gtInput.value.trim();
  if (gtRaw) safeStorageSet(STORAGE_KEYS.gt, gtRaw);
  else safeStorageRemove(STORAGE_KEYS.gt);
  const gt = parseMoneyValue(gtInput.value);
  const hasValidGt = Number.isFinite(gt) && gt > 0;
  if (tariffInput) {
    tariffInput.value = hasValidGt ? String(getPilotageBaseByGt(gt)) : '';
  }

  let arrivalTotal = 0;
  let departureTotal = 0;
  let shiftingTotal = 0;
  const cards = Array.from(cardsWrap.querySelectorAll('.card[data-pilotage-card]')).map((card) => {
    const id = card.dataset.cardId;
    return { id, state: getPilotageCardState(id) };
  });
  const cardStates = [];

  cards.forEach((card) => {
    cardStates.push(card.state);
    const operation = normalizePilotageOperation(card.state.operation);
    const totalEl = document.getElementById(`pilotageCardTotal_${card.id}`);
    if (!hasValidGt) {
      if (totalEl) totalEl.textContent = 'Card total: €0.00';
      return;
    }

    const calculation = getPilotageCardCalculation(card.state, gt);
    if (operation === 'departure') departureTotal += calculation.totalAmount;
    else if (operation === 'shifting') shiftingTotal += calculation.totalAmount;
    else arrivalTotal += calculation.totalAmount;
    if (totalEl) totalEl.textContent = `Card total: €${calculation.totalAmount.toFixed(2)}`;
  });

  const grandTotal = arrivalTotal + departureTotal + shiftingTotal;
  const amountKey = getPilotageAmountKey();
  final.style.display = 'block';
  final.innerHTML = `
    <div class="summary">
      <div><strong>Arrival total</strong><br>€${arrivalTotal.toFixed(2)}</div>
      <div><strong>Departure total</strong><br>€${departureTotal.toFixed(2)}</div>
      <div><strong>Shifting total</strong><br>€${shiftingTotal.toFixed(2)}</div>
      <div><strong>Grand total</strong><br>€${grandTotal.toFixed(2)}</div>
    </div>
  `;

  if (!hasValidGt) {
    if (amountKey) safeStorageRemove(amountKey);
  } else if (amountKey) {
    safeStorageSet(amountKey, grandTotal.toFixed(2));
  }

  const stateKey = getPilotageStateKey();
  if (stateKey) {
    safeStorageSet(stateKey, JSON.stringify({ gt: gtInput.value, cards: cardStates }));
  } else {
    savePilotageState();
  }
}

function initPilotage() {
  const gtInput = document.getElementById('pilotageGt');
  const cardsWrap = document.getElementById('pilotageCards');
  if (!gtInput || !cardsWrap) return;
  pilotageImoMaster = document.getElementById('pilotageImoMaster');

  restorePilotageState();
  syncPilotageImoMaster();
  syncPilotageImoWithGlobalState(true);
  calculatePilotage();
  syncPilotageImoMaster();

  if (pilotageImoMaster) {
    pilotageImoMaster.addEventListener('change', () => {
      applyPilotageImoMaster();
      setGlobalImoTransportState(Boolean(pilotageImoMaster.checked));
      updateImoToggleLabelColor(pilotageImoMaster);
      calculatePilotage();
    });
    updateImoToggleLabelColor(pilotageImoMaster);
  }

  gtInput.addEventListener('input', calculatePilotage);

  cardsWrap.addEventListener('input', (event) => {
    const target = event.target;
    if (!target) return;
    if (target.id && target.id.startsWith('pilotageOp_')) {
      const id = target.id.replace('pilotageOp_', '');
      setPilotageExtrasState(id, true);
      updatePilotageCardTitles();
    }
    calculatePilotage();
  });

  cardsWrap.addEventListener('change', (event) => {
    const target = event.target;
    if (!target) return;
    if (target.id && target.id.startsWith('pilotageOp_')) {
      const id = target.id.replace('pilotageOp_', '');
      setPilotageExtrasState(id, true);
      updatePilotageCardTitles();
    }
    if (target.id && target.id.startsWith('pilotageImo_')) {
      syncPilotageImoMaster();
      if (pilotageImoMaster && !pilotageImoMaster.indeterminate) {
        setGlobalImoTransportState(Boolean(pilotageImoMaster.checked));
      }
    }
    calculatePilotage();
  });

  cardsWrap.addEventListener('click', (event) => {
    const toggle = event.target.closest('.pilotage-extras-toggle');
    if (!toggle) return;
    const id = toggle.id.replace('pilotageExtrasToggle_', '');
    const checkbox = document.getElementById(`pilotageExtrasEnabled_${id}`);
    if (!checkbox) return;
    checkbox.checked = !checkbox.checked;
    setPilotageExtrasState(id, false);
    calculatePilotage();
  });

  const syncGtFromShared = () => {
    const sharedGt = safeStorageGet(STORAGE_KEYS.gt);
    const nextGt = sharedGt === null ? '' : sharedGt;
    if (gtInput.value !== nextGt) {
      gtInput.value = nextGt;
    }
    calculatePilotage();
  };

  window.addEventListener('storage', (event) => {
    if (event.key === STORAGE_KEYS.gt) {
      syncGtFromShared();
      return;
    }

    if (event.key === STORAGE_KEYS.globalImoTransport) {
      syncPilotageImoWithGlobalState();
      calculatePilotage();
    }
  });

  window.addEventListener('pageshow', () => {
    syncGtFromShared();
  });
}

function isMooringPdaPage() {
  return Boolean(document.body && document.body.classList.contains('page-mooring-pda'));
}

function isMooringSailingPage() {
  return Boolean(document.body && document.body.classList.contains('page-mooring-sailing'));
}

function normalizeMooringOperation(value) {
  if (value === 'departure') return 'departure';
  if (value === 'shifting') return 'shifting';
  return 'arrival';
}

function normalizeMooringOvertimeRate(value) {
  const raw = String(value == null ? '' : value).trim().toLowerCase();
  if (raw === '50' || raw === '0.5' || raw === '50%') return '50';
  if (raw === '25' || raw === '0.25' || raw === '25%') return '25';
  return 'normal';
}

function resolveMooringOvertimeRate(rawValue, overtime25, overtime50) {
  if (rawValue !== undefined && rawValue !== null && String(rawValue).trim() !== '') {
    return normalizeMooringOvertimeRate(rawValue);
  }
  if (Boolean(overtime50)) return '50';
  if (Boolean(overtime25)) return '25';
  return 'normal';
}

function getMooringOvertimeFactor(overtimeRate) {
  if (overtimeRate === '50') return 1.5;
  if (overtimeRate === '25') return 1.25;
  return 1;
}

function getMooringStateKey() {
  if (isMooringSailingPage()) return STORAGE_KEYS.mooringStateSailing;
  if (isMooringPdaPage()) return STORAGE_KEYS.mooringState;
  return null;
}

function getMooringAmountKey() {
  if (isMooringSailingPage()) return STORAGE_KEYS.mooringAmountSailing;
  if (isMooringPdaPage()) return STORAGE_KEYS.mooringAmountPda;
  return null;
}

function getMooringBaseByGt(gt) {
  if (!Number.isFinite(gt) || gt <= 0) return 0;
  const matched = MOORING_GT_TARIFFS.find((item) => gt >= item.min && gt <= item.max);
  if (matched) return matched.amount;

  const baseAt35000 = 285.6;
  const additionalThousands = Math.ceil((gt - 35000) / 1000);
  return baseAt35000 + Math.max(0, additionalThousands) * 10;
}

function setMooringAddonState(enabled, fieldsWrap, ...inputs) {
  if (fieldsWrap) {
    fieldsWrap.style.display = enabled ? '' : 'none';
  }
  inputs.forEach((input) => {
    if (!input) return;
    input.disabled = !enabled;
  });
}

function getMooringLegacyCardFromState(state) {
  const hasLegacyFields =
    Object.prototype.hasOwnProperty.call(state, 'overtimeRate') ||
    Object.prototype.hasOwnProperty.call(state, 'overtime25') ||
    Object.prototype.hasOwnProperty.call(state, 'overtime50') ||
    Object.prototype.hasOwnProperty.call(state, 'additionsEnabled') ||
    Object.prototype.hasOwnProperty.call(state, 'boatEnabled') ||
    Object.prototype.hasOwnProperty.call(state, 'boatHours') ||
    Object.prototype.hasOwnProperty.call(state, 'manEnabled') ||
    Object.prototype.hasOwnProperty.call(state, 'manHours') ||
    Object.prototype.hasOwnProperty.call(state, 'manPersons');
  if (!hasLegacyFields) return null;

  const fallbackAdditionsEnabled =
    Boolean(state.boatEnabled) ||
    Boolean(state.manEnabled) ||
    parseMoneyValue(state.boatHours) > 0 ||
    parseMoneyValue(state.manHours) > 0 ||
    parseMoneyValue(state.manPersons) > 0;
  const additionsEnabled = typeof state.additionsEnabled === 'boolean' ? state.additionsEnabled : fallbackAdditionsEnabled;
  const overtimeRate = resolveMooringOvertimeRate(state.overtimeRate, state.overtime25, state.overtime50);

  return {
    operation: normalizeMooringOperation(state.operation),
    overtimeRate,
    additionsEnabled,
    boatEnabled: Boolean(state.boatEnabled),
    boatHours: state.boatHours == null ? '' : String(state.boatHours),
    manEnabled: Boolean(state.manEnabled),
    manHours: state.manHours == null ? '' : String(state.manHours),
    manPersons: state.manPersons == null ? '' : String(state.manPersons)
  };
}

function getMooringCardsFromState(state) {
  if (!state || typeof state !== 'object') return [];

  if (Array.isArray(state.cards)) {
    return state.cards.map((card) => {
      const fallbackAdditionsEnabled =
        Boolean(card?.boatEnabled) ||
        Boolean(card?.manEnabled) ||
        parseMoneyValue(card?.boatHours) > 0 ||
        parseMoneyValue(card?.manHours) > 0 ||
        parseMoneyValue(card?.manPersons) > 0;
      const additionsEnabled = typeof card?.additionsEnabled === 'boolean' ? card.additionsEnabled : fallbackAdditionsEnabled;
      const overtimeRate = resolveMooringOvertimeRate(card?.overtimeRate, card?.overtime25, card?.overtime50);

      return {
        operation: normalizeMooringOperation(card?.operation),
        overtimeRate,
        additionsEnabled,
        boatEnabled: Boolean(card?.boatEnabled),
        boatHours: card?.boatHours == null ? '' : String(card.boatHours),
        manEnabled: Boolean(card?.manEnabled),
        manHours: card?.manHours == null ? '' : String(card.manHours),
        manPersons: card?.manPersons == null ? '' : String(card.manPersons)
      };
    });
  }

  const legacyCard = getMooringLegacyCardFromState(state);
  return legacyCard ? [legacyCard] : [];
}

function getMooringCardCalculation(cardState, gt) {
  const operation = normalizeMooringOperation(cardState?.operation);
  const overtimeRate = normalizeMooringOvertimeRate(cardState?.overtimeRate);
  const additionsEnabled = Boolean(cardState?.additionsEnabled);
  const boatEnabled = additionsEnabled && Boolean(cardState?.boatEnabled);
  const boatHours = parseMoneyValue(cardState?.boatHours);
  const manEnabled = additionsEnabled && Boolean(cardState?.manEnabled);
  const manHours = parseMoneyValue(cardState?.manHours);
  const manPersons = parseMoneyValue(cardState?.manPersons);

  const baseAmount = getMooringBaseByGt(gt);
  const overtimeFactor = getMooringOvertimeFactor(overtimeRate);
  const baseWithOvertime = baseAmount * overtimeFactor;
  const boatAmount = boatEnabled ? 125 * boatHours : 0;
  const manAmount = manEnabled ? 25 * manHours * manPersons : 0;
  const totalAmount = baseWithOvertime + boatAmount + manAmount;

  return {
    operation,
    overtimeRate,
    additionsEnabled,
    boatEnabled,
    boatHours,
    manEnabled,
    manHours,
    manPersons,
    baseAmount,
    overtimeFactor,
    baseWithOvertime,
    boatAmount,
    manAmount,
    totalAmount
  };
}

function getMooringCalculationFromState(state, gtValue) {
  if (!state || typeof state !== 'object') return null;
  const gt = Number.isFinite(gtValue) && gtValue > 0 ? gtValue : parseMoneyValue(state.gt);
  if (!Number.isFinite(gt) || gt <= 0) return null;

  const cardStates = getMooringCardsFromState(state);
  const cards = cardStates.map((cardState) => getMooringCardCalculation(cardState, gt));

  let arrivalTotal = 0;
  let departureTotal = 0;
  let shiftingTotal = 0;
  cards.forEach((card) => {
    if (card.operation === 'departure') {
      departureTotal += card.totalAmount;
    } else if (card.operation === 'shifting') {
      shiftingTotal += card.totalAmount;
    } else {
      arrivalTotal += card.totalAmount;
    }
  });
  const totalAmount = arrivalTotal + departureTotal + shiftingTotal;
  const firstCard = cards[0] || {
    baseAmount: 0,
    overtimeFactor: 1,
    baseWithOvertime: 0,
    boatAmount: 0,
    manAmount: 0
  };

  return {
    gt,
    cards,
    arrivalTotal,
    departureTotal,
    shiftingTotal,
    baseAmount: firstCard.baseAmount,
    overtimeFactor: firstCard.overtimeFactor,
    baseWithOvertime: firstCard.baseWithOvertime,
    boatAmount: firstCard.boatAmount,
    manAmount: firstCard.manAmount,
    totalAmount
  };
}

let mooringCardCount = 0;
let isRestoringMooring = false;

function getMooringCardState(id) {
  return {
    operation: normalizeMooringOperation(document.getElementById(`mooringOp_${id}`)?.value),
    overtimeRate: normalizeMooringOvertimeRate(document.getElementById(`mooringOtRate_${id}`)?.value),
    additionsEnabled: Boolean(document.getElementById(`mooringAdditionsEnabled_${id}`)?.checked),
    boatEnabled: Boolean(document.getElementById(`mooringBoatEnabled_${id}`)?.checked),
    boatHours: document.getElementById(`mooringBoatHours_${id}`)?.value || '',
    manEnabled: Boolean(document.getElementById(`mooringManEnabled_${id}`)?.checked),
    manHours: document.getElementById(`mooringManHours_${id}`)?.value || '',
    manPersons: document.getElementById(`mooringManPersons_${id}`)?.value || ''
  };
}

function setMooringAdditionsState(id) {
  const additionsEnabledInput = document.getElementById(`mooringAdditionsEnabled_${id}`);
  const additionsWrap = document.getElementById(`mooringAdditionsWrap_${id}`);
  const boatEnabledInput = document.getElementById(`mooringBoatEnabled_${id}`);
  const boatFields = document.getElementById(`mooringBoatFields_${id}`);
  const boatHoursInput = document.getElementById(`mooringBoatHours_${id}`);
  const manEnabledInput = document.getElementById(`mooringManEnabled_${id}`);
  const manFields = document.getElementById(`mooringManFields_${id}`);
  const manHoursInput = document.getElementById(`mooringManHours_${id}`);
  const manPersonsInput = document.getElementById(`mooringManPersons_${id}`);
  const toggleBtn = document.getElementById(`mooringAdditionsToggle_${id}`);
  const toggleLabel = toggleBtn ? toggleBtn.querySelector('.mooring-additions-label') : null;
  if (
    !additionsEnabledInput || !additionsWrap || !boatEnabledInput || !boatFields || !boatHoursInput ||
    !manEnabledInput || !manFields || !manHoursInput || !manPersonsInput
  ) return;

  const additionsEnabled = Boolean(additionsEnabledInput.checked);
  additionsWrap.style.display = additionsEnabled ? '' : 'none';
  boatEnabledInput.disabled = !additionsEnabled;
  manEnabledInput.disabled = !additionsEnabled;
  if (toggleBtn) toggleBtn.setAttribute('aria-expanded', additionsEnabled ? 'true' : 'false');
  if (toggleLabel) toggleLabel.textContent = additionsEnabled ? 'Less' : 'More';
  setMooringAddonState(additionsEnabled && Boolean(boatEnabledInput.checked), boatFields, boatHoursInput);
  setMooringAddonState(additionsEnabled && Boolean(manEnabledInput.checked), manFields, manHoursInput, manPersonsInput);
}

function updateMooringCardTitles() {
  const cards = document.querySelectorAll('#mooringCards .card[data-mooring-card]');
  const operations = Array.from(cards).map((card) => {
    const id = card.dataset.cardId;
    return normalizeMooringOperation(document.getElementById(`mooringOp_${id}`)?.value);
  });
  const operationCounts = operations.reduce((acc, operation) => {
    acc[operation] = (acc[operation] || 0) + 1;
    return acc;
  }, {});

  let arrivalIndex = 0;
  let departureIndex = 0;
  let shiftingIndex = 0;

  cards.forEach((card) => {
    const id = card.dataset.cardId;
    const operation = normalizeMooringOperation(document.getElementById(`mooringOp_${id}`)?.value);
    const title = card.querySelector('.mooring-title');
    if (!title) return;

    if (operation === 'departure') {
      departureIndex += 1;
      const prefix = operationCounts.departure > 1 ? `${departureIndex}${getOrdinal(departureIndex)} ` : '';
      title.textContent = `${prefix}Unmooring on departure`;
      return;
    }

    if (operation === 'shifting') {
      shiftingIndex += 1;
      const prefix = operationCounts.shifting > 1 ? `${shiftingIndex}${getOrdinal(shiftingIndex)} ` : '';
      title.textContent = `${prefix}Shifting berth to berth`;
      return;
    }

    arrivalIndex += 1;
    const prefix = operationCounts.arrival > 1 ? `${arrivalIndex}${getOrdinal(arrivalIndex)} ` : '';
    title.textContent = `${prefix}Mooring on arrival`;
  });
}

function addMooringCard(initialState = {}) {
  const cardsWrap = document.getElementById('mooringCards');
  if (!cardsWrap) return;

  mooringCardCount += 1;
  const id = mooringCardCount;
  const operation = normalizeMooringOperation(initialState.operation);
  const overtimeRate = resolveMooringOvertimeRate(initialState.overtimeRate, initialState.overtime25, initialState.overtime50);
  const derivedAdditionsEnabled =
    Boolean(initialState.boatEnabled) ||
    Boolean(initialState.manEnabled) ||
    parseMoneyValue(initialState.boatHours) > 0 ||
    parseMoneyValue(initialState.manHours) > 0 ||
    parseMoneyValue(initialState.manPersons) > 0;
  const additionsEnabled = typeof initialState.additionsEnabled === 'boolean' ? initialState.additionsEnabled : derivedAdditionsEnabled;

  cardsWrap.insertAdjacentHTML('beforeend', `
    <div class="card" id="mooring_${id}" data-mooring-card="1" data-card-id="${id}">
      <div class="tug-header">
        <button class="icon-btn" onclick="removeMooringCard(${id})" aria-label="Remove mooring card">
          <img src="assets/icons/trash.svg" alt="Remove" />
        </button>
        <h3 class="mooring-title">Mooring</h3>
        <span></span>
      </div>

      <label>Operation</label>
      <select id="mooringOp_${id}">
        <option value="arrival"${operation === 'arrival' ? ' selected' : ''}>Mooring on arrival</option>
        <option value="departure"${operation === 'departure' ? ' selected' : ''}>Unmooring on departure</option>
        <option value="shifting"${operation === 'shifting' ? ' selected' : ''}>Shifting berth to berth</option>
      </select>

      <div class="section-title">Overtime Rate</div>
      <select id="mooringOtRate_${id}">
        <option value="normal"${overtimeRate === 'normal' ? ' selected' : ''}>Normal working hours</option>
        <option value="25"${overtimeRate === '25' ? ' selected' : ''}>25% - 22:00 - 6:00 / Mon-Sun</option>
        <option value="50"${overtimeRate === '50' ? ' selected' : ''}>50% - Holidays</option>
      </select>

      <div class="mooring-additions-header">
        <button type="button" class="mooring-additions-toggle" id="mooringAdditionsToggle_${id}" aria-expanded="${additionsEnabled ? 'true' : 'false'}" aria-controls="mooringAdditionsWrap_${id}">
          <span class="mooring-additions-label">${additionsEnabled ? 'Less' : 'More'}</span>
          <span class="mooring-additions-chevron" aria-hidden="true"></span>
        </button>
      </div>
      <input id="mooringAdditionsEnabled_${id}" type="checkbox"${additionsEnabled ? ' checked' : ''} hidden />
      <div id="mooringAdditionsWrap_${id}">
        <label class="checkbox"><input id="mooringBoatEnabled_${id}" type="checkbox" /> Mooring boat (125 EUR / hour)</label>
        <div id="mooringBoatFields_${id}">
          <label>Boat Hours</label>
          <input id="mooringBoatHours_${id}" type="number" min="0" step="0.25" placeholder="0" />
        </div>

        <label class="checkbox"><input id="mooringManEnabled_${id}" type="checkbox" /> Additional mooring man (25 EUR / hour)</label>
        <div id="mooringManFields_${id}">
          <div class="row">
            <div>
              <label>Man Hours</label>
              <input id="mooringManHours_${id}" type="number" min="0" step="0.25" placeholder="0" />
            </div>
            <div>
              <label>Persons</label>
              <input id="mooringManPersons_${id}" type="number" min="0" step="1" placeholder="0" />
            </div>
          </div>
        </div>
      </div>

      <div class="tug-total" id="mooringCardTotal_${id}">Card total: €0.00</div>
    </div>
  `);

  if (Object.prototype.hasOwnProperty.call(initialState, 'additionsEnabled')) {
    const input = document.getElementById(`mooringAdditionsEnabled_${id}`);
    if (input) input.checked = Boolean(initialState.additionsEnabled);
  }
  if (Object.prototype.hasOwnProperty.call(initialState, 'boatEnabled')) {
    const input = document.getElementById(`mooringBoatEnabled_${id}`);
    if (input) input.checked = Boolean(initialState.boatEnabled);
  }
  if (typeof initialState.boatHours === 'string') {
    const input = document.getElementById(`mooringBoatHours_${id}`);
    if (input) input.value = initialState.boatHours;
  }
  if (Object.prototype.hasOwnProperty.call(initialState, 'manEnabled')) {
    const input = document.getElementById(`mooringManEnabled_${id}`);
    if (input) input.checked = Boolean(initialState.manEnabled);
  }
  if (typeof initialState.manHours === 'string') {
    const input = document.getElementById(`mooringManHours_${id}`);
    if (input) input.value = initialState.manHours;
  }
  if (typeof initialState.manPersons === 'string') {
    const input = document.getElementById(`mooringManPersons_${id}`);
    if (input) input.value = initialState.manPersons;
  }

  setMooringAdditionsState(id);

  updateMooringCardTitles();
  if (!isRestoringMooring) calculateMooring();
}

function addMooringArrivalCard() {
  addMooringCard({ operation: 'arrival' });
}

function addMooringDepartureCard() {
  addMooringCard({ operation: 'departure' });
}

function removeMooringCard(id) {
  const card = document.getElementById(`mooring_${id}`);
  if (!card) return;
  card.remove();
  updateMooringCardTitles();
  calculateMooring();
}

function saveMooringState() {
  const stateKey = getMooringStateKey();
  if (!stateKey) return;

  const gtInput = document.getElementById('mooringGt');
  const cardsWrap = document.getElementById('mooringCards');
  if (!gtInput || !cardsWrap) return;

  const cards = Array.from(cardsWrap.querySelectorAll('.card[data-mooring-card]')).map((card) => {
    const id = card.dataset.cardId;
    return getMooringCardState(id);
  });

  const state = {
    gt: gtInput.value,
    cards
  };
  safeStorageSet(stateKey, JSON.stringify(state));
}

function restoreMooringState() {
  const gtInput = document.getElementById('mooringGt');
  const cardsWrap = document.getElementById('mooringCards');
  if (!gtInput || !cardsWrap) return;

  const stateKey = getMooringStateKey();
  let state = null;
  if (stateKey) {
    state = parseStoredJson(safeStorageGet(stateKey));
    if (state && typeof state.gt === 'string') gtInput.value = state.gt;
  }

  const sharedGt = safeStorageGet(STORAGE_KEYS.gt);
  if (sharedGt !== null) {
    gtInput.value = sharedGt;
  }

  const cards = getMooringCardsFromState(state || {});
  isRestoringMooring = true;
  cardsWrap.innerHTML = '';
  mooringCardCount = 0;
  if (cards.length === 0) {
    addMooringCard({ operation: 'arrival' });
    addMooringCard({ operation: 'departure' });
  } else {
    cards.forEach((cardState) => addMooringCard(cardState));
    const hasArrival = cards.some((cardState) => normalizeMooringOperation(cardState.operation) === 'arrival');
    const hasDeparture = cards.some((cardState) => normalizeMooringOperation(cardState.operation) === 'departure');
    if (!hasArrival) addMooringCard({ operation: 'arrival' });
    if (!hasDeparture) addMooringCard({ operation: 'departure' });
  }
  isRestoringMooring = false;
  updateMooringCardTitles();
}

function calculateMooring() {
  const gtInput = document.getElementById('mooringGt');
  const tariffInput = document.getElementById('mooringTariff');
  const cardsWrap = document.getElementById('mooringCards');
  const final = document.getElementById('finalMooringTotal');
  if (!gtInput || !cardsWrap || !final) return;

  const gtRaw = gtInput.value.trim();
  if (gtRaw) safeStorageSet(STORAGE_KEYS.gt, gtRaw);
  else safeStorageRemove(STORAGE_KEYS.gt);
  const gt = parseMoneyValue(gtInput.value);
  const hasValidGt = Number.isFinite(gt) && gt > 0;
  if (tariffInput) {
    tariffInput.value = hasValidGt ? String(getMooringBaseByGt(gt)) : '';
  }

  let arrivalTotal = 0;
  let departureTotal = 0;
  let shiftingTotal = 0;
  const cardStates = [];
  const cards = Array.from(cardsWrap.querySelectorAll('.card[data-mooring-card]'));
  cards.forEach((card) => {
    const id = card.dataset.cardId;
    const cardState = getMooringCardState(id);
    cardStates.push(cardState);

    const cardTotal = document.getElementById(`mooringCardTotal_${id}`);
    if (!hasValidGt) {
      if (cardTotal) cardTotal.textContent = 'Card total: €0.00';
      return;
    }

    const calculation = getMooringCardCalculation(cardState, gt);
    if (calculation.operation === 'departure') {
      departureTotal += calculation.totalAmount;
    } else if (calculation.operation === 'shifting') {
      shiftingTotal += calculation.totalAmount;
    } else {
      arrivalTotal += calculation.totalAmount;
    }
    if (cardTotal) cardTotal.textContent = `Card total: €${calculation.totalAmount.toFixed(2)}`;
  });

  const grandTotal = arrivalTotal + departureTotal + shiftingTotal;
  final.style.display = 'block';
  final.innerHTML = `
    <div class="summary">
      <div><strong>Arrival total</strong><br>€${arrivalTotal.toFixed(2)}</div>
      <div><strong>Departure total</strong><br>€${departureTotal.toFixed(2)}</div>
      <div><strong>Shifting total</strong><br>€${shiftingTotal.toFixed(2)}</div>
      <div><strong>Grand total</strong><br>€${grandTotal.toFixed(2)}</div>
    </div>
  `;

  const amountKey = getMooringAmountKey();

  if (!hasValidGt) {
    if (amountKey) safeStorageRemove(amountKey);
  } else if (amountKey) {
    safeStorageSet(amountKey, grandTotal.toFixed(2));
  }

  const stateKey = getMooringStateKey();
  if (stateKey) {
    safeStorageSet(stateKey, JSON.stringify({ gt: gtInput.value, cards: cardStates }));
  } else {
    saveMooringState();
  }
}

function initMooring() {
  const gtInput = document.getElementById('mooringGt');
  const cardsWrap = document.getElementById('mooringCards');
  if (!gtInput || !cardsWrap) return;

  restoreMooringState();
  calculateMooring();

  gtInput.addEventListener('input', calculateMooring);

  cardsWrap.addEventListener('input', (event) => {
    const target = event.target;
    if (!target) return;
    if (target.id && target.id.startsWith('mooringOp_')) {
      updateMooringCardTitles();
    }
    calculateMooring();
  });

  cardsWrap.addEventListener('change', (event) => {
    const target = event.target;
    if (!target) return;
    if (target.id && target.id.startsWith('mooringOp_')) {
      updateMooringCardTitles();
    }
    if (target.id && target.id.startsWith('mooringBoatEnabled_')) {
      const id = target.id.replace('mooringBoatEnabled_', '');
      setMooringAdditionsState(id);
    }
    if (target.id && target.id.startsWith('mooringManEnabled_')) {
      const id = target.id.replace('mooringManEnabled_', '');
      setMooringAdditionsState(id);
    }
    calculateMooring();
  });

  cardsWrap.addEventListener('click', (event) => {
    const toggle = event.target.closest('.mooring-additions-toggle');
    if (!toggle) return;
    const id = toggle.id.replace('mooringAdditionsToggle_', '');
    const checkbox = document.getElementById(`mooringAdditionsEnabled_${id}`);
    if (!checkbox) return;
    checkbox.checked = !checkbox.checked;
    setMooringAdditionsState(id);
    calculateMooring();
  });

  const syncGtFromShared = () => {
    const sharedGt = safeStorageGet(STORAGE_KEYS.gt);
    const nextGt = sharedGt === null ? '' : sharedGt;
    if (gtInput.value !== nextGt) {
      gtInput.value = nextGt;
    }
    calculateMooring();
  };

  window.addEventListener('storage', (event) => {
    if (event.key !== STORAGE_KEYS.gt) return;
    syncGtFromShared();
  });

  window.addEventListener('pageshow', () => {
    syncGtFromShared();
  });
}

// Tug calculator logic
let tugCount = 0;
const MIN_VOYAGE = 1;
const MIN_ASSIST = 0.5;
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
const TUG_OT_RATE_NORMAL = '0';
const TUG_OT_RATE_25 = '0.25';
const TUG_OT_RATE_50 = '0.50';
const TUG_STANDBY_RATE = 168;
const TUG_STANDBY_MIN_GT = 10000;

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

function addTug() {
  const tugCards = document.getElementById('tugCards');
  if (!tugCards) return;
  tugCount++;
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

  for (const [key, prop] of fields) {
    const from = document.getElementById(`${key}_${id}`);
    const to = document.getElementById(`${key}_${newId}`);
    if (!from || !to) continue;
    to[prop] = from[prop];
  }
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
  const cards = getTugServiceCards();
  cards.forEach(card => card.classList.remove('selected'));
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
  for (let i = 1; i <= tugCount; i++) {
    const cb = document.getElementById(`imo_${i}`);
    if (cb) cb.checked = checked;
  }
}

function applyLinesMaster() {
  if (!linesMaster) return;
  const checked = linesMaster.checked;
  linesMaster.indeterminate = false;
  for (let i = 1; i <= tugCount; i++) {
    const cb = document.getElementById(`lines_${i}`);
    if (cb) cb.checked = checked;
  }
}

function syncImoMaster() {
  if (!imoMaster) return;
  let total = 0;
  let checked = 0;
  for (let i = 1; i <= tugCount; i++) {
    const cb = document.getElementById(`imo_${i}`);
    if (!cb) continue;
    total++;
    if (cb.checked) checked++;
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
  for (let i = 1; i <= tugCount; i++) {
    const cb = document.getElementById(`lines_${i}`);
    if (!cb) continue;
    total++;
    if (cb.checked) checked++;
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
  for (let i = 1; i <= tugCount; i++) {
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
    const rate = getMasterOvertimeRate(departureOtMaster, departureOtSunday);
    applyOvertimeMaster('departure', rate, id);
    return;
  }
  const rate = getMasterOvertimeRate(arrivalOtMaster, arrivalOtSunday);
  applyOvertimeMaster('arrival', rate, id);
}

function applyKwVoyageRule(id) {
  const kwCheckbox = document.getElementById(`kw_${id}`);
  const voyageSelect = document.getElementById(`voyage_${id}`);
  if (!voyageSelect) return;
  if (kwCheckbox && kwCheckbox.checked) {
    voyageSelect.value = '1.5';
  } else {
    voyageSelect.value = '1';
  }
}

function updateOvertimeMasterVisibility() {
  const arrivalWrap = document.getElementById('arrivalOtSundayWrap');
  if (arrivalWrap) arrivalWrap.hidden = !arrivalOtMaster?.checked;
  if (!arrivalOtMaster?.checked && arrivalOtSunday) arrivalOtSunday.checked = false;
  const departureWrap = document.getElementById('departureOtSundayWrap');
  if (departureWrap) departureWrap.hidden = !departureOtMaster?.checked;
  if (!departureOtMaster?.checked && departureOtSunday) departureOtSunday.checked = false;
}

function updateDiscountEnabledState() {
  if (!discountPercentInput) return;
  const isEnabled = Boolean(discountEnabled && discountEnabled.checked);
  if (isEnabled) {
    discountPercentInput.disabled = false;
    if (discountFieldWrap) discountFieldWrap.hidden = false;
    return;
  }
  discountPercentInput.disabled = true;
  if (discountFieldWrap) discountFieldWrap.hidden = true;
}

function getApprovedDiscountPercent() {
  if (!discountEnabled || !discountEnabled.checked) return 0;
  const raw = discountPercentInput ? discountPercentInput.value : '';
  const percent = parsePercentValue(raw);
  if (!Number.isFinite(percent) || percent <= 0) return 0;
  return Math.min(percent, 100);
}

function formatDiscountLabel(percent) {
  if (!Number.isFinite(percent) || percent <= 0) return '';
  const rounded = Math.round(percent * 1000) / 1000;
  let text = rounded.toFixed(3);
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
  if (standbyCard) {
    standbyCard.hidden = !isEligible || !standbyEnabled.checked;
  }
  if (standbyHoursInput) {
    standbyHoursInput.disabled = !isEligible || !standbyEnabled.checked;
  }
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


function updateTugTitles() {
  const cards = getTugServiceCards();
  const operations = Array.from(cards).map((card) => {
    const id = card.id.split('_')[1];
    return document.getElementById(`op_${id}`)?.value === 'departure' ? 'departure' : 'arrival';
  });
  const operationCounts = operations.reduce((acc, operation) => {
    acc[operation] = (acc[operation] || 0) + 1;
    return acc;
  }, {});

  let arrivalIndex = 0;
  let departureIndex = 0;

  cards.forEach(card => {
    const id = card.id.split('_')[1];
    const op = document.getElementById(`op_${id}`)?.value;
    card.setAttribute('data-operation', op === 'departure' ? 'departure' : 'arrival');
    const title = card.querySelector('.tug-title');
    if (!title) return;

    if (op === 'arrival') {
      arrivalIndex++;
      card.style.setProperty('--print-row', String(arrivalIndex));
      const prefix = operationCounts.arrival > 1 ? `${arrivalIndex}${getOrdinal(arrivalIndex)} ` : '';
      title.innerText = `${prefix}Tugboat on arrival`;
    } else {
      departureIndex++;
      card.style.setProperty('--print-row', String(departureIndex));
      const prefix = operationCounts.departure > 1 ? `${departureIndex}${getOrdinal(departureIndex)} ` : '';
      title.innerText = `${prefix}Tugboat on departure`;
    }
  });
}

function getOrdinal(n) {
  if (n % 10 === 1 && n % 100 !== 11) return 'st';
  if (n % 10 === 2 && n % 100 !== 12) return 'nd';
  if (n % 10 === 3 && n % 100 !== 13) return 'rd';
  return 'th';
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
  const imoEnabled = Boolean(source?.imo);
  const linesEnabled = Boolean(source?.lines);
  const kwEnabled = Boolean(source?.kw);

  const voyageBase = tariff * voyage;
  const voyageOvertimeCharge = voyageBase * voyageOT;
  const voyageSubtotal = voyageBase + voyageOvertimeCharge;

  const assistBase = tariff * assist;
  const assistOvertimeCharge = assistBase * assistOT;
  const assistSubtotal = assistBase + assistOvertimeCharge;

  const imoCharge = imoEnabled ? (assistSubtotal * 0.20) : 0;
  const linesCharge = linesEnabled ? (assistSubtotal * 0.15) : 0;
  const kwCharge = kwEnabled ? (assistSubtotal * 0.30) : 0;

  return {
    voyage,
    assist,
    voyageOT,
    assistOT,
    imoEnabled,
    linesEnabled,
    kwEnabled,
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

  return {
    id,
    operation,
    ...calculation
  };
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
  const voyageLabel = getSelectedOptionText(`voyage_${id}`) || `${calc.voyage} hour`;
  const assistLabel = getSelectedOptionText(`assist_${id}`) || `${calc.assist} hour`;
  const voyageOtLabel = getSelectedOptionText(`voy_ot_${id}`) || '';
  const assistOtLabel = getSelectedOptionText(`assist_ot_${id}`) || '';

  const rows = [
    { amount: calc.voyageBase, label: `Voyage base to base (${voyageLabel})` },
    { amount: calc.assistBase, label: `Assistance (${assistLabel})` }
  ];

  if (calc.voyageOvertimeCharge > 0) rows.push({ amount: calc.voyageOvertimeCharge, label: `Voyage overtime (${voyageOtLabel})` });
  if (calc.assistOvertimeCharge > 0) rows.push({ amount: calc.assistOvertimeCharge, label: `Assistance overtime (${assistOtLabel})` });
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
        <tr>
          <th colspan="3">${escapeHtml(cardTitle)}</th>
        </tr>
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
  const enteredHoursLabel = Number.isFinite(hours) && hours > 0 ? formatMoneyValue(hours) : '0.00';

  standbyBreakdownEl.innerHTML = `
    <table class="tug-breakdown-print-table">
      <thead>
        <tr>
          <th colspan="3">Stand By Tugboat</th>
        </tr>
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
  const tariffInput = document.getElementById('tariff');
  const final = document.getElementById('finalTotal');
  if (!tariffInput || !final) return;

  const tariff = Number(tariffInput.value) || 0;
  if (!tariff) {
    for (let i = 1; i <= tugCount; i++) {
      if (!document.getElementById(`tug_${i}`)) continue;
      const totalEl = document.getElementById(`tugTotal_${i}`);
      if (totalEl) totalEl.innerText = 'Tug total: €0.00';
      renderTugPrintBreakdown(i, null);
    }
    if (standbyTotalEl) standbyTotalEl.innerText = 'Stand by Tugboat total: €0.00';
    renderTugStandbyPrintBreakdown(0, '');
    final.style.display = 'none';
    const keys = getTugStorageKeys(isSailingTugsPage());
    safeStorageSet(keys.towageArrivalCount, 0);
    safeStorageSet(keys.towageDepartureCount, 0);
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

  for (let i = 1; i <= tugCount; i++) {
    if (!document.getElementById(`tug_${i}`)) continue;

    const calc = getTugCardCalculation(i, effectiveTariff);
    if (!calc) continue;

    if (calc.operation === 'arrival') arrivalCount += 1;
    else departureCount += 1;

    const cardTotalText = Number.isFinite(calc.total) ? calc.total.toFixed(2) : '0.00';
    document.getElementById(`tugTotal_${i}`).innerText = `Tug total: €${cardTotalText}${discountSuffix}`;
    renderTugPrintBreakdown(i, calc);

    if (calc.operation === 'arrival') arrivalTotal += calc.total;
    else departureTotal += calc.total;
  }

  if (arrivalTotal === 0 && departureTotal === 0) {
    if (standbyCharge <= 0) {
      final.style.display = 'none';
    }
    const keys = getTugStorageKeys(isSailingTugsPage());
    if (Number.isFinite(standbyCharge) && standbyCharge > 0) {
      safeStorageSet(keys.towageTotal, standbyCharge.toFixed(2));
    }
    safeStorageSet(keys.towageArrivalCount, arrivalCount);
    safeStorageSet(keys.towageDepartureCount, departureCount);
    if (standbyCharge <= 0) return;
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

  if (Number.isFinite(grandTotal)) {
    const keys = getTugStorageKeys(isSailingTugsPage());
    safeStorageSet(keys.towageTotal, grandTotal.toFixed(2));
    safeStorageSet(keys.towageArrivalCount, arrivalCount);
    safeStorageSet(keys.towageDepartureCount, departureCount);
  }
}

function initTugs() {
  const tugCards = document.getElementById('tugCards');
  if (!tugCards) return;

  const copyTugsFromPdaBtn = document.getElementById('copyTugsFromPdaBtn');
  if (copyTugsFromPdaBtn && isSailingTugsPage()) {
    copyTugsFromPdaBtn.addEventListener('click', () => {
      const copied = copyTugboatsFromPdaToSailing();
      if (!copied) {
        window.alert('No tugboat cards found in PDA to copy.');
      }
    });
  }

  const vesselNameInput = document.getElementById('vesselName');
  if (vesselNameInput) {
    updateVesselNameFromStorage(vesselNameInput);
    window.addEventListener('storage', (event) => {
      if (event.key === STORAGE_KEYS.vesselName) {
        updateVesselNameFromStorage(vesselNameInput);
      }
    });
  }

  const gtInput = document.getElementById('gt');
  if (gtInput) {
    const syncGtAndTariff = () => {
      const storedGt = safeStorageGet(STORAGE_KEYS.gt);
      if (storedGt !== null && gtInput.value !== storedGt) {
        gtInput.value = storedGt;
      }
      const gt = Number(gtInput.value);
      const tariff = getTariffFromGT(gt);
      const tariffInput = document.getElementById('tariff');
      if (tariffInput) tariffInput.value = tariff || '';
      updateTugStandbyVisibility();
    };

    syncGtAndTariff();
    gtInput.addEventListener('input', () => {
      const gt = Number(gtInput.value);
      const tariff = getTariffFromGT(gt);
      const tariffInput = document.getElementById('tariff');
      if (tariffInput) tariffInput.value = tariff || '';
      const value = gtInput.value.trim();
      if (value) safeStorageSet(STORAGE_KEYS.gt, value);
      else safeStorageRemove(STORAGE_KEYS.gt);
      updateTugStandbyVisibility();
      calculate();
    });

    window.addEventListener('storage', (event) => {
      if (event.key === STORAGE_KEYS.gt) {
        syncGtAndTariff();
        calculate();
      }
    });

    window.addEventListener('pageshow', () => {
      syncGtAndTariff();
      calculate();
    });
  }

  document.addEventListener('input', calculate);
  document.addEventListener('change', (event) => {
    const target = event.target;
    if (target && target.id && target.id.startsWith('op_')) {
      const id = Number(target.id.replace('op_', ''));
      if (Number.isFinite(id)) {
        applyOvertimeMastersToCard(id);
      }
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
  const tugOptionsToggle = document.getElementById('tugOptionsToggle');
  const tugOptionsBody = document.getElementById('tugOptionsBody');
  if (tugOptionsToggle && tugOptionsBody) {
    const setOptionsExpanded = (expanded) => {
      tugOptionsBody.hidden = !expanded;
      tugOptionsToggle.setAttribute('aria-expanded', expanded ? 'true' : 'false');
    };
    tugOptionsToggle.addEventListener('click', () => {
      const isExpanded = tugOptionsToggle.getAttribute('aria-expanded') === 'true';
      setOptionsExpanded(!isExpanded);
    });
    setOptionsExpanded(false);
  }

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
      const rate = getMasterOvertimeRate(arrivalOtMaster, arrivalOtSunday);
      applyOvertimeMaster('arrival', rate);
      calculate();
    });
  }
  if (arrivalOtSunday) {
    arrivalOtSunday.addEventListener('change', () => {
      const rate = getMasterOvertimeRate(arrivalOtMaster, arrivalOtSunday);
      applyOvertimeMaster('arrival', rate);
      calculate();
    });
  }
  if (departureOtMaster) {
    departureOtMaster.addEventListener('change', () => {
      updateOvertimeMasterVisibility();
      const rate = getMasterOvertimeRate(departureOtMaster, departureOtSunday);
      applyOvertimeMaster('departure', rate);
      calculate();
    });
  }
  if (departureOtSunday) {
    departureOtSunday.addEventListener('change', () => {
      const rate = getMasterOvertimeRate(departureOtMaster, departureOtSunday);
      applyOvertimeMaster('departure', rate);
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
    discountPercentInput.addEventListener('input', () => {
      calculate();
    });
  }
  if (standbyEnabled) {
    standbyEnabled.addEventListener('change', () => {
      updateTugStandbyVisibility();
      calculate();
    });
  }
  if (standbyHoursInput) {
    standbyHoursInput.addEventListener('input', () => {
      calculate();
    });
  }

  document.addEventListener('change', (event) => {
    const target = event.target;
    if (target && target.id && target.id.startsWith('imo_')) {
      syncImoMaster();
      if (imoMaster && !imoMaster.indeterminate) {
        setGlobalImoTransportState(Boolean(imoMaster.checked));
      }
    }
    if (target && target.id && target.id.startsWith('lines_')) {
      syncLinesMaster();
    }
    if (target && target.id && target.id.startsWith('kw_')) {
      const id = Number(target.id.replace('kw_', ''));
      if (Number.isFinite(id)) {
        applyKwVoyageRule(id);
      }
    }
  });

  restoreTugsState();
  updateTugStandbyVisibility();
  syncTugImoWithGlobalState(true);
  updateOvertimeMasterVisibility();
  updateDiscountEnabledState();
  calculate();

  window.addEventListener('storage', (event) => {
    if (event.key !== STORAGE_KEYS.globalImoTransport) return;
    syncTugImoWithGlobalState();
    calculate();
  });

  window.addEventListener('beforeunload', saveTugsState);
  window.addEventListener('pagehide', saveTugsState);
}

if (typeof window !== 'undefined' && window.addEventListener) {
  window.addEventListener('afterprint', () => {
    document.body.classList.remove('print-fit', 'print-desktop-lock', 'print-mobile');
    disableMobilePrintFooterSuppression();
    if (printRestoreDensity === 'comfortable') {
      setDensity('comfortable');
    } else if (printRestoreDensity === 'none') {
      document.body.classList.remove('density-comfortable', 'density-dense');
    }
    printRestoreDensity = null;
    restorePrintTitleIfNeeded();
    clearPrintHidden();
  });

  window.addEventListener('beforeprint', () => {
    setPrintTitleFromVessel();
    const logoLeftNote = document.getElementById('logoLeftNote');
    if (logoLeftNote) autoResizeTextarea(logoLeftNote);
    if (document.body.classList.contains('page-index') && shouldSuppressMobilePrintFooters()) {
      enableMobilePrintFooterSuppression();
    } else {
      disableMobilePrintFooterSuppression();
    }
    if (document.body.classList.contains('page-index') && isLikelyMobileViewport()) {
      document.body.classList.add('print-mobile');
    } else {
      document.body.classList.remove('print-mobile');
    }
    applyPrintDensity();
    updatePrintHidden();
  });

  window.addEventListener('pageshow', () => {
    clearTransientIconAnimationState(document);
  });

  window.addEventListener('DOMContentLoaded', () => {
    document.addEventListener('input', (event) => {
      const target = event.target;
      if (!target || !(target instanceof HTMLElement)) return;
      if (target.matches('input, textarea, select')) {
        target.dataset.userEdited = '1';
      }
    }, true);
    document.addEventListener('change', (event) => {
      const target = event.target;
      if (!target || !(target instanceof HTMLElement)) return;
      if (target.matches('input, textarea, select')) {
        target.dataset.userEdited = '1';
      }
    }, true);
    initIconClickAnimationCompletion();
    clearTransientIconAnimationState(document);
    initIndex();
    initLightDues();
    initPortDues();
    initBunkering();
    initBerthageAnchorage();
    initPilotBoat();
    initPilotage();
    initMooring();
    initTugs();
  });
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    __test__: {
      STORAGE_KEYS,
      getTugChargeBreakdown,
      getTugStandbyCharge,
      getGlobalImoTransportState,
      resolveLightDuesTypeForImo,
      resolvePortDuesCargoTypeForImo,
      setGlobalImoTransportState,
      shouldApplyGlobalImoStateOnInit
    }
  };
}
