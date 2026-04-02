/**
 * @file Investown Partner Summary — Chrome Extension Content Script
 * @description Displays total investment across all projects by the same partner (borrower)
 *              on the Investown.cz property detail and listing pages. Operates as a content
 *              script injected into the Investown SPA, intercepting route changes and injecting
 *              summary widgets and listing-page badges via DOM manipulation.
 */
'use strict';

// ===========================================================================
// Constants
// ===========================================================================

const API_URL = 'https://api.investown.cz/core/api/graphql';
const WIDGET_ID = 'investown-ext-partner-summary';
const MAX_PAGES = 5;
const PER_PAGE = 100;
const CACHE_TTL_MS = 5 * 60 * 1000; // 5 minutes

const WIDGET_REINSERTION_DEBOUNCE_MS = 150;
const LISTING_MUTATION_DEBOUNCE_MS = 300;
const WIDGET_INSERTION_RETRY_BASE_MS = 500;
const MAX_WIDGET_INSERTION_RETRIES = 5;
const MAX_CONCURRENT_API_REQUESTS = 3;
const DIVERSIFICATION_SAFE_THRESHOLD = 0.15;    // < 15% = safe (green)
const DIVERSIFICATION_WARNING_THRESHOLD = 0.25;  // 15-25% = warning (orange), > 25% = danger (red)

const REPAYMENT_STATUS = {
  DELAYED: 'Delayed',
  COLLECTION: 'Collection',
  EXITED_OR_REPAID: 'ExitedOrRepaid',
  REPAID: 'Repaid',
};

const BADGE_ATTR = 'data-ext-badge';
const PROPERTY_LINK_SELECTOR = 'a[href^="/property/"]';

// ===========================================================================
// Safe API References (MAIN world protection against page overrides)
// ===========================================================================

const safeFetch = fetch.bind(window);
const safeJsonParse = JSON.parse;
const safeJsonStringify = JSON.stringify;
const safeAtob = atob;
const safeGetItem = localStorage.getItem.bind(localStorage);
const safeObjectKeys = Object.keys;

// ===========================================================================
// TTL Cache
// ===========================================================================

/**
 * A simple cache with time-to-live expiration and maximum size limit.
 * When the cache exceeds maxSize, the oldest entries are evicted (FIFO).
 */
class TTLCache {
  /**
   * @param {number} ttlMs - Time-to-live in milliseconds. Entries older than this are expired.
   * @param {number} [maxSize=200] - Maximum number of entries before FIFO eviction.
   */
  constructor(ttlMs, maxSize = 200) {
    this._ttlMs = ttlMs;
    this._maxSize = maxSize;
    this._map = new Map();
  }

  /**
   * Retrieves a cached value if it exists and hasn't expired.
   * @param {string} key - The cache key.
   * @returns {*|undefined} The cached value, or undefined if missing/expired.
   */
  get(key) {
    const entry = this._map.get(key);
    if (!entry) return undefined;
    if (Date.now() - entry.timestamp > this._ttlMs) {
      this._map.delete(key);
      return undefined;
    }
    return entry.value;
  }

  /**
   * Stores a value in the cache with the current timestamp.
   * Evicts the oldest entry if the cache exceeds maxSize.
   * @param {string} key - The cache key.
   * @param {*} value - The value to store.
   */
  set(key, value) {
    this._map.delete(key); // re-insert to refresh insertion order
    this._map.set(key, { value, timestamp: Date.now() });
    if (this._map.size > this._maxSize) {
      const oldest = this._map.keys().next().value;
      this._map.delete(oldest);
    }
  }

  /**
   * Checks whether a non-expired entry exists for the given key.
   * @param {string} key - The cache key.
   * @returns {boolean}
   */
  has(key) {
    return this.get(key) !== undefined;
  }

  /**
   * Removes an entry from the cache.
   * @param {string} key - The cache key.
   */
  delete(key) {
    this._map.delete(key);
  }
}

// ===========================================================================
// State Variables
// ===========================================================================

// Current navigation state
let currentSlug = null;
let currentPageType = null;
let listingNavigationId = 0; // incremented on each listing page entry, used as stale guard

// Caches
const summaryCache = new TTLCache(CACHE_TTL_MS);
const propertyMetaCache = new TTLCache(CACHE_TTL_MS);
const partnerCache = new TTLCache(CACHE_TTL_MS);

// In-flight request tracking
const pendingFetches = new Set();    // slugs currently being fetched
const pendingPartnerFetches = new Map(); // companyId → Promise<Summary>
const pendingPropertyFetches = new Map(); // slug → Promise<Object>

// Portfolio statistics cache
const portfolioStatsCache = new TTLCache(CACHE_TTL_MS, 1);

// Observers and timers
let listingObserver = null;
let listingDebounceTimer = null;
let reinsertTimer = null;
// Re-inject widget if React removes it during DOM reconciliation (debounced).
// Managed via connectWidgetObserver/disconnectWidgetObserver — only active on property pages.
const widgetObserver = new MutationObserver(() => {
  if (currentPageType === 'property' && currentSlug && !document.getElementById(WIDGET_ID)) {
    clearTimeout(reinsertTimer);
    reinsertTimer = setTimeout(() => ensureWidgetPresent(currentSlug), WIDGET_REINSERTION_DEBOUNCE_MS);
  }
});

// ===========================================================================
// Utility Functions
// ===========================================================================

/**
 * Formats a numeric value as Czech Koruna (CZK) currency string.
 * Uses cs-CZ locale formatting with "Kc" suffix.
 * @param {number} value - The monetary value to format.
 * @returns {string} Formatted string, e.g. "1 234 567 Kc" or "1 234,50 Kc".
 */
function formatCZK(value) {
  return value.toLocaleString('cs-CZ', {
    minimumFractionDigits: value % 1 === 0 ? 0 : 2,
    maximumFractionDigits: 2,
  }) + ' Kč';
}

/**
 * Retrieves a cached summary for a property slug if it exists and hasn't expired.
 * @param {string} slug - The property slug to look up.
 * @returns {Object|null} The cached summary data, or null if missing or expired.
 */
function getCached(slug) {
  return summaryCache.get(slug) || null;
}

// ===========================================================================
// API Layer — Authentication
// ===========================================================================

/**
 * Decodes a Base64URL-encoded string to a regular string.
 * Handles the URL-safe alphabet (+/- substitution) and padding.
 * @param {string} str - The Base64URL-encoded string.
 * @returns {string} The decoded string.
 */
function base64UrlDecode(str) {
  str = str.replace(/-/g, '+').replace(/_/g, '/');
  const pad = str.length % 4;
  if (pad) str += '='.repeat(4 - pad);
  return safeAtob(str);
}

/**
 * Extracts the Cognito JWT ID token from localStorage.
 * Searches for a key matching the AWS Amplify/Cognito pattern (ending in '.idToken'),
 * decodes the JWT payload, and validates that the token hasn't expired.
 * @returns {string|null} The raw JWT token string, or null if not found or expired.
 */
function getAuthToken() {
  const keys = safeObjectKeys(localStorage);
  const idTokenKey = keys.find(key => key.endsWith('.idToken'));
  if (!idTokenKey) return null;

  const token = safeGetItem(idTokenKey);
  if (!token) return null;

  try {
    const payload = safeJsonParse(base64UrlDecode(token.split('.')[1]));
    if (payload.exp && payload.exp < Date.now() / 1000) {
      return null;
    }
  } catch {
    return null;
  }

  return token;
}

// ===========================================================================
// API Layer — GraphQL
// ===========================================================================

/**
 * Executes a GraphQL query against the Investown API.
 * @param {string} query - The GraphQL query string.
 * @param {Object} variables - Query variables.
 * @param {string} token - JWT bearer token for authentication.
 * @returns {Promise<Object>} The `data` field from the GraphQL response.
 * @throws {Error} If the HTTP response is not OK or the response contains GraphQL errors.
 */
async function executeGraphQLQuery(query, variables, token) {
  const response = await safeFetch(API_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${token}`,
    },
    body: safeJsonStringify({ query, variables }),
  });
  if (!response.ok) throw new Error(`API ${response.status}`);
  const json = await response.json();
  if (json.errors) throw new Error(json.errors[0].message);
  return json.data;
}

/**
 * Fetches a single property by its slug, including the current investment round
 * and borrower (partner) details.
 * @param {string} slug - The property slug.
 * @param {string} token - JWT bearer token.
 * @returns {Promise<Object>} The property object from the API.
 */
async function fetchProperty(slug, token) {
  const query = `
    query Property($slug: String!) {
      property(slug: $slug) {
        id
        name
        currentInvestmentRound {
          currentUsersTotalInvestment
          repaymentStatus
          borrower {
            companyName
            companyIdentifier
          }
        }
      }
    }
  `;
  const data = await executeGraphQLQuery(query, { slug }, token);
  if (!data?.property) throw new Error('Missing property in API response');
  return data.property;
}

/**
 * Fetches all related properties for a given slug using pagination.
 * Iterates through up to {@link MAX_PAGES} pages of {@link PER_PAGE} results each,
 * stopping early if a page returns fewer results than requested.
 * @param {string} slug - The property slug to find related properties for.
 * @param {string} token - JWT bearer token.
 * @returns {Promise<Object[]>} Array of related property objects with investment and interest data.
 */
async function fetchAllRelatedProperties(slug, token) {
  const query = `
    query RelatedProperties($slug: String!, $page: Int, $perPage: Int) {
      relatedProperties(slug: $slug, page: $page, perPage: $perPage) {
        investmentAmount
        interestAmount { value currency }
        name
        slug
      }
    }
  `;

  const allRelatedProperties = [];
  for (let page = 0; page < MAX_PAGES; page++) {
    const data = await executeGraphQLQuery(query, { slug, page, perPage: PER_PAGE }, token);
    const batch = data.relatedProperties || [];
    allRelatedProperties.push(...batch);
    if (batch.length < PER_PAGE) break;
  }
  return allRelatedProperties;
}

/**
 * Fetches the user's total portfolio size for diversification calculations.
 * Uses a simple cache with TTL to avoid redundant API calls.
 * @param {string} token - JWT bearer token.
 * @returns {Promise<number>} The total portfolio size in CZK.
 */
async function fetchPortfolioStatistics(token) {
  const cached = portfolioStatsCache.get('size');
  if (cached !== undefined) return cached;

  const query = `
    query PortfolioStatistics {
      portfolioStatistics {
        portfolioSize
      }
    }
  `;
  const data = await executeGraphQLQuery(query, {}, token);
  const size = data.portfolioStatistics?.portfolioSize || 0;
  portfolioStatsCache.set('size', size);
  return size;
}

// ===========================================================================
// API Layer — Concurrency Limiter & Cached Fetches
// ===========================================================================

/**
 * Creates a concurrency limiter that ensures at most `maxConcurrent` async tasks
 * run simultaneously. Additional tasks are queued in FIFO order and executed as
 * slots become available.
 * @param {number} maxConcurrent - Maximum number of tasks to run in parallel.
 * @returns {function(function(): Promise): Promise} The `enqueue` function — accepts
 *   an async task factory and returns a promise that resolves with the task's result.
 */
function createLimiter(maxConcurrent) {
  let running = 0;
  const queue = [];
  function next() {
    if (running >= maxConcurrent || queue.length === 0) return;
    running++;
    const { task, resolve, reject } = queue.shift();
    task().then(resolve, reject).finally(() => { running--; next(); });
  }
  return function enqueue(task) {
    return new Promise((resolve, reject) => {
      queue.push({ task, resolve, reject });
      next();
    });
  };
}

const apiLimiter = createLimiter(MAX_CONCURRENT_API_REQUESTS);

/**
 * Fetches a property with caching. Returns the cached result if available,
 * otherwise fetches from the API and stores the result in {@link propertyMetaCache}.
 * @param {string} slug - The property slug.
 * @param {string} token - JWT bearer token.
 * @returns {Promise<Object>} The property object.
 */
async function fetchPropertyCached(slug, token) {
  if (propertyMetaCache.has(slug)) return propertyMetaCache.get(slug);
  if (pendingPropertyFetches.has(slug)) return pendingPropertyFetches.get(slug);

  const promise = fetchProperty(slug, token)
    .then(property => { propertyMetaCache.set(slug, property); return property; })
    .finally(() => pendingPropertyFetches.delete(slug));

  pendingPropertyFetches.set(slug, promise);
  return promise;
}

/**
 * Fetches partner summary data for a property slug. Uses two-level caching:
 * first at the property level ({@link propertyMetaCache}), then at the partner
 * level ({@link partnerCache}) keyed by company identifier. This ensures that
 * multiple properties from the same partner share a single summary.
 * @param {string} slug - The property slug.
 * @param {string} token - JWT bearer token.
 * @returns {Promise<Object|null>} The partner summary, or null if no borrower found.
 */
async function fetchPartnerDataForSlug(slug, token) {
  const property = await fetchPropertyCached(slug, token);
  const companyId = property.currentInvestmentRound?.borrower?.companyIdentifier;
  if (!companyId) return null;

  if (partnerCache.has(companyId)) return partnerCache.get(companyId);
  if (pendingPartnerFetches.has(companyId)) return pendingPartnerFetches.get(companyId);

  const promise = fetchAllRelatedProperties(slug, token)
    .then(related => {
      const summary = computeSummary(property, related);
      partnerCache.set(companyId, summary);
      return summary;
    })
    .finally(() => pendingPartnerFetches.delete(companyId));

  pendingPartnerFetches.set(companyId, promise);
  return promise;
}

// ===========================================================================
// Computation
// ===========================================================================

/**
 * Computes an aggregated partner summary from a property and its related properties.
 * Filters to CZK-denominated investments only (or investments with no currency specified),
 * sums investment amounts and yields, and counts projects with active investments.
 * The current property's investment is added on top of the related properties total.
 * @param {Object} property - The primary property object (from fetchProperty).
 * @param {Object[]} relatedProperties - Array of related property objects.
 * @returns {{ partnerName: string, totalInvestment: number, totalYields: number, investedCount: number, totalCount: number }}
 */
function computeSummary(property, relatedProperties) {
  const round = property.currentInvestmentRound;
  const borrower = round?.borrower;
  const currentInvestment = round?.currentUsersTotalInvestment || 0;

  const czkRelated = relatedProperties.filter(
    rp => rp.interestAmount?.currency === 'CZK' || !rp.interestAmount?.currency
  );

  const relatedInvestment = czkRelated.reduce(
    (sum, rp) => sum + (rp.investmentAmount || 0), 0
  );

  const totalYields = relatedProperties
    .filter(rp => rp.interestAmount?.currency === 'CZK')
    .reduce((sum, rp) => sum + (rp.interestAmount?.value || 0), 0);

  const investedCount =
    czkRelated.filter(rp => rp.investmentAmount > 0).length +
    (currentInvestment > 0 ? 1 : 0);

  const totalCount = relatedProperties.length + 1;

  const totalInvestment = currentInvestment + relatedInvestment;

  // Repayment status breakdown — only from current property (relatedProperties don't expose repaymentStatus)
  const repaymentBreakdown = { regular: 0, delayed: 0, collection: 0, repaid: 0 };

  if (currentInvestment > 0) {
    const currentStatus = round?.repaymentStatus;
    if (currentStatus === REPAYMENT_STATUS.DELAYED) {
      repaymentBreakdown.delayed++;
    } else if (currentStatus === REPAYMENT_STATUS.COLLECTION) {
      repaymentBreakdown.collection++;
    } else if (currentStatus === REPAYMENT_STATUS.EXITED_OR_REPAID || currentStatus === REPAYMENT_STATUS.REPAID) {
      repaymentBreakdown.repaid++;
    } else {
      repaymentBreakdown.regular++;
    }
  }

  return {
    partnerName: borrower?.companyName || 'Neznámý partner',
    totalInvestment,
    totalYields,
    investedCount,
    totalCount,
    repaymentBreakdown,
  };
}

// ===========================================================================
// DOM Rendering
// ===========================================================================

/**
 * Creates a widget row element with a label and value.
 * @param {string} label - The row label text.
 * @param {string} value - The row value text.
 * @param {string} [valueClass] - Additional CSS class for the value span.
 * @returns {HTMLDivElement} The row element.
 */
function createWidgetRow(label, value, valueClass) {
  const row = document.createElement('div');
  row.className = 'ext-partner-row';

  const labelEl = document.createElement('span');
  labelEl.className = 'ext-partner-label';
  labelEl.textContent = label;

  const valueEl = document.createElement('span');
  valueEl.className = valueClass ? 'ext-partner-value ' + valueClass : 'ext-partner-value';
  valueEl.textContent = value;

  row.append(labelEl, valueEl);
  return row;
}

/**
 * Creates a spacer element matching the native Investown empty-span spacing pattern.
 * @param {boolean} [large=false] - If true, uses larger spacing (12px instead of 4px).
 * @returns {HTMLSpanElement} The spacer element.
 */
function createWidgetSpacer(large) {
  const spacer = document.createElement('span');
  spacer.className = large ? 'ext-partner-spacer ext-partner-spacer--lg' : 'ext-partner-spacer';
  return spacer;
}

/**
 * Creates the repayment status row element for the widget.
 * @param {{ regular: number, delayed: number, collection: number, repaid: number }} breakdown - Repayment counts.
 * @returns {HTMLDivElement|null} The repayment row element, or null if all counts are zero.
 */
function createRepaymentRow(breakdown) {
  if (!breakdown) return null;
  const { regular, delayed, collection, repaid } = breakdown;
  if (regular === 0 && delayed === 0 && collection === 0 && repaid === 0) return null;

  const row = document.createElement('div');
  row.className = 'ext-partner-row ext-partner-row--repayment';

  const label = document.createElement('span');
  label.className = 'ext-partner-label';
  label.textContent = 'Stav splácení';

  const valueContainer = document.createElement('span');
  valueContainer.className = 'ext-partner-value ext-partner-repayment';

  const parts = [
    { count: regular, cls: 'ext-repayment--regular', text: '✓ ' },
    { count: delayed, cls: 'ext-repayment--delayed', text: '⚠ ' },
    { count: collection, cls: 'ext-repayment--collection', text: '✕ ' },
    { count: repaid, cls: 'ext-repayment--repaid', text: '✓ ', suffix: ' splac.' },
  ];

  let first = true;
  for (const { count, cls, text, suffix } of parts) {
    if (count <= 0) continue;
    if (!first) valueContainer.append(' · ');
    first = false;
    const span = document.createElement('span');
    span.className = cls;
    span.textContent = text + count + (suffix || '');
    valueContainer.appendChild(span);
  }

  row.append(label, valueContainer);
  return row;
}

/**
 * Renders the partner summary widget as a DOM element for the property detail page.
 * Produces either an error message or a full summary with investment totals.
 * @param {Object} options - Render options.
 * @param {Object} [options.summary] - The computed partner summary to display.
 * @param {string} [options.error] - Error type: 'auth' for unauthenticated, 'fetch' for API failure.
 * @returns {HTMLDivElement} Widget element ready for insertion into the DOM.
 */
function renderWidget({ summary, error }) {
  const widget = document.createElement('div');
  widget.id = WIDGET_ID;

  if (error) {
    const messages = {
      auth: 'Nepřihlášen — nelze načíst data',
      fetch: 'Chyba při načítání dat partnera',
    };
    const errorEl = document.createElement('div');
    errorEl.className = 'ext-partner-error';
    errorEl.textContent = messages[error] || 'Neočekávaná chyba';
    widget.appendChild(errorEl);
    return widget;
  }

  // Section heading — matches native "Investice" heading style
  const heading = document.createElement('span');
  heading.className = 'ext-partner-heading';
  heading.textContent = 'Partner';
  widget.appendChild(heading);

  // Partner name — styled as a link-like element
  const name = document.createElement('div');
  name.className = 'ext-partner-name ext-truncate';
  name.title = summary.partnerName;
  name.textContent = summary.partnerName;
  widget.appendChild(name);

  // Data rows with native-style spacers between them
  widget.append(
    createWidgetRow('Celková investice u partnera', formatCZK(summary.totalInvestment), 'ext-partner-value--highlight'),
    createWidgetSpacer(),
    createWidgetRow('Celkové výnosy od partnera', formatCZK(summary.totalYields)),
    createWidgetSpacer(),
    createWidgetRow('Projektů s investicí', summary.investedCount + ' z ' + summary.totalCount),
  );

  // Diversification row
  if (summary.portfolioSize > 0) {
    const concentration = summary.totalInvestment / summary.portfolioSize;
    const diversificationClass = concentration < DIVERSIFICATION_SAFE_THRESHOLD
      ? 'ext-partner-value--safe'
      : concentration < DIVERSIFICATION_WARNING_THRESHOLD
        ? 'ext-partner-value--warning'
        : 'ext-partner-value--danger';
    widget.append(
      createWidgetSpacer(),
      createWidgetRow('Podíl v portfoliu', (concentration * 100).toFixed(1) + ' %', diversificationClass),
    );
  }

  // Repayment breakdown
  const repaymentRow = createRepaymentRow(summary.repaymentBreakdown);
  if (repaymentRow) {
    widget.append(createWidgetSpacer(), repaymentRow);
  }

  return widget;
}

/**
 * Creates a badge element showing the total investment at a partner.
 * Displayed inside the yield column of each listing card.
 * @param {Object} summary - The computed partner summary.
 * @returns {HTMLDivElement} The badge element ready for DOM insertion.
 */
function renderBadge(summary) {
  const el = document.createElement('div');
  el.className = 'ext-partner-badge';
  el.setAttribute(BADGE_ATTR, '');

  const label = document.createElement('span');
  label.className = 'ext-badge-label';
  label.textContent = 'U partnera';

  const value = document.createElement('span');
  value.className = 'ext-badge-value';
  value.textContent = formatCZK(summary.totalInvestment);

  el.append(label, value);
  return el;
}

/**
 * Creates a "new partner" badge for cards with no existing investment at the partner.
 * @returns {HTMLDivElement} The badge element ready for DOM insertion.
 */
function renderNewPartnerBadge() {
  const el = document.createElement('div');
  el.className = 'ext-partner-badge--new';
  el.setAttribute(BADGE_ATTR, '');
  el.textContent = 'Nový partner';
  return el;
}

// ===========================================================================
// DOM Injection
// ===========================================================================

/**
 * Checks whether a DOM element is positioned within the right sidebar.
 * Uses two strategies: (1) walks up ancestors looking for layout container classes,
 * and (2) checks if the element is in the right half of the viewport.
 * @param {HTMLElement} el - The element to check.
 * @returns {boolean} True if the element is in the sidebar.
 */
function isInSidebar(el) {
  // Strategy 1: walk up ancestors looking for sidebar layout classes
  let parent = el;
  while (parent && parent !== document.body) {
    const cls = parent.className || '';
    if (typeof cls === 'string' && /layout[-_]?container/i.test(cls)) return true;
    parent = parent.parentElement;
  }
  // Strategy 2: element is in the right half of the viewport (sidebar position)
  const rect = el.getBoundingClientRect();
  return rect.width > 0 && rect.left > window.innerWidth / 2;
}

/**
 * Finds DOM elements whose direct text content matches the given string,
 * using a TreeWalker for efficient traversal instead of querySelectorAll('span, div').
 * @param {Node} root - The root node to search within.
 * @param {string} text - The exact text to match (trimmed).
 * @returns {HTMLElement[]} Array of parent elements whose text nodes match.
 */
function findElementsByText(root, text) {
  const results = [];
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      return node.textContent.trim() === text ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP;
    },
  });
  while (walker.nextNode()) {
    const parent = walker.currentNode.parentElement;
    if (parent) results.push(parent);
  }
  return results;
}

/**
 * Locates the best DOM position for injecting the partner summary widget.
 * Tries two strategies:
 *   1. Finds the "Investice" section heading in the right sidebar
 *   2. Falls back to the invest button's parent section
 * @returns {{ target: HTMLElement, position: string }|null} Injection point descriptor,
 *   or null if no suitable position is found.
 */
function findInjectionPoint() {
  // Strategy 1: Find "Investice" section heading in the right sidebar
  // Note: page may have multiple .layout-container-content elements; search all of them
  const sidebars = document.querySelectorAll('.layout-container-content');
  const searchRoots = sidebars.length > 0 ? [...sidebars] : [document.body];
  const headings = searchRoots.flatMap(root => findElementsByText(root, 'Investice'));
  // Prefer the one actually in the sidebar; fall back to first match
  const heading = headings.find(el => isInSidebar(el)) || headings[0];
  if (heading?.parentElement) {
    return { target: heading.parentElement, position: 'beforebegin' };
  }

  // Strategy 2: Find invest button by data-testid — widget goes after its section
  const investBtn = document.querySelector('[data-testid="invest"]');
  if (investBtn?.parentElement) {
    return { target: investBtn.parentElement, position: 'afterend' };
  }

  return null;
}

/**
 * Removes the partner summary widget from the DOM.
 */
function removeWidget() {
  document.getElementById(WIDGET_ID)?.remove();
}

/**
 * Inserts the partner summary widget element into the page.
 * Uses {@link findInjectionPoint} to locate the correct position.
 * Retries with linear backoff (base {@link WIDGET_INSERTION_RETRY_BASE_MS})
 * up to {@link MAX_WIDGET_INSERTION_RETRIES} times if the injection point isn't available yet.
 * @param {HTMLElement} element - The widget element to insert.
 */
function insertWidget(element) {
  let attempts = 0;
  const tryInsert = () => {
    const point = findInjectionPoint();
    if (point) {
      removeWidget();
      point.target.insertAdjacentElement(point.position, element);
    } else if (attempts < MAX_WIDGET_INSERTION_RETRIES) {
      attempts++;
      setTimeout(tryInsert, WIDGET_INSERTION_RETRY_BASE_MS * attempts);
    }
  };
  tryInsert();
}

/**
 * Injects a partner badge into the yield column of a listing card element.
 * Shows the investment amount or a "new partner" indicator.
 * @param {HTMLElement} cardElement - The listing card (anchor) element.
 * @param {Object|null} summary - The computed partner summary, or null.
 */
function injectBadge(cardElement, summary) {
  if (cardElement.querySelector(`[${BADGE_ATTR}]`)) return;
  const yieldCol = cardElement.children[1];
  if (!yieldCol) return;

  // null summary = failed fetch — don't show misleading badge
  if (!summary) return;

  // Switch yield column to vertical stack so badge appears below yield text
  yieldCol.classList.add('ext-yield-col--stacked');

  if (summary.totalInvestment === 0) {
    yieldCol.appendChild(renderNewPartnerBadge());
  } else {
    yieldCol.appendChild(renderBadge(summary));
  }
}

/**
 * Removes all injected listing badges and restores yield column styles.
 */
function removeAllBadges() {
  document.querySelectorAll(`[${BADGE_ATTR}]`).forEach(el => {
    const parent = el.parentElement;
    el.remove();
    if (parent) parent.classList.remove('ext-yield-col--stacked');
  });
}

/**
 * Re-injects the partner summary widget if it was removed (e.g. by React DOM reconciliation).
 * Uses cached data when available; otherwise triggers a fresh fetch.
 * @param {string} slug - The property slug to ensure the widget is present for.
 */
function ensureWidgetPresent(slug) {
  if (document.getElementById(WIDGET_ID)) return;
  const cached = getCached(slug);
  if (cached) {
    insertWidget(renderWidget({ summary: cached }));
  } else {
    injectPartnerSummary(slug);
  }
}

// ===========================================================================
// Page Handler — Property Detail
// ===========================================================================

/**
 * Orchestrates the partner summary widget injection on a property detail page.
 * Three phases:
 *   1. Authenticates the user via JWT from localStorage
 *   2. Shows cached data instantly if available (stale-while-revalidate)
 *   3. Fetches fresh data in the background and updates the widget
 * Guards against stale results when the user navigates away during fetch.
 * @param {string} slug - The property slug to display partner data for.
 */
async function injectPartnerSummary(slug) {
  if (pendingFetches.has(slug)) return;
  pendingFetches.add(slug);
  try {
    removeWidget();

    // Phase 1: Authentication
    const token = getAuthToken();
    if (!token) {
      insertWidget(renderWidget({ error: 'auth' }));
      return;
    }

    // Phase 2: Show cached data instantly
    const cached = getCached(slug);
    if (cached) {
      insertWidget(renderWidget({ summary: cached }));
    }

    // Phase 3: Fetch fresh data in background
    try {
      const [property, relatedProperties, portfolioSize] = await Promise.all([
        fetchProperty(slug, token),
        fetchAllRelatedProperties(slug, token),
        fetchPortfolioStatistics(token),
      ]);

      // Guard: user navigated away during fetch — discard stale result
      if (currentSlug !== slug) return;

      const summary = { ...computeSummary(property, relatedProperties), portfolioSize };
      summaryCache.set(slug, summary);

      // Update widget (replaces cached version or inserts for first time)
      removeWidget();
      insertWidget(renderWidget({ summary }));
    } catch (err) {
      // Guard: don't show error for stale request
      if (currentSlug !== slug) return;

      // If we already showed cached data, keep it visible on refresh failure
      if (cached) return;

      console.error('[Investown Extension]', err);
      removeWidget();
      insertWidget(renderWidget({ error: 'fetch' }));
    }
  } finally {
    pendingFetches.delete(slug);
  }
}

// ===========================================================================
// Page Handler — Listing
// ===========================================================================

/**
 * Waits until listing cards are present in the DOM, then calls the callback.
 * If cards are already present, calls the callback synchronously.
 * Uses a MutationObserver on document.body to detect when React renders the listing.
 * @param {function} callback - The function to call once listing cards are ready.
 */
function waitForListingCards(callback) {
  if (document.querySelectorAll(PROPERTY_LINK_SELECTOR).length > 0) {
    callback();
    return;
  }
  const observer = new MutationObserver(() => {
    if (document.querySelectorAll(PROPERTY_LINK_SELECTOR).length > 0) {
      observer.disconnect();
      clearTimeout(timeout);
      callback();
    }
  });
  observer.observe(document.body, { childList: true, subtree: true });
  const timeout = setTimeout(() => {
    observer.disconnect();
    console.warn('[Investown Extension] waitForListingCards: timed out after 30 s');
  }, 30_000);
}

/**
 * Extracts property slugs and their corresponding link elements from the listing page.
 * Finds all anchor elements pointing to property detail pages.
 * @returns {{ link: HTMLAnchorElement, slug: string }[]} Array of link-slug pairs.
 */
function extractSlugsFromCards() {
  const links = document.querySelectorAll(PROPERTY_LINK_SELECTOR);
  const results = [];
  for (const link of links) {
    const slug = getSlugFromPath(new URL(link.href).pathname);
    if (slug) results.push({ link, slug });
  }
  return results;
}

/**
 * Processes the listing page: fetches partner data for each card and injects
 * inline badges into the first column. Uses the concurrency limiter to control
 * API request parallelism. A navigation ID guards against stale results when
 * the user leaves the listing page.
 */
async function processListingPage() {
  const capturedNavigationId = ++listingNavigationId;

  const token = getAuthToken();
  if (!token) return;

  const cards = extractSlugsFromCards();
  if (cards.length === 0) return;

  const tasks = cards.map(({ link, slug }) => {
    return apiLimiter(async () => {
      if (listingNavigationId !== capturedNavigationId) return;

      try {
        const summary = await fetchPartnerDataForSlug(slug, token);
        if (listingNavigationId !== capturedNavigationId) return;

        if (!link) return;
        injectBadge(link, summary);
      } catch (err) {
        console.error('[Investown Extension] fetchPartnerData failed for', slug, err);
      }
    });
  });

  await Promise.allSettled(tasks);
}

/**
 * Sets up a MutationObserver on the listing page to detect dynamically loaded cards
 * (e.g. from infinite scroll). When new cards without badges are detected,
 * triggers a re-processing of the listing page.
 * Debounced by {@link LISTING_MUTATION_DEBOUNCE_MS} to avoid excessive re-processing.
 */
function observeListingMutations() {
  if (listingObserver) listingObserver.disconnect();
  listingObserver = new MutationObserver(() => {
    if (currentPageType !== 'listing') return;
    clearTimeout(listingDebounceTimer);
    listingDebounceTimer = setTimeout(() => {
      const cards = document.querySelectorAll(PROPERTY_LINK_SELECTOR);
      const hasMissing = [...cards].some(link => {
        return link && !link.querySelector(`[${BADGE_ATTR}]`);
      });
      if (hasMissing) processListingPage();
    }, LISTING_MUTATION_DEBOUNCE_MS);
  });
  listingObserver.observe(document.body, { childList: true, subtree: true });
}

// ===========================================================================
// Route Management
// ===========================================================================

/**
 * Determines the page type from a URL path.
 * @param {string} path - The URL pathname.
 * @returns {'property'|'listing'|'other'} The detected page type.
 */
function getPageType(path) {
  if (/^\/property\/[^/#?]+/.test(path)) return 'property';
  if (path === '/' || path === '') return 'listing';
  return 'other';
}

/**
 * Extracts the property slug from a URL path.
 * @param {string} path - The URL pathname (e.g. "/property/some-slug").
 * @returns {string|null} The slug, or null if the path doesn't match a property URL.
 */
function getSlugFromPath(path) {
  const match = path.match(/^\/property\/([^/#?]+)/);
  return match ? match[1] : null;
}

/**
 * Handles SPA route changes. Detects page type transitions, tears down previous page
 * state (widgets, badges, observers), and initializes the new page type.
 * Called on pushState, replaceState, popstate, and hashchange events.
 */
function handleRouteChange() {
  const pageType = getPageType(location.pathname);
  const newSlug = getSlugFromPath(location.pathname);
  // --- Cleanup: tear down previous page state ---
  if (currentPageType === 'property' && pageType !== 'property') {
    currentSlug = null;
    removeWidget();
    disconnectWidgetObserver();
  }
  if (currentPageType === 'listing' && pageType !== 'listing') {
    removeAllBadges();
    if (listingObserver) { listingObserver.disconnect(); listingObserver = null; }
  }

  // --- Initialize: set up new page state ---
  currentPageType = pageType;

  if (pageType === 'property') {
    connectWidgetObserver();
    if (newSlug && newSlug !== currentSlug) {
      currentSlug = newSlug;
      injectPartnerSummary(newSlug);
    } else if (newSlug && newSlug === currentSlug) {
      ensureWidgetPresent(newSlug);
    }
  } else if (pageType === 'listing') {
    currentSlug = null;
    waitForListingCards(async () => {
      if (currentPageType !== 'listing') return;
      await processListingPage();
      if (currentPageType !== 'listing') return;
      observeListingMutations();
    });
  } else {
    currentSlug = null;
  }
}

// SPA History API interception
const origPushState = history.pushState;
history.pushState = function (...args) {
  origPushState.apply(this, args);
  handleRouteChange();
};

const origReplaceState = history.replaceState;
history.replaceState = function (...args) {
  origReplaceState.apply(this, args);
  handleRouteChange();
};

window.addEventListener('popstate', handleRouteChange);

window.addEventListener('hashchange', () => {
  if (currentSlug) ensureWidgetPresent(currentSlug);
});

/**
 * Connects the widget observer to watch for React DOM reconciliation removing the widget.
 * Only needed on property detail pages.
 */
function connectWidgetObserver() {
  widgetObserver.observe(document.body, { childList: true, subtree: true });
}

/**
 * Disconnects the widget observer to stop watching for widget removal.
 */
function disconnectWidgetObserver() {
  widgetObserver.disconnect();
  clearTimeout(reinsertTimer);
}

// ===========================================================================
// Bootstrap — must be after all declarations to avoid TDZ
// ===========================================================================

/**
 * Waits for the page content to be ready before executing a callback.
 * For property pages, waits until the "Investice" heading is present in the sidebar.
 * For listing pages, waits until property links are present.
 * Uses a MutationObserver to detect readiness if the page isn't ready yet.
 * @param {function} callback - The function to call once the page is ready.
 */
function onPageReady(callback) {
  const pageType = getPageType(location.pathname);

  if (pageType === 'listing') {
    waitForListingCards(callback);
    return;
  }

  const isReady = () => {
    if (pageType === 'property') {
      return findElementsByText(document.body, 'Investice').some(el => isInSidebar(el));
    }
    return true;
  };
  if (isReady()) { callback(); return; }
  const readinessObserver = new MutationObserver(() => {
    if (isReady()) { readinessObserver.disconnect(); clearTimeout(readinessTimeout); callback(); }
  });
  readinessObserver.observe(document.body, { childList: true, subtree: true });
  const readinessTimeout = setTimeout(() => readinessObserver.disconnect(), 30_000);
}

onPageReady(() => handleRouteChange());
