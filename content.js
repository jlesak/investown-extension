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
const MIN_COLUMNS_FOR_HEADER_ROW = 6;
const FIRST_COLUMN_CONSTRAINED_WIDTH = '430px';

const BADGE_ATTR = 'data-ext-badge';
const HEADER_ATTR = 'data-ext-header';
const COL_WIDTH_ADJUSTED = 'data-ext-col-adjusted';

// ===========================================================================
// Safe API References (MAIN world protection against page overrides)
// ===========================================================================

const safeFetch = fetch.bind(window);
const safeJsonParse = JSON.parse;
const safeJsonStringify = JSON.stringify;
const safeAtob = atob;

// ===========================================================================
// State Variables
// ===========================================================================

// Current navigation state
let currentSlug = null;
let currentPageType = null;
let listingNavigationId = 0; // incremented on each listing page entry, used as stale guard

// Caches
const summaryCache = new Map();      // slug → { data: Summary, timestamp: number }
const propertyMetaCache = new Map(); // slug → Property (GraphQL response object)
const partnerCache = new Map();      // companyIdentifier → Summary

// In-flight request tracking
const pendingFetches = new Set();    // slugs currently being fetched

// Observers and timers
let listingObserver = null;
let listingDebounceTimer = null;
let reinsertTimer = null;

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
 * Escapes a string for safe insertion into HTML.
 * Uses a temporary DOM element to leverage the browser's built-in escaping.
 * @param {string} str - The raw string to escape.
 * @returns {string} HTML-safe string with entities escaped.
 */
function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

/**
 * Retrieves a cached summary for a property slug if it exists and hasn't expired.
 * @param {string} slug - The property slug to look up.
 * @returns {Object|null} The cached summary data, or null if missing or expired.
 */
function getCached(slug) {
  const entry = summaryCache.get(slug);
  if (!entry) return null;
  if (Date.now() - entry.timestamp > CACHE_TTL_MS) {
    summaryCache.delete(slug);
    return null;
  }
  return entry.data;
}

/**
 * Inserts an element as the second child of a parent container.
 * If the parent has at least one child, the element is placed immediately after it;
 * otherwise, the element is simply appended.
 * @param {HTMLElement} parent - The parent container.
 * @param {HTMLElement} element - The element to insert.
 */
function insertAsSecondChild(parent, element) {
  const firstChild = parent.children[0];
  if (firstChild?.nextSibling) {
    parent.insertBefore(element, firstChild.nextSibling);
  } else {
    parent.appendChild(element);
  }
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
  const keys = Object.keys(localStorage);
  const idTokenKey = keys.find(key => key.endsWith('.idToken'));
  if (!idTokenKey) return null;

  const token = localStorage.getItem(idTokenKey);
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
          borrower {
            companyName
            companyIdentifier
          }
        }
      }
    }
  `;
  const data = await executeGraphQLQuery(query, { slug }, token);
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
  const property = await fetchProperty(slug, token);
  propertyMetaCache.set(slug, property);
  return property;
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

  const relatedProperties = await fetchAllRelatedProperties(slug, token);
  const summary = computeSummary(property, relatedProperties);
  partnerCache.set(companyId, summary);
  return summary;
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
    property => property.interestAmount?.currency === 'CZK' || !property.interestAmount?.currency
  );

  const relatedInvestment = czkRelated.reduce(
    (sum, property) => sum + (property.investmentAmount || 0), 0
  );

  const totalYields = relatedProperties
    .filter(property => property.interestAmount?.currency === 'CZK')
    .reduce((sum, property) => sum + (property.interestAmount?.value || 0), 0);

  const investedCount =
    czkRelated.filter(property => property.investmentAmount > 0).length +
    (currentInvestment > 0 ? 1 : 0);

  const totalCount = relatedProperties.length + 1;

  const totalInvestment = currentInvestment + relatedInvestment;
  return {
    partnerName: borrower?.companyName || 'Neznámý partner',
    totalInvestment,
    totalYields,
    investedCount,
    totalCount,
  };
}

// ===========================================================================
// DOM Rendering
// ===========================================================================

/**
 * Renders the partner summary widget HTML for the property detail page.
 * Produces either an error message or a full summary with investment totals.
 * @param {Object} options - Render options.
 * @param {Object} [options.summary] - The computed partner summary to display.
 * @param {string} [options.error] - Error type: 'auth' for unauthenticated, 'fetch' for API failure.
 * @returns {string} HTML string ready for insertion into the DOM.
 */
function renderWidget({ summary, error }) {
  if (error) {
    const messages = {
      auth: 'Nepřihlášen — nelze načíst data',
      fetch: 'Chyba při načítání dat partnera',
    };
    return `
      <div id="${WIDGET_ID}">
        <div class="ext-partner-error">${messages[error] || 'Neočekávaná chyba'}</div>
      </div>
    `;
  }

  return `
    <div id="${WIDGET_ID}">
      <div class="ext-partner-title ext-truncate" title="${escapeHtml(summary.partnerName)}">
        ${escapeHtml(summary.partnerName)}
      </div>
      <div class="ext-partner-row">
        <span class="ext-partner-label">Celková investice u partnera</span>
        <span class="ext-partner-value ext-partner-value--highlight">
          ${formatCZK(summary.totalInvestment)}
        </span>
      </div>
      <div class="ext-partner-divider"></div>
      <div class="ext-partner-row">
        <span class="ext-partner-label">Celkové výnosy od partnera</span>
        <span class="ext-partner-value">${formatCZK(summary.totalYields)}</span>
      </div>
      <div class="ext-partner-divider"></div>
      <div class="ext-partner-row">
        <span class="ext-partner-label">Projektů s investicí</span>
        <span class="ext-partner-value">
          ${summary.investedCount} z ${summary.totalCount}
        </span>
      </div>
    </div>
  `;
}

/**
 * Creates a badge element for the listing page showing partner data for a single card row.
 * Displays partner name, total investment amount, and project count.
 * @param {Object} summary - The computed partner summary.
 * @returns {HTMLDivElement} The badge element ready for DOM insertion.
 */
function renderBadge(summary) {
  const el = document.createElement('div');
  el.className = 'ext-partner-col';
  el.setAttribute(BADGE_ATTR, '');

  const name = document.createElement('span');
  name.className = 'ext-col-name ext-truncate';
  name.title = summary.partnerName;
  name.textContent = summary.partnerName;

  const investment = document.createElement('span');
  investment.className = 'ext-col-investment';
  investment.textContent = formatCZK(summary.totalInvestment);

  const projects = document.createElement('span');
  projects.className = 'ext-col-projects';
  projects.textContent = summary.investedCount + ' z ' + summary.totalCount + ' proj.';

  el.append(name, investment, projects);
  return el;
}

/**
 * Creates an empty placeholder badge column for the listing page.
 * Used as a layout placeholder while data is loading, or for properties with no investment.
 * @returns {HTMLDivElement} An empty badge element with the badge attribute.
 */
function renderEmptyColumn() {
  const el = document.createElement('div');
  el.className = 'ext-partner-col';
  el.setAttribute(BADGE_ATTR, '');
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
 * Locates the best DOM position for injecting the partner summary widget.
 * Tries two strategies:
 *   1. Finds the "Investice" section heading in the right sidebar
 *   2. Falls back to the invest button's parent section
 * @returns {{ target: HTMLElement, position: string }|null} Injection point descriptor,
 *   or null if no suitable position is found.
 */
function findInjectionPoint() {
  // Strategy 1: Find "Investice" section heading in the right sidebar
  const sidebar = document.querySelector('.layout-container-content');
  const searchRoot = sidebar || document;
  const headings = [...searchRoot.querySelectorAll('span, div')].filter(
    el => el.textContent.trim() === 'Investice'
  );
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
 * Inserts the partner summary widget HTML into the page.
 * Uses {@link findInjectionPoint} to locate the correct position.
 * Retries with exponential backoff (base {@link WIDGET_INSERTION_RETRY_BASE_MS})
 * up to {@link MAX_WIDGET_INSERTION_RETRIES} times if the injection point isn't available yet.
 * @param {string} html - The widget HTML string to insert.
 */
function insertWidget(html) {
  let attempts = 0;
  const tryInsert = () => {
    const point = findInjectionPoint();
    if (point) {
      removeWidget();
      point.target.insertAdjacentHTML(point.position, html);
    } else if (attempts < MAX_WIDGET_INSERTION_RETRIES) {
      attempts++;
      setTimeout(tryInsert, WIDGET_INSERTION_RETRY_BASE_MS * attempts);
    }
  };
  tryInsert();
}

/**
 * Constrains the first column of a listing row to a fixed width, making room
 * for the injected partner column. Marks the row with {@link COL_WIDTH_ADJUSTED}
 * to prevent duplicate adjustment.
 * @param {HTMLElement} row - The listing row element to adjust.
 */
function shrinkFirstColumn(row) {
  if (row.getAttribute(COL_WIDTH_ADJUSTED)) return;
  const firstCol = row.children[0];
  if (firstCol) {
    firstCol.style.minWidth = FIRST_COLUMN_CONSTRAINED_WIDTH;
    firstCol.style.maxWidth = FIRST_COLUMN_CONSTRAINED_WIDTH;
    firstCol.style.overflow = 'hidden';
  }
  row.setAttribute(COL_WIDTH_ADJUSTED, '');
}

/**
 * Injects the "Partner" header column into the listing page header row.
 * Finds the header by locating the "Výnos" sort button and walking up to
 * the parent row with {@link MIN_COLUMNS_FOR_HEADER_ROW}+ column children.
 */
function injectHeaderColumn() {
  if (document.querySelector(`[${HEADER_ATTR}]`)) return;

  // Find the "Výnos" sort button, then walk up to the header row
  const allButtons = document.querySelectorAll('.layout-container-content button');
  const vynosBtn = [...allButtons].find(button => button.textContent?.trim().startsWith('Výnos'));
  if (!vynosBtn) return;

  // Walk up from the button to find the header row (the flex parent with enough column children)
  let headerRow = vynosBtn.parentElement;
  while (headerRow && headerRow.children.length < MIN_COLUMNS_FOR_HEADER_ROW) {
    headerRow = headerRow.parentElement;
  }
  if (!headerRow) return;

  shrinkFirstColumn(headerRow);

  const header = document.createElement('div');
  header.className = 'ext-partner-col-header';
  header.setAttribute(HEADER_ATTR, '');

  const label = document.createElement('span');
  label.className = 'ext-col-header-label';
  label.textContent = 'Partner';
  header.appendChild(label);

  insertAsSecondChild(headerRow, header);
}

/**
 * Injects a partner badge into a listing card element.
 * Shows the partner summary if the total investment is positive,
 * otherwise shows an empty placeholder column.
 * @param {HTMLElement} cardElement - The listing card (anchor) element.
 * @param {Object|null} summary - The computed partner summary, or null.
 */
function injectBadge(cardElement, summary) {
  if (cardElement.querySelector(`[${BADGE_ATTR}]`)) return;

  shrinkFirstColumn(cardElement);

  const col = (summary && summary.totalInvestment > 0)
    ? renderBadge(summary)
    : renderEmptyColumn();

  insertAsSecondChild(cardElement, col);
}

/**
 * Removes all injected listing badges, header columns, and width adjustments.
 * Restores the original listing layout.
 */
function removeAllBadges() {
  document.querySelectorAll(`[${BADGE_ATTR}]`).forEach(el => el.remove());
  document.querySelectorAll(`[${HEADER_ATTR}]`).forEach(el => el.remove());
  document.querySelectorAll(`[${COL_WIDTH_ADJUSTED}]`).forEach(row => {
    const firstCol = row.children[0];
    if (firstCol) {
      firstCol.style.minWidth = '';
      firstCol.style.maxWidth = '';
      firstCol.style.overflow = '';
    }
    row.removeAttribute(COL_WIDTH_ADJUSTED);
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
      const [property, relatedProperties] = await Promise.all([
        fetchProperty(slug, token),
        fetchAllRelatedProperties(slug, token),
      ]);

      // Guard: user navigated away during fetch — discard stale result
      if (currentSlug !== slug) return;

      const summary = computeSummary(property, relatedProperties);
      summaryCache.set(slug, { data: summary, timestamp: Date.now() });

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
 * Extracts property slugs and their corresponding link elements from the listing page.
 * Finds all anchor elements pointing to property detail pages.
 * @returns {{ link: HTMLAnchorElement, slug: string }[]} Array of link-slug pairs.
 */
function extractSlugsFromCards() {
  const links = document.querySelectorAll('a[href^="/property/"]');
  const results = [];
  for (const link of links) {
    const slug = getSlugFromPath(new URL(link.href).pathname);
    if (slug) results.push({ link, slug });
  }
  return results;
}

/**
 * Processes the listing page: injects the "Partner" header column, adds placeholder
 * badges to all visible cards, then fetches partner data and replaces placeholders
 * with real badges. Uses the concurrency limiter to control API request parallelism.
 * A navigation ID guards against stale results when the user leaves the listing page.
 */
async function processListingPage() {
  const capturedNavigationId = ++listingNavigationId;

  const token = getAuthToken();
  if (!token) return;

  const cards = extractSlugsFromCards();
  if (cards.length === 0) return;

  // Phase 1: Prepare layout (header + placeholders)
  injectHeaderColumn();

  for (const { link } of cards) {
    if (link && !link.querySelector(`[${BADGE_ATTR}]`)) {
      shrinkFirstColumn(link);
      const placeholder = renderEmptyColumn();
      insertAsSecondChild(link, placeholder);
    }
  }

  // Phase 2: Fetch data and inject badges
  const tasks = cards.map(({ link, slug }) => {
    return apiLimiter(async () => {
      // Stale navigation guard
      if (listingNavigationId !== capturedNavigationId) return;

      const summary = await fetchPartnerDataForSlug(slug, token);
      if (listingNavigationId !== capturedNavigationId) return;

      if (!link) return;

      // Replace placeholder with actual data (or keep empty if no investment)
      const existing = link.querySelector(`[${BADGE_ATTR}]`);
      if (existing) existing.remove();
      injectBadge(link, summary);
    });
  });

  await Promise.allSettled(tasks);
}

/**
 * Sets up a MutationObserver on the listing page to detect dynamically loaded cards
 * (e.g. from infinite scroll). When new cards without badges are detected or the
 * header column is missing, triggers a re-processing of the listing page.
 * Debounced by {@link LISTING_MUTATION_DEBOUNCE_MS} to avoid excessive re-processing.
 */
function observeListingMutations() {
  if (listingObserver) listingObserver.disconnect();
  listingObserver = new MutationObserver(() => {
    if (currentPageType !== 'listing') return;
    clearTimeout(listingDebounceTimer);
    listingDebounceTimer = setTimeout(() => {
      const cards = document.querySelectorAll('a[href^="/property/"]');
      const hasMissing = [...cards].some(link => {
        return link && !link.querySelector(`[${BADGE_ATTR}]`);
      });
      const headerMissing = !document.querySelector(`[${HEADER_ATTR}]`);
      if (hasMissing || headerMissing) processListingPage();
    }, LISTING_MUTATION_DEBOUNCE_MS);
  });
  const content = document.querySelector('.layout-container-content') || document.body;
  listingObserver.observe(content, { childList: true, subtree: true });
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
  }
  if (currentPageType === 'listing' && pageType !== 'listing') {
    removeAllBadges();
    if (listingObserver) { listingObserver.disconnect(); listingObserver = null; }
  }

  // --- Initialize: set up new page state ---
  currentPageType = pageType;

  if (pageType === 'property') {
    if (newSlug && newSlug !== currentSlug) {
      currentSlug = newSlug;
      injectPartnerSummary(newSlug);
    } else if (newSlug && newSlug === currentSlug) {
      ensureWidgetPresent(newSlug);
    }
  } else if (pageType === 'listing') {
    currentSlug = null;
    processListingPage();
    observeListingMutations();
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

// Re-inject widget if React removes it during DOM reconciliation (debounced)
new MutationObserver(() => {
  if (currentPageType === 'property' && currentSlug && !document.getElementById(WIDGET_ID)) {
    clearTimeout(reinsertTimer);
    reinsertTimer = setTimeout(() => ensureWidgetPresent(currentSlug), WIDGET_REINSERTION_DEBOUNCE_MS);
  }
}).observe(document.body, { childList: true, subtree: true });

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
  const isReady = () => {
    const pageType = getPageType(location.pathname);
    if (pageType === 'property') {
      return [...document.querySelectorAll('span, div')].some(
        el => el.textContent.trim() === 'Investice' && isInSidebar(el)
      );
    }
    if (pageType === 'listing') {
      return document.querySelectorAll('a[href^="/property/"]').length > 0;
    }
    return true;
  };
  if (isReady()) { callback(); return; }
  const readinessObserver = new MutationObserver(() => {
    if (isReady()) { readinessObserver.disconnect(); callback(); }
  });
  readinessObserver.observe(document.body, { childList: true, subtree: true });
}

onPageReady(() => handleRouteChange());
