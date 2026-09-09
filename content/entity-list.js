(function () {
  const D365IA = (window.D365IA = window.D365IA || {});
  const STORAGE_KEY = 'entityList';

  let names = [];

  async function load() {
    const data = await chrome.storage.local.get(STORAGE_KEY);
    const stored = data[STORAGE_KEY];
    if (stored && Array.isArray(stored.names)) names = stored.names;
    return { count: names.length, fetchedAt: stored ? stored.fetchedAt : null };
  }

  // The environment's own OData service document lists every entity exposed
  // by this tenant, so the list is always right here rather than a hardcoded
  // one that drifts with version and customizations. It's same-origin, so
  // the session already open in the tab authenticates the request.
  async function refresh() {
    const response = await fetch(`${location.origin}/data`, {
      credentials: 'include',
      headers: { Accept: 'application/json' }
    });
    if (!response.ok) {
      throw new Error(`/data returned ${response.status}`);
    }

    const body = await response.json();
    const entries = Array.isArray(body.value) ? body.value : [];
    const fetched = entries.map((entry) => entry.name).filter(Boolean);
    if (fetched.length === 0) throw new Error('/data listed no entities');

    names = fetched.sort();
    const fetchedAt = Date.now();
    // Thousands of names — too big for storage.sync's per-item quota.
    await chrome.storage.local.set({ [STORAGE_KEY]: { names, fetchedAt } });
    return { count: names.length, fetchedAt };
  }

  // Fuzzy-matching a name against several thousand entities costs real time
  // (measured ~1.2s for a 37-file queue), and the panel re-renders on every
  // status change. Results only depend on the name and the loaded list, so
  // they're cached until the list is replaced — otherwise this blocks the
  // same main thread D365 renders on, which itself causes timing failures.
  let cache = new Map();

  function clearCache() {
    cache = new Map();
  }

  // Reports whether a cleaned file name corresponds to a real entity. This is
  // a hint only: OData lists technical names ("OperationalSitesV2") while the
  // import form shows display labels ("Sites V2"), which are often
  // legitimately different strings for the same entity.
  function validate(cleanedName) {
    if (names.length === 0) return { status: 'unknown' };

    const key = String(cleanedName || '');
    if (cache.has(key)) return cache.get(key);

    const { candidate, score } = D365IA.matcher.bestMatch(key, names);
    let result;
    if (score >= 1) result = { status: 'match', name: candidate };
    else if (score >= 0.7) result = { status: 'close', name: candidate };
    else result = { status: 'none', name: candidate };

    cache.set(key, result);
    return result;
  }

  D365IA.entityList = { load, refresh, validate, count: () => names.length };
})();
