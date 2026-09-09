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

  // Reports whether a cleaned file name corresponds to a real entity, so a
  // bad guess surfaces before a long run rather than partway through it.
  // Comparison ignores case and spacing, since OData spells an entity
  // "CustomerGroups" where the import form labels it "Customer groups".
  function validate(cleanedName) {
    if (names.length === 0) return { status: 'unknown' };

    const { candidate, score } = D365IA.matcher.bestMatch(cleanedName, names);
    if (score >= 1) return { status: 'match', name: candidate };
    if (score >= 0.7) return { status: 'close', name: candidate };
    return { status: 'none', name: candidate };
  }

  D365IA.entityList = { load, refresh, validate, count: () => names.length };
})();
