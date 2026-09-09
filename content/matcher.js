(function () {
  const D365IA = (window.D365IA = window.D365IA || {});

  const DEFAULT_RULES = {
    stripExtension: true,
    stripLeadingNumbers: true,
    stripTrailingNumbers: true,
    stripDates: true,
    stripVersionSuffixes: false,
    separatorsToSpaces: true,
    collapseWhitespace: true,
    titleCase: true
  };

  const EXTENSION_RE = /\.(xlsx|xlsm|xls|csv|txt)$/i;

  const DATE_TOKEN_RES = [
    /^\d{4}[-_.]?\d{2}[-_.]?\d{2}$/, // yyyymmdd / yyyy-mm-dd
    /^\d{2}[-_.]?\d{2}[-_.]?\d{4}$/, // ddmmyyyy / mm-dd-yyyy
    /^\d{6}$/, // yymmdd or hhmmss
    /^\d{8}$/, // yyyymmdd, no separators
    /^\d{1,2}[-_.]\d{1,2}[-_.]\d{2,4}$/ // d-m-yy etc.
  ];

  const VERSION_TOKEN_RES = [/^v\d+$/i, /^\(\d+\)$/, /^final$/i, /^copy$/i, /^draft$/i];

  // Leading classification codes like "01", "02B.03SYS" or "80.ALOG.WM".
  // Anything starting with a digit at the front of a file name is a
  // numbering/classification prefix, never part of an entity name.
  const CODE_PREFIX_RE = /^\d[\w.]*$/;

  function isPureNumberToken(t) {
    return /^\d+$/.test(t);
  }

  function isCodePrefixToken(t) {
    return CODE_PREFIX_RE.test(t);
  }

  function isDateToken(t) {
    return DATE_TOKEN_RES.some((re) => re.test(t));
  }

  function isVersionToken(t) {
    return VERSION_TOKEN_RES.some((re) => re.test(t));
  }

  // Turns a raw Excel file name into a best-guess entity name by stripping
  // sequence numbers, dates and separators from the edges, leaving anything
  // embedded mid-word (e.g. "CustomersV3") untouched.
  function cleanFileName(rawName, rules) {
    rules = Object.assign({}, DEFAULT_RULES, rules || {});
    let name = rawName;

    if (rules.stripExtension) {
      name = name.replace(EXTENSION_RE, '');
    }

    let tokens = name.split(/[_\-\s]+/).filter(Boolean);

    while (
      tokens.length > 1 &&
      ((rules.stripLeadingNumbers && isCodePrefixToken(tokens[0])) ||
        (rules.stripDates && isDateToken(tokens[0])) ||
        (rules.stripVersionSuffixes && isVersionToken(tokens[0])))
    ) {
      tokens.shift();
    }

    while (
      tokens.length > 1 &&
      ((rules.stripTrailingNumbers && isPureNumberToken(tokens[tokens.length - 1])) ||
        (rules.stripDates && isDateToken(tokens[tokens.length - 1])) ||
        (rules.stripVersionSuffixes && isVersionToken(tokens[tokens.length - 1])))
    ) {
      tokens.pop();
    }

    let result = tokens.join(rules.separatorsToSpaces ? ' ' : '');

    if (rules.collapseWhitespace) {
      result = result.replace(/\s+/g, ' ').trim();
    }

    if (rules.titleCase) {
      result = toTitleCaseSmart(result);
    }

    return result;
  }

  // Capitalizes plain words but leaves camelCase/PascalCase tokens
  // (e.g. "CustomersV3") as the user wrote them.
  function toTitleCaseSmart(str) {
    return str
      .split(' ')
      .map((word) => {
        if (!word) return word;
        if (/[a-z]/.test(word) && /[A-Z]/.test(word)) return word;
        return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
      })
      .join(' ');
  }

  function normalize(s) {
    return s.toLowerCase().replace(/[^a-z0-9]+/g, '');
  }

  function levenshtein(a, b) {
    const m = a.length;
    const n = b.length;
    if (m === 0) return n;
    if (n === 0) return m;
    const dp = new Array(n + 1);
    for (let j = 0; j <= n; j++) dp[j] = j;
    for (let i = 1; i <= m; i++) {
      let prev = dp[0];
      dp[0] = i;
      for (let j = 1; j <= n; j++) {
        const temp = dp[j];
        dp[j] = Math.min(dp[j] + 1, dp[j - 1] + 1, prev + (a[i - 1] === b[j - 1] ? 0 : 1));
        prev = temp;
      }
    }
    return dp[n];
  }

  function scoreMatch(cleanedName, candidate) {
    const a = normalize(cleanedName);
    const b = normalize(candidate);
    if (!a || !b) return 0;
    if (a === b) return 1;
    if (b.startsWith(a) || a.startsWith(b)) return 0.9;
    if (b.includes(a) || a.includes(b)) return 0.75;
    const dist = levenshtein(a, b);
    const maxLen = Math.max(a.length, b.length);
    return Math.max(0, 1 - dist / maxLen) * 0.7;
  }

  function bestMatch(cleanedName, candidates) {
    let best = null;
    let bestScore = -1;
    for (const c of candidates) {
      const score = scoreMatch(cleanedName, c);
      if (score > bestScore) {
        bestScore = score;
        best = c;
      }
    }
    return { candidate: best, score: bestScore };
  }

  // Orders file names the way their numbering reads rather than the way
  // ASCII sorts, so "10_" follows "9_" and "04.12PMD-" precedes "80.ALOG.WM-".
  // Import order matters: these files depend on each other, and a batch run
  // in the wrong order fails.
  function naturalChunks(text) {
    return text.toLowerCase().match(/\d+|\D+/g) || [];
  }

  function compareNaturally(a, b) {
    const left = naturalChunks(a);
    const right = naturalChunks(b);

    for (let i = 0; i < Math.min(left.length, right.length); i++) {
      const x = left[i];
      const y = right[i];
      const bothNumeric = /^\d/.test(x) && /^\d/.test(y);

      if (bothNumeric) {
        const difference = parseInt(x, 10) - parseInt(y, 10);
        if (difference) return difference;
      } else if (x !== y) {
        return x < y ? -1 : 1;
      }
    }
    return left.length - right.length;
  }

  D365IA.matcher = {
    DEFAULT_RULES,
    cleanFileName,
    scoreMatch,
    bestMatch,
    compareNaturally
  };
})();
