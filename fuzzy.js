/**
 * FuzzySearch — Self-contained fuzzy search module for label matching.
 * Combines exact, prefix, substring, subsequence, and Levenshtein matching.
 * Returns scored results with match ranges for highlighting.
 */
var FuzzySearch = (function () {

  /**
   * Levenshtein distance between two strings.
   */
  function levenshtein(a, b) {
    if (a.length === 0) return b.length;
    if (b.length === 0) return a.length;

    var matrix = [];
    for (var i = 0; i <= b.length; i++) {
      matrix[i] = [i];
    }
    for (var j = 0; j <= a.length; j++) {
      matrix[0][j] = j;
    }
    for (var i = 1; i <= b.length; i++) {
      for (var j = 1; j <= a.length; j++) {
        if (b.charAt(i - 1) === a.charAt(j - 1)) {
          matrix[i][j] = matrix[i - 1][j - 1];
        } else {
          matrix[i][j] = Math.min(
            matrix[i - 1][j - 1] + 1,
            matrix[i][j - 1] + 1,
            matrix[i - 1][j] + 1
          );
        }
      }
    }
    return matrix[b.length][a.length];
  }

  /**
   * Check if all query chars appear in candidate in order (subsequence).
   * Returns { score, ranges } where ranges are [start, end] pairs.
   */
  function subsequenceMatch(query, candidate) {
    var qi = 0;
    var ranges = [];
    var candidateLower = candidate.toLowerCase();

    for (var ci = 0; ci < candidate.length && qi < query.length; ci++) {
      if (candidateLower[ci] === query[qi]) {
        ranges.push([ci, ci + 1]);
        qi++;
      }
    }

    if (qi === query.length) {
      var span = ranges[ranges.length - 1][1] - ranges[0][0];
      var density = query.length / span;
      return { score: 0.5 + (0.3 * density), ranges: ranges };
    }

    return { score: 0, ranges: [] };
  }

  /**
   * Merge overlapping or adjacent ranges and sort them.
   */
  function mergeRanges(ranges) {
    if (ranges.length === 0) return [];

    var sorted = ranges.slice().sort(function (a, b) { return a[0] - b[0]; });
    var merged = [sorted[0]];

    for (var i = 1; i < sorted.length; i++) {
      var last = merged[merged.length - 1];
      if (sorted[i][0] <= last[1]) {
        last[1] = Math.max(last[1], sorted[i][1]);
      } else {
        merged.push(sorted[i]);
      }
    }
    return merged;
  }

  /**
   * Score a candidate against a query.
   * Returns { score, matchType, matchRanges }
   *   score: 0 (no match) to 1 (perfect)
   *   matchType: 'exact' | 'prefix' | 'substring' | 'fuzzy' | 'none'
   *   matchRanges: [[start, end], ...] for highlighting
   */
  function score(query, candidate) {
    if (!query || !candidate) return { score: 0, matchType: 'none', matchRanges: [] };

    var q = query.toLowerCase().trim();
    var c = candidate.toLowerCase();

    if (!q) return { score: 0, matchType: 'none', matchRanges: [] };

    // Exact match
    if (q === c) {
      return { score: 1.0, matchType: 'exact', matchRanges: [[0, candidate.length]] };
    }

    // Prefix match
    if (c.startsWith(q)) {
      return {
        score: 0.9 + (0.1 * q.length / c.length),
        matchType: 'prefix',
        matchRanges: [[0, q.length]]
      };
    }

    // Substring match
    var subIdx = c.indexOf(q);
    if (subIdx !== -1) {
      return {
        score: 0.7 + (0.1 * q.length / c.length),
        matchType: 'substring',
        matchRanges: [[subIdx, subIdx + q.length]]
      };
    }

    // Word-start match: check if query matches the start of any word in the candidate
    var words = c.split(/[\s\-_\/]+/);
    var wordOffset = 0;
    for (var w = 0; w < words.length; w++) {
      // Recalculate offset by finding the word in the remaining string
      var wIdx = c.indexOf(words[w], wordOffset);
      if (words[w].startsWith(q)) {
        return {
          score: 0.8 + (0.1 * q.length / c.length),
          matchType: 'prefix',
          matchRanges: [[wIdx, wIdx + q.length]]
        };
      }
      wordOffset = wIdx + words[w].length;
    }

    // Fuzzy match (only for queries >= 2 chars)
    if (q.length >= 2) {
      var dist = levenshtein(q, c);
      var maxLen = Math.max(q.length, c.length);
      var similarity = 1 - (dist / maxLen);

      var seqResult = subsequenceMatch(q, candidate);
      var fuzzyScore = Math.max(similarity, seqResult.score);

      // Dynamic threshold: shorter queries need closer matches
      var threshold = q.length <= 3 ? 0.5 : 0.4;

      if (fuzzyScore > threshold) {
        return {
          score: fuzzyScore * 0.65,
          matchType: 'fuzzy',
          matchRanges: seqResult.ranges.length > 0 ? seqResult.ranges : []
        };
      }
    }

    return { score: 0, matchType: 'none', matchRanges: [] };
  }

  /**
   * Search a list of categories against a query.
   * Returns sorted array of { category, score, matchType, matchRanges }.
   */
  function search(query, categories) {
    if (!query || !query.trim()) return [];

    var results = [];
    for (var i = 0; i < categories.length; i++) {
      var cat = categories[i];
      var result = score(query, cat.displayName);
      if (result.score > 0) {
        results.push({
          category: cat,
          score: result.score,
          matchType: result.matchType,
          matchRanges: result.matchRanges
        });
      }
    }

    results.sort(function (a, b) { return b.score - a.score; });
    return results;
  }

  /**
   * Check if query exactly matches any category name (case-insensitive).
   */
  function hasExactMatch(query, categories) {
    var q = query.toLowerCase().trim();
    for (var i = 0; i < categories.length; i++) {
      if (categories[i].displayName.toLowerCase() === q) return true;
    }
    return false;
  }

  return {
    search: search,
    score: score,
    hasExactMatch: hasExactMatch,
    mergeRanges: mergeRanges,
    levenshtein: levenshtein
  };
})();
