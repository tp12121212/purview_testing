const SAMPLE_DETECTORS = [
  {
    id: 'email-address',
    name: 'Email address',
    description: 'Basic email address pattern.',
    pattern: /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi,
    minCount: 1,
    confidence: 60,
    sensitiveTypeId: ''
  },
  {
    id: 'us-ssn',
    name: 'U.S. Social Security Number',
    description: 'Detects ###-##-#### patterns.',
    pattern: /\b(?!000|666|9\d\d)\d{3}[- ]?(?!00)\d{2}[- ]?(?!0000)\d{4}\b/g,
    minCount: 1,
    confidence: 80,
    sensitiveTypeId: ''
  },
  {
    id: 'credit-card',
    name: 'Credit card number',
    description: 'Detects common credit card digit patterns.',
    pattern: /\b(?:\d[ -]*?){13,16}\b/g,
    minCount: 1,
    confidence: 75,
    sensitiveTypeId: ''
  }
];

const parseRegexDefinition = (pattern, defaultFlags = 'gi') => {
  if (pattern instanceof RegExp) {
    return { source: pattern.source, flags: pattern.flags || 'g' };
  }
  if (typeof pattern === 'string') {
    let source = pattern.trim();
    if (!source) {
      return null;
    }
    let flags = defaultFlags;

    if (source.startsWith('/') && source.lastIndexOf('/') > 0) {
      const lastSlash = source.lastIndexOf('/');
      const rawFlags = source.slice(lastSlash + 1);
      source = source.slice(1, lastSlash);
      flags += rawFlags;
    }

    const inlineFlags = source.match(/^\(\?([imsu]+)\)/);
    if (inlineFlags) {
      source = source.replace(/^\(\?[imsu]+\)/, '');
      flags += inlineFlags[1];
    }

    if (source.includes('\\p{') || source.includes('\\P{') || source.includes('\\p')) {
      flags += 'u';
    }

    return { source, flags };
  }
  return null;
};

const ensureRegex = (pattern) => {
  const parsed = parseRegexDefinition(pattern);
  if (!parsed) {
    return null;
  }
  const flagSet = new Set(parsed.flags.split(''));
  flagSet.add('g');
  const flags = Array.from(flagSet).join('');
  try {
    return new RegExp(parsed.source, flags);
  } catch (error) {
    return null;
  }
};

export const getSampleDetectors = () => SAMPLE_DETECTORS.map((detector) => ({
  ...detector,
  pattern: detector.pattern instanceof RegExp ? detector.pattern.source : detector.pattern
}));

export const findInvalidDetectors = (detectors) => detectors.filter((detector) => !ensureRegex(detector.pattern));

export const detectSensitiveInfo = (text, detectors) => {
  if (!text) {
    return [];
  }
  const results = [];

  detectors.forEach((detector) => {
    const regex = ensureRegex(detector.pattern);
    if (!regex) {
      return;
    }
    const matches = Array.from(text.matchAll(regex));
    if (matches.length < (Number(detector.minCount) || 1)) {
      return;
    }

    results.push({
      id: detector.id,
      name: detector.name,
      count: matches.length,
      confidence: Number(detector.confidence) || 50,
      sensitiveTypeId: detector.sensitiveTypeId?.trim() || '',
      samples: matches.slice(0, 5).map((match) => match[0])
    });
  });

  return results;
};

export const aggregateClassificationResults = (classificationResults) => {
  const map = new Map();
  classificationResults.forEach((result) => {
    const key = result.sensitiveTypeId || result.id;
    if (!key) {
      return;
    }
    const existing = map.get(key);
    if (!existing) {
      map.set(key, { ...result });
      return;
    }
    existing.count += result.count;
    existing.confidence = Math.max(existing.confidence, result.confidence);
    existing.samples = Array.from(new Set([...(existing.samples ?? []), ...(result.samples ?? [])])).slice(0, 5);
  });
  return Array.from(map.values());
};

export const buildGraphClassificationResults = (classificationResults) => aggregateClassificationResults(classificationResults)
  .filter((result) => result.sensitiveTypeId)
  .map((result) => ({
    sensitiveTypeId: result.sensitiveTypeId,
    count: result.count,
    confidenceLevel: result.confidence
  }));

const mergeSamples = (existing, additional, limit = 5) => {
  const merged = [...existing];
  for (const entry of additional) {
    if (!entry) {
      continue;
    }
    if (!merged.includes(entry)) {
      merged.push(entry);
    }
    if (merged.length >= limit) {
      break;
    }
  }
  return merged.slice(0, limit);
};

const runPatternMatches = (text, pattern, flags) => {
  try {
    const regex = new RegExp(pattern, flags);
    const matches = Array.from(text.matchAll(regex));
    return {
      count: matches.length,
      samples: matches.slice(0, 5).map((match) => match[0])
    };
  } catch (error) {
    return { count: 0, samples: [] };
  }
};

const evaluateRegexNode = (text, node) => {
  const parsed = parseRegexDefinition(node.pattern);
  if (!parsed) {
    return { matched: false, count: 0, samples: [] };
  }
  const { count, samples } = runPatternMatches(text, parsed.source, parsed.flags);
  const matchesRequired = node.minMatches || 1;
  return {
    matched: count >= matchesRequired,
    count,
    samples
  };
};

const evaluateKeywordNode = (text, node) => {
  let totalCount = 0;
  let sampleValues = [];
  node.entries.forEach((entry) => {
    const { count, samples } = runPatternMatches(text, entry.pattern, entry.flags);
    if (count > 0) {
      totalCount += count;
      sampleValues = mergeSamples(sampleValues, samples);
    }
  });
  const matchesRequired = node.minMatches || 1;
  return {
    matched: totalCount >= matchesRequired,
    count: totalCount,
    samples: sampleValues
  };
};

const evaluateAnyNode = (text, node) => {
  let totalCount = 0;
  let sampleValues = [];
  node.children.forEach((child) => {
    const result = evaluateNode(text, child);
    totalCount += result.count;
    sampleValues = mergeSamples(sampleValues, result.samples);
  });
  const minMatches = node.minMatches ?? 1;
  const maxMatches = node.maxMatches;
  const withinMax = maxMatches === null || typeof maxMatches === 'undefined' || totalCount <= maxMatches;
  return {
    matched: totalCount >= minMatches && withinMax,
    count: totalCount,
    samples: sampleValues
  };
};

const evaluateNode = (text, node) => {
  if (!node) {
    return { matched: false, count: 0, samples: [] };
  }
  if (node.type === 'regex') {
    return evaluateRegexNode(text, node);
  }
  if (node.type === 'keyword') {
    return evaluateKeywordNode(text, node);
  }
  if (node.type === 'any') {
    return evaluateAnyNode(text, node);
  }
  return { matched: false, count: 0, samples: [] };
};

const evaluatePattern = (text, pattern) => {
  let totalCount = 0;
  let sampleValues = [];
  for (const node of pattern.nodes) {
    const result = evaluateNode(text, node);
    if (!result.matched) {
      return { matched: false };
    }
    totalCount += result.count;
    sampleValues = mergeSamples(sampleValues, result.samples);
  }
  return {
    matched: true,
    count: totalCount,
    samples: sampleValues,
    confidence: pattern.confidence ?? 50
  };
};

export const evaluateSitDetectors = (text, detectors) => {
  if (!text || !detectors?.length) {
    return [];
  }
  return detectors.map((detector) => {
    const patternResult = evaluatePattern(text, detector.pattern);
    if (!patternResult.matched || patternResult.count === 0) {
      return null;
    }
    return {
      id: detector.id,
      name: detector.name,
      description: detector.description,
      count: patternResult.count,
      confidence: Math.max(detector.confidence ?? 50, patternResult.confidence ?? 50),
      sensitiveTypeId: detector.sitId,
      samples: patternResult.samples,
      source: 'rulepack'
    };
  }).filter(Boolean);
};
