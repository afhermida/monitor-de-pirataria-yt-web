(function (global) {
  'use strict';

  const VERSION = '2.0.0-web';
  const CACHE_VERSION = 1;
  const MAX_API_KEYS = 10;
  const DEFAULT_SEARCH_BUDGET_PER_KEY = 95;

  const WEIGHTS = Object.freeze({
    titlePiracyPhrase: 0.35,
    suspiciousChannel: 0.25,
    longDuration: 0.20,
    unapprovedChannel: 0.10,
    bookTitleMatch: 0.10,
    descriptionPhrase: 0.10,
  });

  const TITLE_PIRACY_PHRASES = Object.freeze([
    ['audiobook completo', 'audiobook completo'],
    ['audio book completo', 'audiobook completo'],
    ['audiolivro completo', 'livro completo'],
    ['audio livro completo', 'livro completo'],
    ['livro completo', 'livro completo'],
    ['full audiobook', 'full audiobook'],
    ['complete audiobook', 'complete audiobook'],
    ['audiobook full', 'full audiobook'],
    ['full audio book', 'full audiobook'],
  ]);

  const DESCRIPTION_PHRASES = Object.freeze([
    ['completo', 'completo'],
    ['full', 'full'],
    ['gratuito', 'gratuito'],
    ['gratis', 'grátis'],
    ['free', 'free'],
    ['download', 'download'],
    ['mp3', 'mp3'],
  ]);

  const SUMMARY_PHRASES = Object.freeze([
    'resumo', 'resenha', 'analise', 'review', 'explicado', 'explicacao',
    'sinopse', 'comentado', 'comentario sobre',
  ]);

  const AUDIOBOOK_MARKERS = Object.freeze([
    'audiobook', 'audio book', 'audiolivro', 'audio livro', 'livro narrado',
    'narracao humana', 'voz humana',
  ]);

  const DEFAULT_SUSPICIOUS_CHANNEL_TERMS = Object.freeze([
    'audio livros', 'audiolivros', 'audiobooks', 'audiobook br', 'acervo audio',
    'livros narrados', 'ouca livros', 'voz das letras', 'livros gratis',
    'livros gratuitos',
  ]);

  const HEADER_TITLE_ALIASES = new Set([
    'titulo', 'titulo do livro', 'livro', 'nome do livro', 'title', 'book title',
  ]);
  const HEADER_AUTHOR_ALIASES = new Set([
    'autor', 'autora', 'autores', 'author', 'authors',
  ]);

  const STOPWORDS = new Set([
    'a', 'o', 'as', 'os', 'um', 'uma', 'uns', 'umas', 'de', 'da', 'do', 'das',
    'dos', 'e', 'em', 'no', 'na', 'nos', 'nas', 'por', 'para', 'com', 'sem',
    'the', 'of', 'and', 'an', 'in', 'on', 'to', 'for',
  ]);

  const EDITION_MARKERS = new Set([
    'edicao', 'nova edicao', 'bolso', 'luxo', 'traducao', 'traducao direta',
    'revista', 'ampliada', 'volume', 'vol', 'capa dura', 'ed especial',
    'edicao especial',
  ]);

  class MonitorError extends Error {
    constructor(message, details = null) {
      super(message);
      this.name = 'MonitorError';
      this.details = details;
    }
  }

  class ApiRequestError extends MonitorError {
    constructor(message, status, payload, reasons) {
      super(message, payload);
      this.name = 'ApiRequestError';
      this.status = status;
      this.payload = payload;
      this.reasons = reasons || new Set();
    }
  }

  class QuotaReachedError extends MonitorError {
    constructor(message) {
      super(message);
      this.name = 'QuotaReachedError';
    }
  }

  function normalize(value) {
    const text = value == null ? '' : String(value);
    return text
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .replace(/&/g, ' e ')
      .replace(/[^a-z0-9]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function cleanDisplayText(value) {
    return String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
  }

  function meaningfulTokens(value) {
    return normalize(value).split(' ').filter((token) => token && !STOPWORDS.has(token) && token.length >= 2);
  }

  function titleAliases(title) {
    const original = cleanDisplayText(title);
    const noParentheses = original.replace(/\s*\([^)]*\)\s*/g, ' ').replace(/\s+/g, ' ').trim();
    const candidates = [noParentheses];
    if (noParentheses.includes(':')) {
      const splitAt = noParentheses.indexOf(':');
      const head = noParentheses.slice(0, splitAt).trim();
      const tail = noParentheses.slice(splitAt + 1).trim();
      const normalizedTail = normalize(tail);
      if ([...EDITION_MARKERS].some((marker) => normalizedTail.includes(marker))) {
        candidates.push(head);
      }
    }

    const expanded = [];
    for (const candidate of candidates) {
      expanded.push(candidate);
      for (const part of candidate.split(/\s*\/\s*/)) {
        if (part.trim()) expanded.push(part.trim());
      }
    }

    const seen = new Set();
    const unique = [];
    for (const candidate of expanded) {
      const normalized = normalize(candidate);
      if (normalized && !seen.has(normalized)) {
        seen.add(normalized);
        unique.push(normalized);
      }
    }
    return unique.length ? unique : [normalize(title)];
  }

  function authorTokens(author) {
    const tokens = meaningfulTokens(author);
    if (!tokens.length) return [];
    return [...new Set([tokens[tokens.length - 1], ...tokens.slice(0, -1)])];
  }

  function escapeRegExp(value) {
    return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function containsToken(normalizedText, token) {
    return new RegExp(`(^|[^a-z0-9])${escapeRegExp(token)}([^a-z0-9]|$)`).test(normalizedText);
  }

  function authorMatches(author, videoTitle, description) {
    const haystack = `${normalize(videoTitle)} ${normalize(description)}`;
    const tokens = authorTokens(author);
    return tokens.slice(0, 2).some((token) => containsToken(haystack, token));
  }

  function levenshteinRatio(a, b) {
    if (a === b) return 1;
    if (!a.length || !b.length) return 0;
    let left = a;
    let right = b;
    if (left.length > right.length) [left, right] = [right, left];
    const previous = new Uint16Array(left.length + 1);
    const current = new Uint16Array(left.length + 1);
    for (let i = 0; i <= left.length; i += 1) previous[i] = i;
    for (let j = 1; j <= right.length; j += 1) {
      current[0] = j;
      const rb = right.charCodeAt(j - 1);
      for (let i = 1; i <= left.length; i += 1) {
        const cost = left.charCodeAt(i - 1) === rb ? 0 : 1;
        current[i] = Math.min(
          current[i - 1] + 1,
          previous[i] + 1,
          previous[i - 1] + cost,
        );
      }
      previous.set(current);
    }
    return 1 - (previous[left.length] / Math.max(left.length, right.length));
  }

  function computeTitleMatch(bookTitle, videoTitle) {
    const target = normalize(videoTitle);
    let bestScore = 0;
    let exact = false;
    let bestAlias = '';

    for (const alias of titleAliases(bookTitle)) {
      const aliasTokens = alias.split(' ').filter((token) => token && !STOPWORDS.has(token));
      if (!aliasTokens.length) continue;
      let score;
      let isExact;
      if (target.includes(alias)) {
        score = 1;
        isExact = true;
      } else {
        const matched = aliasTokens.filter((token) => containsToken(target, token)).length;
        const coverage = matched / aliasTokens.length;
        const compactTarget = target.split(' ').filter((token) => !STOPWORDS.has(token)).join(' ');
        const charRatio = levenshteinRatio(alias, target);
        const tokenRatio = levenshteinRatio(aliasTokens.join(' '), compactTarget);
        score = Math.max((coverage * 0.82) + (tokenRatio * 0.18), charRatio * 0.85);
        isExact = false;
      }
      if (score > bestScore) {
        bestScore = score;
        exact = isExact;
        bestAlias = alias;
      }
    }
    return { score: Math.round(bestScore * 10000) / 10000, exact, alias: bestAlias };
  }

  function bibliographicGate(book, titleMatchScore, exactTitleMatch, authorMatch) {
    const aliases = titleAliases(book.title);
    const alias = aliases.length ? aliases.reduce((a, b) => (a.length >= b.length ? a : b)) : normalize(book.title);
    const tokens = alias.split(' ').filter((token) => token && !STOPWORDS.has(token));
    const compactLength = alias.replace(/ /g, '').length;
    if (!tokens.length) return { ok: false, reason: 'título bibliográfico vazio após normalização' };

    let ok = false;
    if (tokens.length >= 4) {
      ok = titleMatchScore >= 0.58 || (titleMatchScore >= 0.48 && authorMatch);
    } else if (tokens.length === 3) {
      ok = exactTitleMatch || titleMatchScore >= 0.68 || (titleMatchScore >= 0.55 && authorMatch);
    } else if (tokens.length === 2) {
      ok = exactTitleMatch && (authorMatch || compactLength >= 11);
    } else {
      const token = tokens[0];
      const distinctiveNumeric = /^\d{3,}$/.test(token);
      ok = exactTitleMatch && (authorMatch || distinctiveNumeric);
    }
    return ok ? { ok: true, reason: null } : { ok: false, reason: 'correspondência bibliográfica fraca' };
  }

  function firstPhrase(text, phrases) {
    const normalizedText = normalize(text);
    for (const [normalizedPhrase, displayPhrase] of phrases) {
      if (normalizedText.includes(normalizedPhrase)) return { normalizedPhrase, displayPhrase };
    }
    return null;
  }

  function channelIsSuspicious(channelTitle, suspiciousTerms) {
    const normalizedTitle = normalize(channelTitle);
    return suspiciousTerms.some((term) => normalize(term) && normalizedTitle.includes(normalize(term)));
  }

  function parseChannelRules(text) {
    const ids = new Set();
    const names = new Set();
    for (const rawLine of String(text || '').split(/\r?\n/)) {
      const line = rawLine.trim();
      if (!line || line.startsWith('#')) continue;
      const lower = line.toLowerCase();
      if (lower.startsWith('id:')) {
        const value = line.split(':').slice(1).join(':').trim();
        if (value) ids.add(value);
      } else if (lower.startsWith('nome:') || lower.startsWith('name:')) {
        const value = line.split(':').slice(1).join(':').trim();
        if (value) names.add(normalize(value));
      } else if (line.startsWith('UC') && line.length >= 20) {
        ids.add(line);
      } else {
        names.add(normalize(line));
      }
    }
    return { ids, names };
  }

  function parseTerms(text, defaults = DEFAULT_SUSPICIOUS_CHANNEL_TERMS) {
    const terms = [...defaults];
    for (const rawLine of String(text || '').split(/\r?\n/)) {
      const line = rawLine.trim();
      if (line && !line.startsWith('#')) terms.push(line);
    }
    const unique = [];
    const seen = new Set();
    for (const term of terms) {
      const normalized = normalize(term);
      if (normalized && !seen.has(normalized)) {
        seen.add(normalized);
        unique.push(term.trim());
      }
    }
    return unique;
  }

  function isAuthorizedChannel(video, authorizedIds, authorizedNames) {
    return authorizedIds.has(video.channelId) || authorizedNames.has(normalize(video.channelTitle));
  }

  function parseIsoDuration(value) {
    const match = /^P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/.exec(String(value || ''));
    if (!match) return 0;
    return (Number(match[1] || 0) * 86400) + (Number(match[2] || 0) * 3600)
      + (Number(match[3] || 0) * 60) + Number(match[4] || 0);
  }

  function formatDuration(seconds) {
    const safe = Math.max(0, Number(seconds) || 0);
    const totalMinutes = Math.floor(safe / 60);
    const hours = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;
    return hours ? `${hours}h${minutes}m` : `${minutes}m`;
  }

  function evaluateVideo(book, video, config) {
    const {
      profile = 'equilibrado', threshold = 0.5, longDurationMinutes = 20,
      authorizedIds = new Set(), authorizedNames = new Set(),
      suspiciousChannelTerms = DEFAULT_SUSPICIOUS_CHANNEL_TERMS,
      includeSummaries = false,
    } = config || {};

    const reasons = [];
    let score = 0;
    if (isAuthorizedChannel(video, authorizedIds, authorizedNames)) {
      return {
        book, video, score: 0, reasons: ['canal autorizado'], titleMatchScore: 0,
        authorMatch: false, accepted: false, rejectedReason: 'canal autorizado',
      };
    }

    const titleMatch = computeTitleMatch(book.title, video.title);
    const authorMatch = authorMatches(book.author, video.title, video.description);
    const titlePhrase = firstPhrase(video.title, TITLE_PIRACY_PHRASES);
    if (titlePhrase) {
      score += WEIGHTS.titlePiracyPhrase;
      reasons.push(`título contém "${titlePhrase.displayPhrase}"`);
    }
    if (channelIsSuspicious(video.channelTitle, suspiciousChannelTerms)) {
      score += WEIGHTS.suspiciousChannel;
      reasons.push(`canal suspeito: "${normalize(video.channelTitle)}"`);
    }
    const durationMinutes = Math.floor((Number(video.durationSeconds) || 0) / 60);
    if (durationMinutes >= longDurationMinutes) {
      score += WEIGHTS.longDuration;
      reasons.push(`duração longa: ${durationMinutes} min`);
    }
    score += WEIGHTS.unapprovedChannel;
    reasons.push('canal não verificado');

    const titleReasonMatch = titleMatch.exact || titleMatch.score >= 0.72;
    if (titleReasonMatch) {
      score += WEIGHTS.bookTitleMatch;
      reasons.push('título do livro no vídeo');
    }
    const descriptionPhrase = firstPhrase(video.description, DESCRIPTION_PHRASES);
    if (descriptionPhrase) {
      score += WEIGHTS.descriptionPhrase;
      reasons.push(`descrição contém "${descriptionPhrase.displayPhrase}"`);
    }

    let rejectedReason = null;
    if (profile === 'equilibrado') {
      const gate = bibliographicGate(book, titleMatch.score, titleMatch.exact, authorMatch);
      if (!gate.ok) rejectedReason = gate.reason;
      const normalizedTitle = normalize(video.title);
      const hasSummarySignal = SUMMARY_PHRASES.some((phrase) => normalizedTitle.includes(phrase));
      if (hasSummarySignal && !includeSummaries) {
        score -= 0.30;
        reasons.push('possível resumo/resenha: -0,30');
      }
      const hasAudioMarker = AUDIOBOOK_MARKERS.some((marker) => normalizedTitle.includes(marker));
      if (durationMinutes < 45 && !titlePhrase && !hasAudioMarker && !includeSummaries) {
        score -= 0.25;
        reasons.push('vídeo curto sem indicação clara de audiobook: -0,25');
      }
    }

    score = Math.round(Math.min(1, Math.max(0, score)) * 100) / 100;
    return {
      book, video, score, reasons, titleMatchScore: titleMatch.score, authorMatch,
      accepted: rejectedReason == null && score >= threshold,
      rejectedReason,
    };
  }

  function bookIdentity(book) {
    return `${normalize(book.title)}|${normalize(book.author)}`;
  }

  function buildQuery(book) {
    const displayAlias = titleAliases(book.title)[0] || normalize(book.title);
    const author = cleanDisplayText(book.author);
    return `"${displayAlias}" "${author}" audiobook|audiolivro`.trim();
  }

  function fnv1a64(text) {
    let hash = 0xcbf29ce484222325n;
    const prime = 0x100000001b3n;
    const bytes = new TextEncoder().encode(String(text));
    for (const byte of bytes) {
      hash ^= BigInt(byte);
      hash = BigInt.asUintN(64, hash * prime);
    }
    return hash.toString(16).padStart(16, '0');
  }

  function cacheKey(book, options) {
    const payload = [
      bookIdentity(book), buildQuery(book), String(options.maxResults), options.region,
      options.language, options.publishedAfter || '',
    ].join('|');
    return `v${CACHE_VERSION}-${fnv1a64(payload)}`;
  }

  function decodeHtmlEntities(text) {
    const value = String(text || '');
    if (typeof document !== 'undefined') {
      const textarea = document.createElement('textarea');
      textarea.innerHTML = value;
      return textarea.value;
    }
    return value
      .replace(/&amp;/g, '&').replace(/&#39;/g, "'").replace(/&quot;/g, '"')
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>');
  }

  function extractErrorReasons(payload) {
    const reasons = new Set();
    if (!payload || typeof payload !== 'object') return reasons;
    const error = payload.error && typeof payload.error === 'object' ? payload.error : {};
    if (error.status) reasons.add(String(error.status).toLowerCase());
    if (Array.isArray(error.errors)) {
      for (const detail of error.errors) {
        if (detail && detail.reason) reasons.add(String(detail.reason).toLowerCase());
      }
    }
    return reasons;
  }

  function isQuotaError(error) {
    const reasons = error && error.reasons ? error.reasons : new Set();
    const message = String(error && error.message ? error.message : '').toLowerCase();
    const quotaReasons = new Set([
      'quotaexceeded', 'dailylimitexceeded', 'dailylimitexceededunreg',
      'ratelimitexceeded', 'userratelimitexceeded', 'uploadlimitexceeded',
      'resource_exhausted', 'too_many_requests',
    ]);
    if (error && error.status === 429) return true;
    return error && error.status === 403 && (
      [...reasons].some((reason) => quotaReasons.has(reason))
      || /quota|ratelimit|rate limit|dailylimit/.test(message)
    );
  }

  function isKeyConfigurationError(error) {
    const reasons = error && error.reasons ? error.reasons : new Set();
    const message = String(error && error.message ? error.message : '').toLowerCase();
    const keyReasons = new Set([
      'accessnotconfigured', 'apikeynotvalid', 'autherror', 'forbidden',
      'iprefererblocked', 'keyexpired', 'keyinvalid', 'permission_denied',
      'api_key_invalid',
    ]);
    if (error && error.status === 401) return true;
    return error && [400, 403].includes(error.status) && (
      [...reasons].some((reason) => keyReasons.has(reason))
      || /api key not valid|keyinvalid|access not configured|accessnotconfigured|iprefererblocked|permission_denied|forbidden|api key expired/.test(message)
    );
  }

  function delay(ms, signal) {
    return new Promise((resolve, reject) => {
      const timer = setTimeout(resolve, ms);
      if (signal) {
        signal.addEventListener('abort', () => {
          clearTimeout(timer);
          reject(new DOMException('Operação interrompida', 'AbortError'));
        }, { once: true });
      }
    });
  }

  async function fetchJson(url, { signal, fetchImpl = global.fetch } = {}) {
    let response;
    try {
      response = await fetchImpl(url, {
        method: 'GET', mode: 'cors', cache: 'no-store', credentials: 'omit',
        referrerPolicy: 'no-referrer', signal,
      });
    } catch (error) {
      if (error && error.name === 'AbortError') throw error;
      throw new ApiRequestError(`Falha de rede: ${error && error.message ? error.message : error}`, 0, null, new Set());
    }

    let payload = null;
    const raw = await response.text();
    if (raw) {
      try { payload = JSON.parse(raw); } catch (_) { payload = { raw: raw.slice(0, 800) }; }
    }
    if (!response.ok) {
      const reasons = extractErrorReasons(payload);
      const apiMessage = payload && payload.error && payload.error.message
        ? String(payload.error.message) : `HTTP ${response.status}`;
      throw new ApiRequestError(apiMessage, response.status, payload, reasons);
    }
    return payload || {};
  }

  class YouTubeKeyPool {
    constructor(apiKeys, options = {}) {
      const unique = [...new Set((apiKeys || []).map((key) => String(key).trim()).filter(Boolean))];
      if (!unique.length) throw new MonitorError('Informe pelo menos uma chave da API do YouTube.');
      if (unique.length > MAX_API_KEYS) throw new MonitorError(`Informe no máximo ${MAX_API_KEYS} chaves de API.`);
      this.apiKeys = unique;
      this.currentIndex = 0;
      this.usedIndices = new Set();
      this.unavailableIndices = new Set();
      this.rotations = 0;
      this.fetchImpl = options.fetchImpl || global.fetch;
      this.onEvent = typeof options.onEvent === 'function' ? options.onEvent : () => {};
      this.sleep = options.sleep || delay;
    }

    get configuredCount() { return this.apiKeys.length; }
    get usedCount() { return this.usedIndices.size; }
    get unavailableCount() { return this.unavailableIndices.size; }
    get activeDisplayIndex() { return this.currentIndex + 1; }

    snapshot() {
      return {
        configured: this.configuredCount,
        used: this.usedCount,
        unavailable: this.unavailableCount,
        rotations: this.rotations,
        currentIndex: this.currentIndex,
        unavailableIndices: [...this.unavailableIndices],
        usedIndices: [...this.usedIndices],
      };
    }

    rotate(reason) {
      const oldIndex = this.currentIndex;
      this.unavailableIndices.add(oldIndex);
      let nextIndex = null;
      for (let index = oldIndex + 1; index < this.apiKeys.length; index += 1) {
        if (!this.unavailableIndices.has(index)) { nextIndex = index; break; }
      }
      if (nextIndex == null) {
        this.onEvent({ type: 'key-exhausted', keyIndex: oldIndex, reason, snapshot: this.snapshot() });
        return false;
      }
      this.currentIndex = nextIndex;
      this.rotations += 1;
      this.onEvent({ type: 'key-rotated', from: oldIndex, to: nextIndex, reason, snapshot: this.snapshot() });
      return true;
    }

    async execute(buildUrl, { signal, retries = 3, label = 'consulta' } = {}) {
      while (true) {
        const index = this.currentIndex;
        this.usedIndices.add(index);
        this.onEvent({ type: 'key-active', keyIndex: index, label, snapshot: this.snapshot() });
        let retryDelay = 1200;
        for (let attempt = 0; attempt <= retries; attempt += 1) {
          try {
            return await fetchJson(buildUrl(this.apiKeys[index]), { signal, fetchImpl: this.fetchImpl });
          } catch (error) {
            if (error && error.name === 'AbortError') throw error;
            const quotaError = isQuotaError(error);
            const keyError = isKeyConfigurationError(error);
            if (quotaError || keyError) {
              const reason = quotaError ? 'cota ou limite atingido' : 'chave inválida, restrita ou API desativada';
              if (this.rotate(reason)) break;
              throw new QuotaReachedError(`Todas as ${this.apiKeys.length} chaves ficaram indisponíveis nesta execução.`);
            }
            const retryable = !error.status || [429, 500, 502, 503, 504].includes(error.status);
            if (!retryable || attempt >= retries) {
              throw new MonitorError(`Falha ao consultar a API do YouTube (${label}): ${error.message || error}`);
            }
            this.onEvent({ type: 'retry', keyIndex: index, label, attempt: attempt + 1, snapshot: this.snapshot() });
            await this.sleep(retryDelay, signal);
            retryDelay *= 2;
          }
        }
      }
    }
  }

  function makeYouTubeUrl(path, params, apiKey) {
    const url = new URL(`https://www.googleapis.com/youtube/v3/${path}`);
    for (const [key, value] of Object.entries(params)) {
      if (value != null && value !== '') url.searchParams.set(key, String(value));
    }
    url.searchParams.set('key', apiKey);
    return url.toString();
  }

  async function searchYouTube(keyPool, book, options, signal) {
    const query = buildQuery(book);
    const searchParams = {
      part: 'id', q: query, type: 'video', maxResults: options.maxResults,
      order: 'relevance', safeSearch: 'none', regionCode: options.region,
      relevanceLanguage: options.language,
    };
    if (options.publishedAfter) searchParams.publishedAfter = options.publishedAfter;

    const searchResponse = await keyPool.execute(
      (key) => makeYouTubeUrl('search', searchParams, key),
      { signal, label: 'search.list' },
    );
    const orderedIds = [];
    const seen = new Set();
    for (const item of searchResponse.items || []) {
      const id = item && item.id ? item.id.videoId : null;
      if (id && !seen.has(id)) { seen.add(id); orderedIds.push(id); }
    }
    if (!orderedIds.length) return [];

    const detailsResponse = await keyPool.execute(
      (key) => makeYouTubeUrl('videos', {
        part: 'snippet,contentDetails,statistics,status', id: orderedIds.join(','),
        maxResults: orderedIds.length,
      }, key),
      { signal, label: 'videos.list' },
    );
    const byId = new Map((detailsResponse.items || []).map((item) => [String(item.id), item]));
    const videos = [];
    orderedIds.forEach((videoId, rankIndex) => {
      const item = byId.get(videoId);
      if (!item) return;
      const snippet = item.snippet || {};
      const details = item.contentDetails || {};
      const statistics = item.statistics || {};
      videos.push({
        videoId,
        title: cleanDisplayText(decodeHtmlEntities(snippet.title || '')),
        description: decodeHtmlEntities(snippet.description || ''),
        channelId: String(snippet.channelId || ''),
        channelTitle: cleanDisplayText(decodeHtmlEntities(snippet.channelTitle || '')),
        publishedAt: String(snippet.publishedAt || ''),
        durationSeconds: parseIsoDuration(details.duration),
        viewCount: Number.parseInt(statistics.viewCount || '0', 10) || 0,
        searchRank: rankIndex + 1,
        url: `https://www.youtube.com/watch?v=${videoId}`,
      });
    });
    return videos;
  }

  function detectHeader(rows, options = {}) {
    const explicitTitle = options.titleColumnName ? normalize(options.titleColumnName) : '';
    const explicitAuthor = options.authorColumnName ? normalize(options.authorColumnName) : '';
    for (let rowIndex = 0; rowIndex < Math.min(rows.length, 20); rowIndex += 1) {
      const row = rows[rowIndex] || [];
      const normalizedCells = row.map(normalize);
      let titleIndex = -1;
      let authorIndex = -1;
      normalizedCells.forEach((value, columnIndex) => {
        if (value === explicitTitle || (!explicitTitle && HEADER_TITLE_ALIASES.has(value))) titleIndex = columnIndex;
        if (value === explicitAuthor || (!explicitAuthor && HEADER_AUTHOR_ALIASES.has(value))) authorIndex = columnIndex;
      });
      if (titleIndex >= 0 && authorIndex >= 0) return { headerRow: rowIndex, titleIndex, authorIndex };
    }
    return null;
  }

  function chooseCatalogSheet(workbook) {
    const candidates = [];
    (workbook.sheets || []).forEach((sheet, sheetIndex) => {
      const header = detectHeader(sheet.rows || []);
      if (!header) return;
      const normalizedName = normalize(sheet.name);
      const preferred = ['catalogo', 'catalog', 'livros', 'titulos'].some((marker) => normalizedName.includes(marker)) ? 1 : 0;
      candidates.push({ sheetIndex, header, preferred, rowCount: (sheet.rows || []).length });
    });
    if (!candidates.length) return null;
    candidates.sort((a, b) => b.preferred - a.preferred || b.rowCount - a.rowCount || a.sheetIndex - b.sheetIndex);
    return candidates[0];
  }

  function extractBooks(sheet, headerRow, titleIndex, authorIndex) {
    const books = [];
    const seen = new Set();
    const rows = sheet && sheet.rows ? sheet.rows : [];
    for (let index = Number(headerRow) + 1; index < rows.length; index += 1) {
      const row = rows[index] || [];
      const title = cleanDisplayText(row[titleIndex]);
      const author = cleanDisplayText(row[authorIndex]);
      if (!title) continue;
      const identity = `${normalize(title)}|${normalize(author)}`;
      if (seen.has(identity)) continue;
      seen.add(identity);
      books.push({ position: books.length + 1, title, author, identity });
    }
    return books;
  }

  function columnLabel(index) {
    let number = Number(index) + 1;
    let label = '';
    while (number > 0) {
      const remainder = (number - 1) % 26;
      label = String.fromCharCode(65 + remainder) + label;
      number = Math.floor((number - 1) / 26);
    }
    return label;
  }

  function buildAnalysis(books, state, config) {
    const evaluations = [];
    const acceptedByBook = new Map();
    const resultsByBook = state.resultsByBook instanceof Map
      ? state.resultsByBook : new Map(Object.entries(state.resultsByBook || {}));
    for (const book of books) {
      const rawVideos = resultsByBook.get(book.identity || bookIdentity(book)) || [];
      const seenVideoIds = new Set();
      for (const video of rawVideos) {
        if (seenVideoIds.has(video.videoId)) continue;
        seenVideoIds.add(video.videoId);
        const evaluation = evaluateVideo(book, video, config);
        evaluations.push(evaluation);
        if (evaluation.accepted) {
          const key = book.identity || bookIdentity(book);
          if (!acceptedByBook.has(key)) acceptedByBook.set(key, []);
          acceptedByBook.get(key).push(evaluation);
        }
      }
    }

    const rankedBooks = [];
    for (const book of books) {
      const key = book.identity || bookIdentity(book);
      const accepted = acceptedByBook.get(key) || [];
      if (!accepted.length) continue;
      const mostViewed = accepted.reduce((best, item) => (
        !best || item.video.viewCount > best.video.viewCount
        || (item.video.viewCount === best.video.viewCount && item.score > best.score)
        || (item.video.viewCount === best.video.viewCount && item.score === best.score && item.video.searchRank < best.video.searchRank)
          ? item : best
      ), null);
      const maxScore = Math.max(...accepted.map((item) => item.score));
      rankedBooks.push({ book, accepted, mostViewed, maxScore });
    }
    rankedBooks.sort((a, b) => (
      b.mostViewed.video.viewCount - a.mostViewed.video.viewCount
      || b.maxScore - a.maxScore
      || b.book.title.localeCompare(a.book.title, 'pt-BR')
    ));

    const acceptedEvaluations = evaluations.filter((evaluation) => evaluation.accepted);
    const pendingBooks = state.pendingBooks instanceof Set ? state.pendingBooks : new Set(state.pendingBooks || []);
    const pending = books.filter((book) => pendingBooks.has(book.identity || bookIdentity(book))).length;
    const processed = books.length - pending;
    const suspiciousTitles = acceptedByBook.size;
    const cleanTitles = Math.max(0, processed - suspiciousTitles);

    return {
      evaluations,
      acceptedEvaluations,
      acceptedByBook,
      rankedBooks,
      summary: {
        total: books.length,
        processed,
        pending,
        suspiciousTitles,
        cleanTitles,
        percentage: processed ? suspiciousTitles / processed : 0,
        suspiciousVideos: acceptedEvaluations.length,
      },
    };
  }

  function selfTest() {
    const failures = [];
    const book = { position: 1, title: 'O Pequeno Príncipe', author: 'Antoine de Saint-Exupéry', identity: 'x' };
    const video = {
      videoId: 'abc', title: 'O Pequeno Príncipe - Audiobook Completo',
      description: 'Ouça gratuitamente o livro completo.', channelId: 'UC_TEST',
      channelTitle: 'Audiolivros Brasil', publishedAt: '2026-01-01T00:00:00Z',
      durationSeconds: 7200, viewCount: 1000, searchRank: 1,
      url: 'https://www.youtube.com/watch?v=abc',
    };
    const result = evaluateVideo(book, video, {
      profile: 'equilibrado', threshold: 0.5, longDurationMinutes: 20,
      authorizedIds: new Set(), authorizedNames: new Set(),
      suspiciousChannelTerms: DEFAULT_SUSPICIOUS_CHANNEL_TERMS, includeSummaries: false,
    });
    if (!result.accepted || result.score < 0.9) failures.push(`classificação principal inesperada: ${JSON.stringify(result)}`);

    const falsePositive = evaluateVideo(
      { position: 2, title: 'Manual do estoicismo', author: 'Epicteto', identity: 'y' },
      { ...video, videoId: 'def', title: 'Meditações de Marco Aurélio | Audiolivro Completo', description: '' },
      {
        profile: 'equilibrado', threshold: 0.5, longDurationMinutes: 20,
        authorizedIds: new Set(), authorizedNames: new Set(),
        suspiciousChannelTerms: DEFAULT_SUSPICIOUS_CHANNEL_TERMS, includeSummaries: false,
      },
    );
    if (falsePositive.accepted) failures.push('o bloqueio bibliográfico não rejeitou um falso positivo conhecido');
    if (parseIsoDuration('PT1H2M3S') !== 3723) failures.push('parseIsoDuration incorreto');
    if (titleAliases('Livro: edição especial')[0] !== 'livro edicao especial') failures.push('titleAliases inesperado');
    return failures;
  }

  const api = {
    VERSION, CACHE_VERSION, MAX_API_KEYS, DEFAULT_SEARCH_BUDGET_PER_KEY,
    WEIGHTS, DEFAULT_SUSPICIOUS_CHANNEL_TERMS, MonitorError, ApiRequestError,
    QuotaReachedError, normalize, cleanDisplayText, meaningfulTokens, titleAliases,
    authorTokens, containsToken, authorMatches, computeTitleMatch, bibliographicGate,
    parseChannelRules, parseTerms, parseIsoDuration, formatDuration, evaluateVideo,
    bookIdentity, buildQuery, cacheKey, decodeHtmlEntities, extractErrorReasons,
    isQuotaError, isKeyConfigurationError, fetchJson, YouTubeKeyPool, searchYouTube,
    detectHeader, chooseCatalogSheet, extractBooks, columnLabel, buildAnalysis, selfTest,
  };

  global.MonitorCore = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
}(typeof window !== 'undefined' ? window : globalThis));
