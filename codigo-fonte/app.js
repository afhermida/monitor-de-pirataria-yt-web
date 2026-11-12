(function () {
  'use strict';

  const Core = window.MonitorCore;
  const Xlsx = window.XlsxLite;
  const $ = (selector) => document.querySelector(selector);
  const $$ = (selector) => Array.from(document.querySelectorAll(selector));
  const numberFormat = new Intl.NumberFormat('pt-BR');
  const percentFormat = new Intl.NumberFormat('pt-BR', { style: 'percent', maximumFractionDigits: 1 });

  const state = {
    workbook: null,
    inputFileName: '',
    books: [],
    selectedSheetIndex: 0,
    resultsByBook: new Map(),
    pendingBooks: new Set(),
    analysis: null,
    scan: null,
    logs: [],
    resultTab: 'videos',
    resultPage: 1,
    cacheDb: null,
    cacheAvailable: false,
    memoryCache: new Map(),
  };

  class CacheDatabase {
    constructor() {
      this.db = null;
    }

    open() {
      return new Promise((resolve, reject) => {
        if (!window.indexedDB) {
          reject(new Error('IndexedDB indisponível'));
          return;
        }
        const request = indexedDB.open('monitor-youtube-audiobooks', 1);
        request.onupgradeneeded = () => {
          const db = request.result;
          if (!db.objectStoreNames.contains('searchCache')) {
            const store = db.createObjectStore('searchCache', { keyPath: 'key' });
            store.createIndex('fetchedAt', 'fetchedAt', { unique: false });
            store.createIndex('bookIdentity', 'bookIdentity', { unique: false });
          }
        };
        request.onsuccess = () => { this.db = request.result; resolve(this); };
        request.onerror = () => reject(request.error || new Error('Falha ao abrir o cache local'));
      });
    }

    transaction(mode, callback) {
      return new Promise((resolve, reject) => {
        const tx = this.db.transaction('searchCache', mode);
        const store = tx.objectStore('searchCache');
        let result;
        try { result = callback(store); } catch (error) { reject(error); return; }
        tx.oncomplete = () => resolve(result);
        tx.onerror = () => reject(tx.error || new Error('Falha no cache local'));
        tx.onabort = () => reject(tx.error || new Error('Operação de cache cancelada'));
      });
    }

    async get(key) {
      return new Promise((resolve, reject) => {
        const tx = this.db.transaction('searchCache', 'readonly');
        const request = tx.objectStore('searchCache').get(key);
        request.onsuccess = () => resolve(request.result || null);
        request.onerror = () => reject(request.error);
      });
    }

    async put(entry) {
      await this.transaction('readwrite', (store) => store.put(entry));
    }

    async getAll() {
      return new Promise((resolve, reject) => {
        const tx = this.db.transaction('searchCache', 'readonly');
        const request = tx.objectStore('searchCache').getAll();
        request.onsuccess = () => resolve(request.result || []);
        request.onerror = () => reject(request.error);
      });
    }

    async clear() {
      await this.transaction('readwrite', (store) => store.clear());
    }

    async putMany(entries) {
      await this.transaction('readwrite', (store) => {
        for (const entry of entries) store.put(entry);
      });
    }
  }

  function showToast(message, type = 'info', timeout = 5000) {
    const container = $('#toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.setAttribute('role', 'status');
    toast.textContent = message;
    container.appendChild(toast);
    requestAnimationFrame(() => toast.classList.add('is-visible'));
    setTimeout(() => {
      toast.classList.remove('is-visible');
      setTimeout(() => toast.remove(), 250);
    }, timeout);
  }

  function setStatus(text, tone = 'neutral') {
    const badge = $('#globalStatus');
    badge.textContent = text;
    badge.className = `status-badge status-${tone}`;
  }

  function log(message, tone = 'info') {
    const timestamp = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    state.logs.push({ timestamp, message, tone });
    if (state.logs.length > 250) state.logs.shift();
    const list = $('#activityLog');
    const item = document.createElement('li');
    item.className = `log-${tone}`;
    const time = document.createElement('time');
    time.textContent = timestamp;
    const text = document.createElement('span');
    text.textContent = message;
    item.append(time, text);
    list.appendChild(item);
    while (list.children.length > 250) list.firstElementChild.remove();
    list.scrollTop = list.scrollHeight;
  }

  function escapeHtml(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  }

  function downloadBlob(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = fileName;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  async function readTextFile(file) {
    return file.text();
  }

  function getHeaderRow() {
    return Math.max(0, Number($('#headerRowInput').value || 1) - 1);
  }

  function maxColumnsAroundHeader(sheet, headerRow) {
    let max = 0;
    const rows = sheet.rows || [];
    for (let index = Math.max(0, headerRow - 2); index <= Math.min(rows.length - 1, headerRow + 5); index += 1) {
      max = Math.max(max, (rows[index] || []).length);
    }
    return Math.max(max, 2);
  }

  function populateColumnSelectors(preferredTitle = null, preferredAuthor = null) {
    if (!state.workbook) return;
    const sheet = state.workbook.sheets[Number($('#sheetSelect').value) || 0];
    if (!sheet) return;
    const headerRow = getHeaderRow();
    const header = sheet.rows[headerRow] || [];
    const count = maxColumnsAroundHeader(sheet, headerRow);
    const titleSelect = $('#titleColumnSelect');
    const authorSelect = $('#authorColumnSelect');
    titleSelect.innerHTML = '';
    authorSelect.innerHTML = '';
    for (let index = 0; index < count; index += 1) {
      const label = `${Core.columnLabel(index)} — ${Core.cleanDisplayText(header[index]) || '(sem cabeçalho)'}`;
      for (const select of [titleSelect, authorSelect]) {
        const option = document.createElement('option');
        option.value = String(index);
        option.textContent = label;
        select.appendChild(option);
      }
    }
    if (preferredTitle != null) titleSelect.value = String(preferredTitle);
    if (preferredAuthor != null) authorSelect.value = String(preferredAuthor);
  }

  function applyCatalogSelection({ quiet = false } = {}) {
    if (!state.workbook) return;
    const sheetIndex = Number($('#sheetSelect').value) || 0;
    const sheet = state.workbook.sheets[sheetIndex];
    const headerRow = getHeaderRow();
    const titleIndex = Number($('#titleColumnSelect').value);
    const authorIndex = Number($('#authorColumnSelect').value);
    const books = Core.extractBooks(sheet, headerRow, titleIndex, authorIndex);
    state.selectedSheetIndex = sheetIndex;
    state.books = books;
    state.resultsByBook = new Map();
    state.pendingBooks = new Set(books.map((book) => book.identity));
    state.scan = null;
    state.analysis = null;
    $('#catalogCount').textContent = numberFormat.format(books.length);
    $('#catalogSheetName').textContent = sheet.name;
    $('#catalogFileName').textContent = state.inputFileName;
    $('#catalogStatusPanel').hidden = false;
    $('#startButton').disabled = books.length === 0;
    $('#exportReportButton').disabled = books.length === 0;
    renderCatalogPreview();
    refreshAnalysis();
    updateWorkflowState();
    if (!quiet) {
      if (books.length) showToast(`${numberFormat.format(books.length)} títulos carregados.`, 'success');
      else showToast('Nenhum título válido foi encontrado com essa seleção.', 'warning');
    }
  }

  function renderCatalogPreview() {
    const body = $('#catalogPreviewBody');
    body.innerHTML = '';
    state.books.slice(0, 8).forEach((book) => {
      const row = document.createElement('tr');
      row.innerHTML = `<td>${book.position}</td><td>${escapeHtml(book.title)}</td><td>${escapeHtml(book.author)}</td>`;
      body.appendChild(row);
    });
    $('#catalogPreview').hidden = state.books.length === 0;
  }

  async function handleCatalogFile(file) {
    if (!file) return;
    if (!/\.xlsx$/i.test(file.name)) {
      showToast('Selecione uma planilha no formato .xlsx.', 'error');
      return;
    }
    setStatus('Lendo planilha', 'working');
    $('#dropZone').classList.add('is-loading');
    try {
      const workbook = await Xlsx.readWorkbook(await file.arrayBuffer());
      state.workbook = workbook;
      state.inputFileName = file.name;
      const sheetSelect = $('#sheetSelect');
      sheetSelect.innerHTML = '';
      workbook.sheets.forEach((sheet, index) => {
        const option = document.createElement('option');
        option.value = String(index);
        option.textContent = `${sheet.name} (${numberFormat.format(sheet.rows.length)} linhas)`;
        sheetSelect.appendChild(option);
      });
      const candidate = Core.chooseCatalogSheet(workbook);
      const selected = candidate || { sheetIndex: 0, header: { headerRow: 0, titleIndex: 0, authorIndex: 1 } };
      sheetSelect.value = String(selected.sheetIndex);
      $('#headerRowInput').value = String(selected.header.headerRow + 1);
      populateColumnSelectors(selected.header.titleIndex, selected.header.authorIndex);
      applyCatalogSelection({ quiet: true });
      $('#mappingPanel').hidden = false;
      $('#dropZoneLabel').textContent = file.name;
      $('#dropZoneHint').textContent = `${numberFormat.format(state.books.length)} títulos encontrados; revise o mapeamento abaixo se necessário.`;
      log(`Planilha “${file.name}” carregada com ${state.books.length} títulos.`, 'success');
      setStatus('Catálogo pronto', 'success');
      showToast(`Catálogo carregado: ${numberFormat.format(state.books.length)} títulos.`, 'success');
    } catch (error) {
      console.error(error);
      setStatus('Erro na planilha', 'error');
      showToast(error.message || 'Não foi possível ler a planilha.', 'error', 8000);
      log(`Erro ao ler planilha: ${error.message || error}`, 'error');
    } finally {
      $('#dropZone').classList.remove('is-loading');
    }
  }

  function createKeyRow(value = '') {
    const rows = $('#apiKeysList');
    if (rows.children.length >= Core.MAX_API_KEYS) {
      showToast(`O limite é de ${Core.MAX_API_KEYS} chaves.`, 'warning');
      return;
    }
    const index = rows.children.length;
    const row = document.createElement('div');
    row.className = 'key-row';
    row.innerHTML = `
      <span class="key-number">${index + 1}</span>
      <label class="sr-only" for="apiKey${index}">Chave ${index + 1}</label>
      <input id="apiKey${index}" class="api-key-input configuration-control" type="password" autocomplete="off" spellcheck="false" placeholder="Cole a chave da API" value="${escapeHtml(value)}">
      <button type="button" class="icon-button toggle-key" title="Mostrar ou ocultar chave" aria-label="Mostrar ou ocultar chave">◉</button>
      <button type="button" class="icon-button remove-key" title="Remover chave" aria-label="Remover chave">×</button>`;
    row.querySelector('.toggle-key').addEventListener('click', () => {
      const input = row.querySelector('input');
      input.type = input.type === 'password' ? 'text' : 'password';
    });
    row.querySelector('.remove-key').addEventListener('click', () => {
      if (rows.children.length === 1) row.querySelector('input').value = '';
      else row.remove();
      renumberKeyRows();
      updateBudgetHint();
    });
    row.querySelector('input').addEventListener('input', () => { updateBudgetHint(); updateWorkflowState(); });
    rows.appendChild(row);
    updateBudgetHint();
    updateWorkflowState();
  }

  function renumberKeyRows() {
    $$('#apiKeysList .key-row').forEach((row, index) => {
      row.querySelector('.key-number').textContent = String(index + 1);
      const input = row.querySelector('input');
      input.id = `apiKey${index}`;
    });
  }

  function getApiKeys() {
    return [...new Set($$('.api-key-input').map((input) => input.value.trim()).filter(Boolean))];
  }

  function updateBudgetHint() {
    const count = getApiKeys().length;
    const defaultBudget = Core.DEFAULT_SEARCH_BUDGET_PER_KEY * count;
    $('#keysCount').textContent = String(count);
    $('#budgetHint').textContent = count
      ? `Sem valor manual, o limite local será ${numberFormat.format(defaultBudget)} títulos.`
      : 'Adicione pelo menos uma chave para pesquisar títulos ainda não armazenados no cache.';
    $('#addKeyButton').disabled = $('#apiKeysList').children.length >= Core.MAX_API_KEYS;
  }

  function getSearchOptions(apiKeyCount) {
    const publishedDate = $('#publishedAfter').value;
    const rawBudget = $('#searchBudget').value.trim();
    const budget = rawBudget === '' ? Core.DEFAULT_SEARCH_BUDGET_PER_KEY * apiKeyCount : Math.max(0, Number(rawBudget) || 0);
    return {
      maxResults: Math.max(1, Math.min(50, Number($('#maxResults').value) || 15)),
      region: ($('#region').value || 'BR').trim().toUpperCase(),
      language: ($('#language').value || 'pt').trim().toLowerCase(),
      publishedAfter: publishedDate ? `${publishedDate}T00:00:00Z` : null,
      searchBudget: budget,
      cacheTtlDays: Number($('#cacheTtl').value),
      refresh: $('#refreshCache').checked,
    };
  }

  function getEvaluationConfig() {
    const rules = Core.parseChannelRules($('#authorizedChannels').value);
    return {
      profile: $('#profile').value,
      threshold: Number($('#threshold').value),
      longDurationMinutes: Number($('#longDuration').value),
      authorizedIds: rules.ids,
      authorizedNames: rules.names,
      suspiciousChannelTerms: Core.parseTerms($('#suspiciousTerms').value),
      includeSummaries: $('#includeSummaries').checked,
    };
  }

  function cacheIsFresh(entry, ttlDays) {
    if (!entry || !entry.fetchedAt) return false;
    if (ttlDays < 0) return true;
    const fetchedAt = new Date(entry.fetchedAt).getTime();
    if (!Number.isFinite(fetchedAt)) return false;
    return Date.now() - fetchedAt <= ttlDays * 86400000;
  }

  async function getCacheEntry(key) {
    if (state.cacheAvailable) {
      try { return await state.cacheDb.get(key); } catch (error) { log(`Cache local indisponível: ${error.message}`, 'warning'); }
    }
    return state.memoryCache.get(key) || null;
  }

  async function putCacheEntry(entry) {
    state.memoryCache.set(entry.key, entry);
    if (state.cacheAvailable) {
      try { await state.cacheDb.put(entry); } catch (error) { log(`Não foi possível salvar checkpoint local: ${error.message}`, 'warning'); }
    }
  }

  function normalizeCachedVideos(videos) {
    return (videos || []).map((video) => ({
      videoId: String(video.videoId || video.video_id || ''),
      title: String(video.title || ''),
      description: String(video.description || ''),
      channelId: String(video.channelId || video.channel_id || ''),
      channelTitle: String(video.channelTitle || video.channel_title || ''),
      publishedAt: String(video.publishedAt || video.published_at || ''),
      durationSeconds: Number(video.durationSeconds || video.duration_seconds || 0),
      viewCount: Number(video.viewCount || video.view_count || 0),
      searchRank: Number(video.searchRank || video.search_rank || 0),
      url: video.url || `https://www.youtube.com/watch?v=${video.videoId || video.video_id || ''}`,
    })).filter((video) => video.videoId);
  }

  function keyPoolEvent(event) {
    renderKeyStatus(event.snapshot);
    if (event.type === 'key-rotated') {
      log(`Chave ${event.from + 1} indisponível (${event.reason}); alternando para a chave ${event.to + 1}.`, 'warning');
      showToast(`Chave ${event.from + 1} indisponível; usando a chave ${event.to + 1}.`, 'warning');
    } else if (event.type === 'key-exhausted') {
      log(`Chave ${event.keyIndex + 1} indisponível. Não há outra chave disponível.`, 'error');
    } else if (event.type === 'retry') {
      log(`Instabilidade na API; repetindo ${event.label} (tentativa ${event.attempt}).`, 'warning');
    }
  }

  function renderKeyStatus(snapshot) {
    const container = $('#keyStatusList');
    container.innerHTML = '';
    if (!snapshot || !snapshot.configured) {
      container.innerHTML = '<span class="muted">Nenhuma chave ativa.</span>';
      return;
    }
    for (let index = 0; index < snapshot.configured; index += 1) {
      const chip = document.createElement('span');
      const unavailable = snapshot.unavailableIndices.includes(index);
      const used = snapshot.usedIndices.includes(index);
      const active = index === snapshot.currentIndex && !unavailable;
      chip.className = `key-chip ${unavailable ? 'key-unavailable' : active ? 'key-active' : used ? 'key-used' : 'key-unused'}`;
      chip.textContent = `Chave ${index + 1}`;
      chip.title = unavailable ? 'Indisponível nesta execução' : active ? 'Ativa' : used ? 'Já utilizada' : 'Ainda não utilizada';
      container.appendChild(chip);
    }
  }

  function createScanState(apiKeys, options) {
    const pending = new Set(state.books.map((book) => book.identity));
    return {
      running: false,
      paused: false,
      pauseRequested: false,
      cursor: 0,
      controller: null,
      options,
      resultsByBook: new Map(),
      pendingBooks: pending,
      apiSearchCalls: 0,
      cachedBooks: 0,
      quotaBlocked: false,
      startedAt: new Date(),
      keyPool: new Core.YouTubeKeyPool(apiKeys, { onEvent: keyPoolEvent }),
    };
  }

  async function processBook(scan, book) {
    const options = scan.options;
    const key = Core.cacheKey(book, options);
    const entry = await getCacheEntry(key);
    const fresh = entry && !options.refresh && cacheIsFresh(entry, options.cacheTtlDays);
    if (fresh) {
      scan.resultsByBook.set(book.identity, normalizeCachedVideos(entry.videos));
      scan.cachedBooks += 1;
      scan.pendingBooks.delete(book.identity);
      return;
    }

    const budgetExhausted = options.searchBudget > 0 && scan.apiSearchCalls >= options.searchBudget;
    if (budgetExhausted || scan.quotaBlocked) {
      if (entry) {
        scan.resultsByBook.set(book.identity, normalizeCachedVideos(entry.videos));
        scan.cachedBooks += 1;
        scan.pendingBooks.delete(book.identity);
      }
      return;
    }

    try {
      const videos = await Core.searchYouTube(scan.keyPool, book, options, scan.controller.signal);
      scan.apiSearchCalls += 1;
      scan.resultsByBook.set(book.identity, videos);
      scan.pendingBooks.delete(book.identity);
      await putCacheEntry({
        key,
        version: Core.CACHE_VERSION,
        bookIdentity: book.identity,
        book: { title: book.title, author: book.author },
        query: Core.buildQuery(book),
        searchOptions: {
          maxResults: options.maxResults, region: options.region,
          language: options.language, publishedAfter: options.publishedAfter,
        },
        fetchedAt: new Date().toISOString(),
        videos,
      });
    } catch (error) {
      if (error && error.name === 'AbortError') throw error;
      if (error instanceof Core.QuotaReachedError) {
        scan.quotaBlocked = true;
        if (entry) {
          scan.resultsByBook.set(book.identity, normalizeCachedVideos(entry.videos));
          scan.cachedBooks += 1;
          scan.pendingBooks.delete(book.identity);
        }
        return;
      }
      throw error;
    }
  }

  async function runScan({ resume = false } = {}) {
    if (!state.books.length) {
      showToast('Carregue uma planilha antes de iniciar.', 'warning');
      return;
    }
    const apiKeys = getApiKeys();
    if (!resume && !apiKeys.length) {
      showToast('Adicione pelo menos uma chave da API do YouTube.', 'warning');
      $('#apiKeysSection').scrollIntoView({ behavior: 'smooth', block: 'start' });
      return;
    }

    let scan = state.scan;
    if (!resume || !scan) {
      const options = getSearchOptions(apiKeys.length);
      scan = createScanState(apiKeys, options);
      state.scan = scan;
      state.resultsByBook = scan.resultsByBook;
      state.pendingBooks = scan.pendingBooks;
      log(`Monitoramento iniciado para ${state.books.length} títulos.`, 'info');
    }

    scan.running = true;
    scan.paused = false;
    scan.pauseRequested = false;
    scan.controller = new AbortController();
    setRunUi('running');
    setStatus('Monitorando', 'working');
    renderKeyStatus(scan.keyPool.snapshot());

    try {
      while (scan.cursor < state.books.length) {
        if (scan.pauseRequested) break;
        const book = state.books[scan.cursor];
        $('#currentBook').textContent = `${book.title}${book.author ? ` — ${book.author}` : ''}`;
        try {
          await processBook(scan, book);
        } catch (error) {
          if (error && error.name === 'AbortError' && scan.pauseRequested) break;
          throw error;
        }
        scan.cursor += 1;
        updateProgress();
        if (scan.cursor % 5 === 0 || scan.cursor === state.books.length) refreshAnalysis();
        await new Promise((resolve) => setTimeout(resolve, 0));
      }

      state.resultsByBook = scan.resultsByBook;
      state.pendingBooks = scan.pendingBooks;
      refreshAnalysis();
      if (scan.cursor >= state.books.length) {
        scan.running = false;
        scan.paused = false;
        setRunUi('complete');
        setStatus(scan.pendingBooks.size ? 'Concluído com pendências' : 'Concluído', scan.pendingBooks.size ? 'warning' : 'success');
        log(`Monitoramento concluído: ${scan.apiSearchCalls} buscas novas, ${scan.cachedBooks} títulos do cache e ${scan.pendingBooks.size} pendentes.`, scan.pendingBooks.size ? 'warning' : 'success');
        showToast(scan.pendingBooks.size ? `Concluído com ${scan.pendingBooks.size} títulos pendentes.` : 'Monitoramento concluído.', scan.pendingBooks.size ? 'warning' : 'success', 7000);
        updateWorkflowState();
      } else {
        scan.running = false;
        scan.paused = true;
        setRunUi('paused');
        setStatus('Pausado', 'warning');
        log(`Monitoramento pausado em ${scan.cursor}/${state.books.length}.`, 'warning');
      }
    } catch (error) {
      console.error(error);
      scan.running = false;
      scan.paused = true;
      setRunUi('paused');
      setStatus('Erro na consulta', 'error');
      log(`Erro: ${error.message || error}`, 'error');
      showToast(error.message || 'Falha ao consultar o YouTube.', 'error', 10000);
    }
  }

  function pauseScan() {
    const scan = state.scan;
    if (!scan || !scan.running) return;
    scan.pauseRequested = true;
    if (scan.controller) scan.controller.abort();
    $('#pauseButton').disabled = true;
    setStatus('Pausando', 'working');
  }

  function resetScan() {
    if (state.scan && state.scan.running) {
      showToast('Pause a execução antes de reiniciar.', 'warning');
      return;
    }
    state.scan = null;
    state.resultsByBook = new Map();
    state.pendingBooks = new Set(state.books.map((book) => book.identity));
    state.analysis = null;
    $('#progressBar').style.width = '0%';
    $('#progressLabel').textContent = `0 / ${state.books.length}`;
    $('#currentBook').textContent = 'Aguardando início';
    renderKeyStatus(null);
    refreshAnalysis();
    setRunUi('idle');
    setStatus('Pronto', 'neutral');
    log('Execução reiniciada. O cache local foi preservado.', 'info');
  }

  function setRunUi(mode) {
    const start = $('#startButton');
    const pause = $('#pauseButton');
    const resume = $('#resumeButton');
    const reset = $('#resetButton');
    start.hidden = mode !== 'idle' && mode !== 'complete';
    start.textContent = mode === 'complete' ? 'Executar novamente' : 'Iniciar monitoramento';
    pause.hidden = mode !== 'running';
    pause.disabled = false;
    resume.hidden = mode !== 'paused';
    reset.disabled = mode === 'running';
    $$('.configuration-control').forEach((control) => { control.disabled = mode === 'running'; });
    $$('#apiKeysList button, #addKeyButton, #importKeysButton').forEach((control) => { control.disabled = mode === 'running'; });
  }

  function updateProgress() {
    const scan = state.scan;
    if (!scan) return;
    const total = state.books.length || 1;
    const percent = Math.min(100, (scan.cursor / total) * 100);
    $('#progressBar').style.width = `${percent}%`;
    $('#progressLabel').textContent = `${numberFormat.format(scan.cursor)} / ${numberFormat.format(state.books.length)}`;
    $('#apiCallsCount').textContent = numberFormat.format(scan.apiSearchCalls);
    $('#cacheHitsCount').textContent = numberFormat.format(scan.cachedBooks);
    $('#pendingCount').textContent = numberFormat.format(scan.pendingBooks.size);
    $('#rotationCount').textContent = numberFormat.format(scan.keyPool.rotations);
  }

  function refreshAnalysis() {
    const scanState = state.scan || {
      resultsByBook: state.resultsByBook,
      pendingBooks: state.pendingBooks,
    };
    state.analysis = Core.buildAnalysis(state.books, scanState, getEvaluationConfig());
    renderSummaryCards();
    renderResults();
  }

  function renderSummaryCards() {
    const summary = state.analysis ? state.analysis.summary : {
      total: state.books.length, processed: 0, pending: state.books.length,
      suspiciousTitles: 0, cleanTitles: 0, suspiciousVideos: 0,
    };
    $('#summaryTotal').textContent = numberFormat.format(summary.total);
    $('#summaryProcessed').textContent = numberFormat.format(summary.processed);
    $('#summarySuspiciousTitles').textContent = numberFormat.format(summary.suspiciousTitles);
    $('#summarySuspiciousVideos').textContent = numberFormat.format(summary.suspiciousVideos);
    $('#summaryPending').textContent = numberFormat.format(summary.pending);
  }

  function filteredVideos() {
    if (!state.analysis) return [];
    const text = Core.normalize($('#resultsSearch').value);
    const minScore = Number($('#minScoreFilter').value || 0);
    const sort = $('#sortResults').value;
    const items = state.analysis.acceptedEvaluations.filter((item) => {
      if (item.score < minScore) return false;
      if (!text) return true;
      const haystack = Core.normalize([
        item.book.title, item.book.author, item.video.title, item.video.channelTitle,
        item.reasons.join(' '),
      ].join(' '));
      return haystack.includes(text);
    });
    items.sort((a, b) => {
      if (sort === 'score') return b.score - a.score || b.video.viewCount - a.video.viewCount;
      if (sort === 'book') return a.book.title.localeCompare(b.book.title, 'pt-BR') || b.video.viewCount - a.video.viewCount;
      return b.video.viewCount - a.video.viewCount || b.score - a.score;
    });
    return items;
  }

  function scoreClass(score) {
    if (score >= 0.75) return 'score-high';
    if (score >= 0.5) return 'score-medium';
    return 'score-low';
  }

  function renderResults() {
    $$('.result-tab').forEach((button) => button.classList.toggle('is-active', button.dataset.tab === state.resultTab));
    const container = $('#resultsTableContainer');
    const empty = $('#resultsEmpty');
    if (!state.books.length) {
      container.innerHTML = '';
      empty.hidden = false;
      empty.querySelector('strong').textContent = 'Carregue um catálogo para começar.';
      empty.querySelector('span').textContent = 'Os resultados aparecerão aqui durante o monitoramento.';
      $('#pagination').hidden = true;
      return;
    }
    empty.hidden = true;
    if (state.resultTab === 'videos') renderVideosTable(container);
    else if (state.resultTab === 'ranking') renderRankingTable(container);
    else renderFullCatalogTable(container);
  }

  function renderVideosTable(container) {
    const items = filteredVideos();
    const pageSize = 50;
    const pages = Math.max(1, Math.ceil(items.length / pageSize));
    state.resultPage = Math.min(Math.max(1, state.resultPage), pages);
    const start = (state.resultPage - 1) * pageSize;
    const pageItems = items.slice(start, start + pageSize);
    const rows = pageItems.map((item) => `
      <tr>
        <td><strong>${escapeHtml(item.book.title)}</strong><small>${escapeHtml(item.book.author)}</small></td>
        <td>${escapeHtml(item.video.title)}</td>
        <td>${escapeHtml(item.video.channelTitle)}</td>
        <td class="numeric">${numberFormat.format(item.video.viewCount)}</td>
        <td>${Core.formatDuration(item.video.durationSeconds)}</td>
        <td><span class="score-pill ${scoreClass(item.score)}">${item.score.toFixed(2)}</span></td>
        <td class="reasons-cell">${escapeHtml(item.reasons.join(' · '))}</td>
        <td><a class="table-link" href="${escapeHtml(item.video.url)}" target="_blank" rel="noopener noreferrer">Abrir ↗</a></td>
      </tr>`).join('');
    container.innerHTML = `<table class="results-table"><thead><tr><th>Obra</th><th>Vídeo</th><th>Canal</th><th>Views</th><th>Duração</th><th>Score</th><th>Razões</th><th>Link</th></tr></thead><tbody>${rows || '<tr><td colspan="8" class="table-empty">Nenhum vídeo corresponde aos filtros atuais.</td></tr>'}</tbody></table>`;
    renderPagination(items.length, pages);
  }

  function renderRankingTable(container) {
    const rows = (state.analysis ? state.analysis.rankedBooks : []).map((item, index) => `
      <tr>
        <td class="numeric">${index + 1}</td>
        <td><strong>${escapeHtml(item.book.title)}</strong></td>
        <td>${escapeHtml(item.book.author)}</td>
        <td class="numeric">${numberFormat.format(item.accepted.length)}</td>
        <td class="numeric">${numberFormat.format(item.mostViewed.video.viewCount)}</td>
        <td><span class="score-pill ${scoreClass(item.maxScore)}">${item.maxScore.toFixed(2)}</span></td>
        <td><a class="table-link" href="${escapeHtml(item.mostViewed.video.url)}" target="_blank" rel="noopener noreferrer">Abrir ↗</a></td>
      </tr>`).join('');
    container.innerHTML = `<table class="results-table"><thead><tr><th>Posição</th><th>Título</th><th>Autor</th><th>Suspeitos</th><th>Maior nº de views</th><th>Score</th><th>Vídeo mais visto</th></tr></thead><tbody>${rows || '<tr><td colspan="7" class="table-empty">Ainda não há títulos classificados como suspeitos.</td></tr>'}</tbody></table>`;
    $('#pagination').hidden = true;
  }

  function renderFullCatalogTable(container) {
    const acceptedByBook = state.analysis ? state.analysis.acceptedByBook : new Map();
    const rows = state.books.map((book) => {
      const accepted = acceptedByBook.get(book.identity) || [];
      const pending = state.pendingBooks.has(book.identity);
      const status = accepted.length ? 'SIM' : pending ? 'pendente' : 'não';
      const statusClass = accepted.length ? 'catalog-yes' : pending ? 'catalog-pending' : 'catalog-no';
      const mostViewed = accepted.length ? accepted.reduce((best, item) => !best || item.video.viewCount > best.video.viewCount ? item : best, null) : null;
      return `<tr><td><strong>${escapeHtml(book.title)}</strong></td><td>${escapeHtml(book.author)}</td><td><span class="catalog-status ${statusClass}">${status}</span></td><td class="numeric">${accepted.length}</td><td>${mostViewed ? `<a class="table-link" href="${escapeHtml(mostViewed.video.url)}" target="_blank" rel="noopener noreferrer">Abrir ↗</a>` : '—'}</td></tr>`;
    }).join('');
    container.innerHTML = `<table class="results-table"><thead><tr><th>Título</th><th>Autor</th><th>Suspeita</th><th>Qtd.</th><th>Mais visto</th></tr></thead><tbody>${rows}</tbody></table>`;
    $('#pagination').hidden = true;
  }

  function renderPagination(total, pages) {
    const pagination = $('#pagination');
    pagination.hidden = total <= 50;
    $('#pageLabel').textContent = `Página ${state.resultPage} de ${pages} · ${numberFormat.format(total)} resultados`;
    $('#prevPage').disabled = state.resultPage <= 1;
    $('#nextPage').disabled = state.resultPage >= pages;
  }

  function styleForScore(score, alt = false) {
    if (score >= 0.75) return alt ? 'scoreRedAlt' : 'scoreRed';
    if (score >= 0.5) return alt ? 'scoreOrangeAlt' : 'scoreOrange';
    return alt ? 'scoreGreenAlt' : 'scoreGreen';
  }

  function reportSheets() {
    const analysis = state.analysis || Core.buildAnalysis(state.books, {
      resultsByBook: state.resultsByBook,
      pendingBooks: state.pendingBooks,
    }, getEvaluationConfig());
    const scan = state.scan;
    const generatedAt = new Date();
    const summaryRows = [
      [{ value: 'Relatório de Monitoramento de Pirataria — YouTube', style: 'title' }, { value: '' }],
      [{ value: `Gerado em: ${generatedAt.toLocaleString('pt-BR')}`, style: 'subtitle' }, { value: '' }],
      [{ value: '' }, { value: '' }],
    ];
    const summaryItems = [
      ['Total de títulos no catálogo', analysis.summary.total, 'summaryValue'],
      ['Títulos processados', analysis.summary.processed, 'summaryValue'],
      ['Títulos pendentes', analysis.summary.pending, 'summaryValue'],
      ['Títulos com suspeita de pirataria', analysis.summary.suspiciousTitles, 'summaryValue'],
      ['Títulos sem suspeita', analysis.summary.cleanTitles, 'summaryValue'],
      ['% dos processados com pirataria suspeita', analysis.summary.percentage, 'summaryPercent'],
      ['Consultas search.list nesta execução', scan ? scan.apiSearchCalls : 0, 'summaryValue'],
      ['Títulos reaproveitados do cache', scan ? scan.cachedBooks : 0, 'summaryValue'],
      ['Chaves de API configuradas', scan ? scan.keyPool.configuredCount : getApiKeys().length, 'summaryValue'],
      ['Chaves de API utilizadas', scan ? scan.keyPool.usedCount : 0, 'summaryValue'],
      ['Chaves indisponíveis nesta execução', scan ? scan.keyPool.unavailableCount : 0, 'summaryValue'],
      ['Trocas automáticas de chave', scan ? scan.keyPool.rotations : 0, 'summaryValue'],
      ['Perfil de classificação', $('#profile').value, 'summaryValue'],
    ];
    summaryItems.forEach(([label, value, style], index) => {
      summaryRows.push([
        { value: label, style: index % 2 === 0 ? 'summaryLabel' : 'default' },
        { value, style },
      ]);
    });

    const rankingRows = [[
      'Posição', 'Título', 'Autor', 'Qtd suspeitos', 'Maior nº de views',
      'Score pirataria', 'URL mais acessada', 'Razões',
    ].map((value) => ({ value, style: 'header' }))];
    analysis.rankedBooks.forEach((item, index) => {
      const alt = (index + 2) % 2 === 0;
      rankingRows.push([
        { value: index + 1, style: alt ? 'integerAlt' : 'integer' },
        { value: item.book.title, style: alt ? 'textAlt' : 'default' },
        { value: item.book.author, style: alt ? 'textAlt' : 'default' },
        { value: item.accepted.length, style: alt ? 'integerAlt' : 'integer' },
        { value: item.mostViewed.video.viewCount, style: alt ? 'integerAlt' : 'integer' },
        { value: item.maxScore, style: styleForScore(item.maxScore, alt) },
        { value: '▶ Abrir vídeo', style: alt ? 'hyperlinkAlt' : 'hyperlink', link: item.mostViewed.video.url },
        { value: item.mostViewed.reasons.join(' | '), style: alt ? 'wrapAlt' : 'wrap' },
      ]);
    });

    const videosRows = [[
      'Título do livro', 'Autor', 'Título do vídeo', 'Canal', 'Views', 'Duração',
      'Score', 'Razões', 'Data publicação', 'URL',
    ].map((value) => ({ value, style: 'header' }))];
    for (const ranked of analysis.rankedBooks) {
      const accepted = [...ranked.accepted].sort((a, b) => b.video.viewCount - a.video.viewCount || b.score - a.score || a.video.searchRank - b.video.searchRank);
      for (const evaluation of accepted) {
        const rowIndex = videosRows.length + 1;
        const alt = rowIndex % 2 === 0;
        videosRows.push([
          { value: ranked.book.title, style: alt ? 'textAlt' : 'default' },
          { value: ranked.book.author, style: alt ? 'textAlt' : 'default' },
          { value: evaluation.video.title, style: alt ? 'wrapAlt' : 'wrap' },
          { value: evaluation.video.channelTitle, style: alt ? 'wrapAlt' : 'wrap' },
          { value: evaluation.video.viewCount, style: alt ? 'integerAlt' : 'integer' },
          { value: Core.formatDuration(evaluation.video.durationSeconds), style: alt ? 'textAlt' : 'default' },
          { value: evaluation.score, style: styleForScore(evaluation.score, alt) },
          { value: evaluation.reasons.join(' | '), style: alt ? 'wrapAlt' : 'wrap' },
          { value: evaluation.video.publishedAt.slice(0, 10), style: alt ? 'dateTextAlt' : 'dateText' },
          { value: '▶ Abrir', style: alt ? 'hyperlinkAlt' : 'hyperlink', link: evaluation.video.url },
        ]);
      }
    }

    const catalogRows = [[
      'Título', 'Autor', 'Suspeita de pirataria', 'Qtd suspeitos', 'URL mais visto',
    ].map((value) => ({ value, style: 'header' }))];
    state.books.forEach((book, index) => {
      const accepted = analysis.acceptedByBook.get(book.identity) || [];
      const pending = state.pendingBooks.has(book.identity);
      const status = accepted.length ? 'SIM' : pending ? 'pendente' : 'não';
      const mostViewed = accepted.length ? accepted.reduce((best, item) => !best || item.video.viewCount > best.video.viewCount ? item : best, null) : null;
      const alt = (index + 2) % 2 === 0;
      const statusStyle = accepted.length ? 'statusRed' : pending ? 'statusOrange' : 'statusGreen';
      catalogRows.push([
        { value: book.title, style: alt ? 'textAlt' : 'default' },
        { value: book.author, style: alt ? 'textAlt' : 'default' },
        { value: status, style: statusStyle },
        { value: accepted.length, style: alt ? 'integerAlt' : 'integer' },
        mostViewed
          ? { value: '▶ Abrir', style: alt ? 'hyperlinkAlt' : 'hyperlink', link: mostViewed.video.url }
          : { value: '', style: alt ? 'textAlt' : 'default' },
      ]);
    });

    return [
      { name: 'Resumo', rows: summaryRows, columns: [49, 22], merges: ['A1:B1'], rowHeights: { 1: 28 }, freezeRows: 0, autoFilter: false },
      { name: 'Ranking por Views', rows: rankingRows, columns: [10, 52, 31, 15, 19, 16, 20, 72], rowHeights: { 1: 36 }, freezeRows: 1, autoFilter: true },
      { name: 'Vídeos Suspeitos', rows: videosRows, columns: [40, 28, 48, 28, 14, 11, 11, 64, 16, 14], rowHeights: { 1: 36 }, freezeRows: 1, autoFilter: true },
      { name: 'Catálogo Completo', rows: catalogRows, columns: [58, 34, 23, 16, 18], rowHeights: { 1: 36 }, freezeRows: 1, autoFilter: true },
    ];
  }

  async function exportReport() {
    if (!state.books.length) return;
    const button = $('#exportReportButton');
    button.disabled = true;
    const original = button.textContent;
    button.textContent = 'Gerando Excel…';
    try {
      refreshAnalysis();
      const blob = await Xlsx.createWorkbookBlob(reportSheets(), {
        title: 'Relatório de monitoramento de audiolivros no YouTube',
        creator: 'Monitor de Audiolivros no YouTube',
      });
      const stamp = new Date().toISOString().slice(0, 16).replace(/[-:T]/g, '');
      downloadBlob(blob, `relatorio_pirataria_youtube_${stamp}.xlsx`);
      log('Relatório Excel gerado.', 'success');
      showToast('Relatório Excel gerado.', 'success');
    } catch (error) {
      console.error(error);
      showToast(`Não foi possível gerar o Excel: ${error.message || error}`, 'error', 8000);
    } finally {
      button.disabled = false;
      button.textContent = original;
    }
  }

  async function exportCheckpoint() {
    let entries = [...state.memoryCache.values()];
    if (state.cacheAvailable) {
      try { entries = await state.cacheDb.getAll(); } catch (_) { /* usa memória */ }
    }
    const payload = {
      application: 'monitor-youtube-audiobooks-web',
      version: Core.CACHE_VERSION,
      exportedAt: new Date().toISOString(),
      entries,
    };
    downloadBlob(new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' }), `checkpoint_youtube_${new Date().toISOString().slice(0, 10)}.json`);
    showToast(`${entries.length} entradas de cache exportadas.`, 'success');
  }

  async function importCheckpoint(file) {
    try {
      const payload = JSON.parse(await file.text());
      if (!payload || !Array.isArray(payload.entries)) throw new Error('Formato de checkpoint não reconhecido.');
      const valid = payload.entries.filter((entry) => entry && entry.key && Array.isArray(entry.videos));
      valid.forEach((entry) => state.memoryCache.set(entry.key, entry));
      if (state.cacheAvailable) await state.cacheDb.putMany(valid);
      showToast(`${valid.length} entradas importadas.`, 'success');
      log(`Checkpoint importado com ${valid.length} entradas.`, 'success');
    } catch (error) {
      showToast(`Falha ao importar checkpoint: ${error.message || error}`, 'error', 8000);
    }
  }

  async function clearCache() {
    if (!window.confirm('Apagar todo o cache de pesquisas salvo neste navegador?')) return;
    state.memoryCache.clear();
    if (state.cacheAvailable) await state.cacheDb.clear();
    showToast('Cache local apagado.', 'success');
    log('Cache local apagado.', 'warning');
  }

  function saveSettings() {
    try {
      const ids = [
        'profile', 'threshold', 'longDuration', 'maxResults', 'region', 'language',
        'publishedAfter', 'searchBudget', 'cacheTtl', 'refreshCache',
        'includeSummaries', 'authorizedChannels', 'suspiciousTerms',
      ];
      const settings = {};
      ids.forEach((id) => {
        const element = $(`#${id}`);
        settings[id] = element.type === 'checkbox' ? element.checked : element.value;
      });
      localStorage.setItem('monitor-youtube-settings-v1', JSON.stringify(settings));
    } catch (_) {
      // Alguns navegadores restringem localStorage em páginas file://.
    }
  }

  function loadSettings() {
    try {
      const settings = JSON.parse(localStorage.getItem('monitor-youtube-settings-v1') || '{}');
      Object.entries(settings).forEach(([id, value]) => {
        const element = $(`#${id}`);
        if (!element) return;
        if (element.type === 'checkbox') element.checked = Boolean(value);
        else element.value = value;
      });
    } catch (_) { /* ignora configurações antigas */ }
    $('#thresholdValue').textContent = Number($('#threshold').value).toFixed(2);
  }

  function updateWorkflowState() {
    $('#workflowCatalog').classList.toggle('is-complete', state.books.length > 0);
    $('#workflowKeys').classList.toggle('is-complete', getApiKeys().length > 0);
    $('#workflowRun').classList.toggle('is-complete', Boolean(state.scan && state.scan.cursor >= state.books.length));
  }

  function bindEvents() {
    const fileInput = $('#catalogFile');
    $('#dropZone').addEventListener('click', () => fileInput.click());
    $('#dropZone').addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); fileInput.click(); }
    });
    fileInput.addEventListener('change', () => handleCatalogFile(fileInput.files[0]));
    for (const eventName of ['dragenter', 'dragover']) {
      $('#dropZone').addEventListener(eventName, (event) => { event.preventDefault(); $('#dropZone').classList.add('is-dragging'); });
    }
    for (const eventName of ['dragleave', 'drop']) {
      $('#dropZone').addEventListener(eventName, (event) => { event.preventDefault(); $('#dropZone').classList.remove('is-dragging'); });
    }
    $('#dropZone').addEventListener('drop', (event) => handleCatalogFile(event.dataTransfer.files[0]));

    $('#sheetSelect').addEventListener('change', () => {
      const sheet = state.workbook.sheets[Number($('#sheetSelect').value) || 0];
      const detected = Core.detectHeader(sheet.rows || []) || { headerRow: 0, titleIndex: 0, authorIndex: 1 };
      $('#headerRowInput').value = String(detected.headerRow + 1);
      populateColumnSelectors(detected.titleIndex, detected.authorIndex);
    });
    $('#headerRowInput').addEventListener('change', () => populateColumnSelectors());
    $('#applyMappingButton').addEventListener('click', () => applyCatalogSelection());

    $('#addKeyButton').addEventListener('click', () => createKeyRow());
    $('#importKeysButton').addEventListener('click', () => $('#keysFile').click());
    $('#keysFile').addEventListener('change', async (event) => {
      const file = event.target.files[0];
      if (!file) return;
      const values = [];
      for (const rawLine of (await readTextFile(file)).split(/\r?\n/)) {
        const withoutComment = rawLine.split('#', 1)[0].trim();
        if (!withoutComment) continue;
        values.push(...withoutComment.split(/[\s,;]+/).map((value) => value.trim()).filter(Boolean));
      }
      $('#apiKeysList').innerHTML = '';
      [...new Set(values)].slice(0, Core.MAX_API_KEYS).forEach((value) => createKeyRow(value));
      if (!$('#apiKeysList').children.length) createKeyRow();
      updateWorkflowState();
      event.target.value = '';
    });

    $('#authorizedFileButton').addEventListener('click', () => $('#authorizedFile').click());
    $('#authorizedFile').addEventListener('change', async (event) => {
      if (event.target.files[0]) $('#authorizedChannels').value = await readTextFile(event.target.files[0]);
      event.target.value = '';
      refreshAnalysis();
    });
    $('#termsFileButton').addEventListener('click', () => $('#termsFile').click());
    $('#termsFile').addEventListener('change', async (event) => {
      if (event.target.files[0]) $('#suspiciousTerms').value = await readTextFile(event.target.files[0]);
      event.target.value = '';
      refreshAnalysis();
    });

    $('#threshold').addEventListener('input', () => {
      $('#thresholdValue').textContent = Number($('#threshold').value).toFixed(2);
      refreshAnalysis();
    });
    $$('.analysis-setting').forEach((element) => element.addEventListener('change', () => { saveSettings(); refreshAnalysis(); }));
    $$('.configuration-control').forEach((element) => element.addEventListener('change', saveSettings));
    $('#authorizedChannels').addEventListener('input', saveSettings);
    $('#suspiciousTerms').addEventListener('input', saveSettings);

    $('#startButton').addEventListener('click', () => runScan({ resume: false }));
    $('#pauseButton').addEventListener('click', pauseScan);
    $('#resumeButton').addEventListener('click', () => runScan({ resume: true }));
    $('#resetButton').addEventListener('click', resetScan);
    $('#exportReportButton').addEventListener('click', exportReport);
    $('#exportCheckpointButton').addEventListener('click', exportCheckpoint);
    $('#importCheckpointButton').addEventListener('click', () => $('#checkpointFile').click());
    $('#checkpointFile').addEventListener('change', (event) => {
      if (event.target.files[0]) importCheckpoint(event.target.files[0]);
      event.target.value = '';
    });
    $('#clearCacheButton').addEventListener('click', clearCache);

    $$('.result-tab').forEach((button) => button.addEventListener('click', () => {
      state.resultTab = button.dataset.tab;
      state.resultPage = 1;
      renderResults();
    }));
    for (const id of ['resultsSearch', 'minScoreFilter', 'sortResults']) {
      $(`#${id}`).addEventListener(id === 'resultsSearch' ? 'input' : 'change', () => { state.resultPage = 1; renderResults(); });
    }
    $('#prevPage').addEventListener('click', () => { state.resultPage -= 1; renderResults(); });
    $('#nextPage').addEventListener('click', () => { state.resultPage += 1; renderResults(); });

    $('#advancedToggle').addEventListener('click', () => {
      const panel = $('#advancedPanel');
      panel.hidden = !panel.hidden;
      $('#advancedToggle').setAttribute('aria-expanded', String(!panel.hidden));
    });
    $('#logToggle').addEventListener('click', () => {
      const panel = $('#activityLogPanel');
      panel.hidden = !panel.hidden;
      $('#logToggle').setAttribute('aria-expanded', String(!panel.hidden));
    });

    window.addEventListener('beforeunload', (event) => {
      if (state.scan && state.scan.running) {
        event.preventDefault();
        event.returnValue = '';
      }
    });
  }

  async function initializeCache() {
    state.cacheDb = new CacheDatabase();
    try {
      await state.cacheDb.open();
      state.cacheAvailable = true;
      $('#cacheAvailability').textContent = 'Cache local ativo';
      $('#cacheAvailability').className = 'cache-status cache-ok';
    } catch (error) {
      state.cacheAvailable = false;
      $('#cacheAvailability').textContent = 'Cache apenas nesta sessão';
      $('#cacheAvailability').className = 'cache-status cache-warning';
      log('O navegador não permitiu cache persistente. A execução continua, mas exporte o checkpoint antes de fechar.', 'warning');
    }
  }

  async function init() {
    const failures = Core.selfTest();
    if (failures.length) {
      console.error('Falhas no autoteste:', failures);
      setStatus('Falha interna', 'error');
      showToast('A aplicação falhou no autoteste interno. Reextraia o pacote.', 'error', 12000);
      return;
    }
    loadSettings();
    bindEvents();
    createKeyRow();
    await initializeCache();
    renderSummaryCards();
    renderResults();
    setRunUi('idle');
    setStatus('Pronto', 'neutral');
    $('#appVersion').textContent = `v${Core.VERSION}`;
    log('Aplicação pronta. Nenhuma chave ou catálogo é enviado a servidores próprios.', 'success');
  }

  document.addEventListener('DOMContentLoaded', init);
}());
