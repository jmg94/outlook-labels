/**
 * Gmail Labels for Outlook — Task Pane Application Logic
 * Uses Office.js Categories API (Mailbox 1.8+) to manage labels.
 *
 * Category filtering strategy:
 * masterCategories.getAsync() on legacy Mac Outlook returns categories from
 * ALL accounts (including shared mailboxes). The API has no filter parameter
 * and no account identifier on categories. Instead of probing (which is
 * unreliable and risky), we maintain a self-curating "known labels" list in
 * localStorage that grows organically:
 *   - Labels created through this add-in are tracked automatically
 *   - Labels found on emails you open are auto-imported (they must be yours)
 *   - A manual "Import from Outlook" dialog lets you pick which categories
 *     from the full master list are yours (one-time setup)
 */
(function () {
  'use strict';

  // --- Color map: Outlook CategoryColor presets -> display values ---
  var CATEGORY_COLORS = {
    Preset0:  { name: 'Red',            hex: '#E74856' },
    Preset1:  { name: 'Orange',         hex: '#FF8C00' },
    Preset2:  { name: 'Brown',          hex: '#847545' },
    Preset3:  { name: 'Yellow',         hex: '#FFD700' },
    Preset4:  { name: 'Green',          hex: '#10893E' },
    Preset5:  { name: 'Teal',           hex: '#038387' },
    Preset6:  { name: 'Olive',          hex: '#7E735F' },
    Preset7:  { name: 'Blue',           hex: '#0078D4' },
    Preset8:  { name: 'Purple',         hex: '#8764B8' },
    Preset9:  { name: 'Cranberry',      hex: '#A4262C' },
    Preset10: { name: 'Steel',          hex: '#617D8A' },
    Preset11: { name: 'Dark Steel',     hex: '#4A5459' },
    Preset12: { name: 'Gray',           hex: '#8F9497' },
    Preset13: { name: 'Dark Gray',      hex: '#626467' },
    Preset14: { name: 'Black',          hex: '#2D2D2D' },
    Preset15: { name: 'Dark Red',       hex: '#750B1C' },
    Preset16: { name: 'Dark Orange',    hex: '#CA5010' },
    Preset17: { name: 'Dark Brown',     hex: '#5D5341' },
    Preset18: { name: 'Dark Yellow',    hex: '#C19C00' },
    Preset19: { name: 'Dark Green',     hex: '#0B6A0B' },
    Preset20: { name: 'Dark Teal',      hex: '#025D5D' },
    Preset21: { name: 'Dark Olive',     hex: '#5C5C2E' },
    Preset22: { name: 'Dark Blue',      hex: '#004E8C' },
    Preset23: { name: 'Dark Purple',    hex: '#5C2D91' },
    Preset24: { name: 'Dark Cranberry', hex: '#6E0811' }
  };

  // --- Application state ---
  var state = {
    masterCategories: [],   // Curated list: only labels the user has chosen/created
    allApiCategories: [],   // Raw list from masterCategories.getAsync() (all accounts)
    itemCategories: [],
    searchQuery: '',
    searchResults: [],
    isAllLabelsExpanded: false,
    focusedResultIndex: -1,
    pendingDeleteLabel: null,
    statusTimer: null,
    isSharedMailbox: false,
    primaryEmail: ''
  };

  // --- DOM references ---
  var dom = {};

  function cacheDom() {
    dom.app = document.getElementById('app');
    dom.appliedList = document.getElementById('applied-labels-list');
    dom.noLabelsMsg = document.getElementById('no-labels-msg');
    dom.searchInput = document.getElementById('label-search');
    dom.searchResults = document.getElementById('search-results');
    dom.toggleAllBtn = document.getElementById('toggle-all-labels');
    dom.toggleArrow = document.getElementById('toggle-arrow');
    dom.allLabelsList = document.getElementById('all-labels-list');
    dom.labelCount = document.getElementById('label-count');
    dom.refreshBtn = document.getElementById('refresh-btn');
    dom.importBtn = document.getElementById('import-btn');
    dom.createOverlay = document.getElementById('create-overlay');
    dom.createDialog = document.getElementById('create-dialog');
    dom.newLabelName = document.getElementById('new-label-name');
    dom.colorPicker = document.getElementById('color-picker');
    dom.createCancel = document.getElementById('create-cancel');
    dom.createConfirm = document.getElementById('create-confirm');
    dom.deleteOverlay = document.getElementById('delete-overlay');
    dom.deleteMsg = document.getElementById('delete-msg');
    dom.deleteCancel = document.getElementById('delete-cancel');
    dom.deleteConfirm = document.getElementById('delete-confirm');
    dom.importOverlay = document.getElementById('import-overlay');
    dom.importList = document.getElementById('import-list');
    dom.importSelectAll = document.getElementById('import-select-all');
    dom.importCancel = document.getElementById('import-cancel');
    dom.importConfirm = document.getElementById('import-confirm');
    dom.statusBar = document.getElementById('status-bar');
    dom.loading = document.getElementById('loading');
    dom.unsupported = document.getElementById('unsupported');
    dom.noItem = document.getElementById('no-item');
    dom.currentLabels = document.getElementById('current-labels');
    dom.searchSection = document.getElementById('search-section');
    dom.allLabelsSection = document.getElementById('all-labels-section');
  }

  // --- Utility ---

  function escapeHtml(str) {
    var div = document.createElement('div');
    div.appendChild(document.createTextNode(str));
    return div.innerHTML;
  }

  function debounce(fn, delay) {
    var timer;
    return function () {
      var args = arguments;
      var ctx = this;
      clearTimeout(timer);
      timer = setTimeout(function () { fn.apply(ctx, args); }, delay);
    };
  }

  function getColorHex(colorEnum) {
    var info = CATEGORY_COLORS[colorEnum];
    return info ? info.hex : '#888888';
  }

  function showStatus(message, type) {
    if (state.statusTimer) clearTimeout(state.statusTimer);
    dom.statusBar.textContent = message;
    dom.statusBar.className = type;
    state.statusTimer = setTimeout(function () {
      dom.statusBar.className = 'hidden';
    }, 3000);
  }

  function showView(view) {
    dom.loading.classList.add('hidden');
    dom.unsupported.classList.add('hidden');
    dom.noItem.classList.add('hidden');
    dom.currentLabels.classList.add('hidden');
    dom.searchSection.classList.add('hidden');
    dom.allLabelsSection.classList.add('hidden');

    if (view === 'loading') {
      dom.loading.classList.remove('hidden');
    } else if (view === 'unsupported') {
      dom.unsupported.classList.remove('hidden');
    } else if (view === 'no-item') {
      dom.noItem.classList.remove('hidden');
    } else if (view === 'main') {
      dom.currentLabels.classList.remove('hidden');
      dom.searchSection.classList.remove('hidden');
      dom.allLabelsSection.classList.remove('hidden');
    }
  }

  // --- localStorage: Known labels list ---

  var STORAGE_KEY_PREFIX = 'outlook_labels_own_';
  var IMPORT_DONE_PREFIX = 'outlook_labels_imported_';

  function getStorageKey() {
    var email = (Office.context.mailbox.userProfile.emailAddress || 'unknown').toLowerCase();
    return STORAGE_KEY_PREFIX + email;
  }

  function getImportDoneKey() {
    var email = (Office.context.mailbox.userProfile.emailAddress || 'unknown').toLowerCase();
    return IMPORT_DONE_PREFIX + email;
  }

  function hasCompletedImport() {
    return localStorage.getItem(getImportDoneKey()) === '1';
  }

  function markImportDone() {
    try { localStorage.setItem(getImportDoneKey(), '1'); } catch (e) {}
  }

  function loadOwnCategoryNames() {
    try {
      var data = localStorage.getItem(getStorageKey());
      return data ? JSON.parse(data) : [];
    } catch (e) {
      return [];
    }
  }

  function saveOwnCategoryNames(names) {
    try {
      localStorage.setItem(getStorageKey(), JSON.stringify(names));
    } catch (e) { /* ignore quota errors */ }
  }

  function addOwnCategoryName(name) {
    var names = loadOwnCategoryNames();
    var lower = name.toLowerCase();
    for (var i = 0; i < names.length; i++) {
      if (names[i].toLowerCase() === lower) return; // already tracked
    }
    names.push(name);
    saveOwnCategoryNames(names);
  }

  function removeOwnCategoryName(name) {
    var names = loadOwnCategoryNames();
    var lower = name.toLowerCase();
    names = names.filter(function (n) { return n.toLowerCase() !== lower; });
    saveOwnCategoryNames(names);
  }

  function isOwnCategory(name) {
    var names = loadOwnCategoryNames();
    var lower = name.toLowerCase();
    for (var i = 0; i < names.length; i++) {
      if (names[i].toLowerCase() === lower) return true;
    }
    return false;
  }

  // --- Office.js Categories API wrappers ---

  function fetchAllApiCategories() {
    return new Promise(function (resolve, reject) {
      Office.context.mailbox.masterCategories.getAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          state.allApiCategories = result.value || [];
          resolve(state.allApiCategories);
        } else {
          reject(result.error);
        }
      });
    });
  }

  function buildMasterCategoriesFromOwn() {
    // Build the display list from our known-own names, cross-referenced
    // with the API list to get current colors
    var ownNames = loadOwnCategoryNames();
    var apiMap = {};
    state.allApiCategories.forEach(function (cat) {
      apiMap[cat.displayName.toLowerCase()] = cat;
    });

    var result = [];
    ownNames.forEach(function (name) {
      var apiCat = apiMap[name.toLowerCase()];
      if (apiCat) {
        result.push(apiCat); // Use API version (has current color)
      } else {
        // Category in our list but not in API — might have been deleted externally
        // Still show it so user can re-create or clean up
        result.push({ displayName: name, color: 'Preset7' });
      }
    });

    result.sort(function (a, b) {
      return a.displayName.localeCompare(b.displayName);
    });

    state.masterCategories = result;
    return result;
  }

  function loadMasterCategories() {
    return fetchAllApiCategories().then(function () {
      buildMasterCategoriesFromOwn();
      return state.masterCategories;
    });
  }

  function loadItemCategories() {
    return new Promise(function (resolve, reject) {
      var item = Office.context.mailbox.item;
      if (!item) {
        state.itemCategories = [];
        resolve([]);
        return;
      }
      item.categories.getAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          state.itemCategories = result.value || [];

          // Auto-import: categories on the user's own emails must be theirs
          if (!state.isSharedMailbox) {
            state.itemCategories.forEach(function (cat) {
              addOwnCategoryName(cat.displayName);
            });
            // Rebuild master list with any newly discovered categories
            buildMasterCategoriesFromOwn();
          }

          resolve(state.itemCategories);
        } else {
          reject(result.error);
        }
      });
    });
  }

  function addMasterCategory(displayName, colorPreset) {
    return new Promise(function (resolve, reject) {
      var newCat = [{ displayName: displayName, color: colorPreset }];
      Office.context.mailbox.masterCategories.addAsync(newCat, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          addOwnCategoryName(displayName);
          resolve();
        } else {
          reject(result.error);
        }
      });
    });
  }

  function deleteMasterCategory(displayName) {
    return new Promise(function (resolve, reject) {
      var done = false;
      var timer = setTimeout(function () {
        if (!done) {
          done = true;
          removeOwnCategoryName(displayName);
          reject({ message: 'Delete timed out \u2014 removed from list only' });
        }
      }, 5000);

      try {
        Office.context.mailbox.masterCategories.removeAsync([displayName], function (result) {
          if (done) return;
          done = true;
          clearTimeout(timer);
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            removeOwnCategoryName(displayName);
            resolve();
          } else {
            removeOwnCategoryName(displayName);
            reject(result.error || { message: 'Unknown error' });
          }
        });
      } catch (e) {
        if (!done) {
          done = true;
          clearTimeout(timer);
          removeOwnCategoryName(displayName);
          reject(e);
        }
      }
    });
  }

  function addLabelToItem(displayName) {
    return new Promise(function (resolve, reject) {
      Office.context.mailbox.item.categories.addAsync([displayName], function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(result.error);
        }
      });
    });
  }

  function removeLabelFromItem(displayName) {
    return new Promise(function (resolve, reject) {
      Office.context.mailbox.item.categories.removeAsync([displayName], function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(result.error);
        }
      });
    });
  }

  // --- Shared Mailbox Detection ---

  function checkSharedMailbox() {
    return new Promise(function (resolve) {
      var item = Office.context.mailbox.item;
      if (!item || !item.getSharedPropertiesAsync) {
        state.isSharedMailbox = false;
        resolve();
        return;
      }
      item.getSharedPropertiesAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
          var owner = result.value.owner || '';
          var userEmail = (Office.context.mailbox.userProfile.emailAddress || '').toLowerCase();
          state.primaryEmail = userEmail;
          state.isSharedMailbox = owner.toLowerCase() !== userEmail;
        } else {
          state.isSharedMailbox = false;
        }
        resolve();
      });
    });
  }

  function applySharedMailboxRestrictions() {
    if (state.isSharedMailbox) {
      dom.searchSection.classList.add('hidden');
      showStatus('Shared mailbox \u2014 view only', 'error');
    }
  }

  // --- Data Loading ---

  function loadAllData() {
    showView('loading');
    checkSharedMailbox()
      .then(function () {
        return Promise.all([loadMasterCategories(), loadItemCategories()]);
      })
      .then(function () {
        showView('main');
        applySharedMailboxRestrictions();
        renderAppliedLabels();
        renderAllLabels();
        updateLabelCount();

        // Auto-open import on first use (user hasn't done an import yet)
        if (!hasCompletedImport() && state.allApiCategories.length > 0 && !state.isSharedMailbox) {
          openImportDialog();
        }
      })
      .catch(function (error) {
        showView('main');
        showStatus('Error loading labels: ' + (error.message || error), 'error');
      });
  }

  // --- Rendering: Applied Labels ---

  function renderAppliedLabels() {
    dom.appliedList.innerHTML = '';

    if (state.itemCategories.length === 0) {
      dom.noLabelsMsg.classList.remove('hidden');
      return;
    }

    dom.noLabelsMsg.classList.add('hidden');

    state.itemCategories.forEach(function (cat) {
      var chip = document.createElement('div');
      chip.className = 'label-chip';
      var hex = getColorHex(cat.color);
      chip.style.backgroundColor = hex + '1A';
      chip.style.borderColor = hex;
      chip.style.color = hex;

      var nameSpan = document.createElement('span');
      nameSpan.className = 'chip-name';
      nameSpan.textContent = cat.displayName;

      chip.appendChild(nameSpan);

      if (!state.isSharedMailbox) {
        var removeBtn = document.createElement('button');
        removeBtn.className = 'chip-remove';
        removeBtn.textContent = '\u00D7';
        removeBtn.title = 'Remove ' + cat.displayName;
        removeBtn.addEventListener('click', function () {
          handleRemoveLabel(cat.displayName);
        });
        chip.appendChild(removeBtn);
      }

      dom.appliedList.appendChild(chip);
    });
  }

  function handleRemoveLabel(displayName) {
    removeLabelFromItem(displayName)
      .then(function () { return loadItemCategories(); })
      .then(function () {
        renderAppliedLabels();
        renderAllLabels();
        renderSearchResults();
      })
      .catch(function (err) {
        showStatus('Error removing label: ' + (err.message || err), 'error');
      });
  }

  // --- Rendering: Search Results ---

  function highlightMatch(text, matchRanges) {
    if (!matchRanges || matchRanges.length === 0) return escapeHtml(text);

    var merged = FuzzySearch.mergeRanges(matchRanges);
    var result = '';
    var lastIdx = 0;

    for (var i = 0; i < merged.length; i++) {
      var start = merged[i][0];
      var end = merged[i][1];
      result += escapeHtml(text.substring(lastIdx, start));
      result += '<mark>' + escapeHtml(text.substring(start, end)) + '</mark>';
      lastIdx = end;
    }
    result += escapeHtml(text.substring(lastIdx));
    return result;
  }

  function performSearch() {
    var query = state.searchQuery;
    state.searchResults = FuzzySearch.search(query, state.masterCategories);
    state.focusedResultIndex = -1;
    renderSearchResults();
  }

  function renderSearchResults() {
    dom.searchResults.innerHTML = '';

    var query = state.searchQuery.trim();
    if (!query) return;

    var results = state.searchResults;
    var exactMatchExists = FuzzySearch.hasExactMatch(query, state.masterCategories);

    results.forEach(function (result, index) {
      var row = document.createElement('div');
      row.className = 'search-result-row';
      row.setAttribute('data-index', String(index));

      var isApplied = isLabelApplied(result.category.displayName);
      if (isApplied) row.classList.add('already-applied');
      if (index === state.focusedResultIndex) row.classList.add('focused');

      var colorDot = document.createElement('span');
      colorDot.className = 'color-dot';
      colorDot.style.backgroundColor = getColorHex(result.category.color);

      var nameSpan = document.createElement('span');
      nameSpan.className = 'result-name';
      nameSpan.innerHTML = highlightMatch(result.category.displayName, result.matchRanges);

      var checkSpan = document.createElement('span');
      checkSpan.className = 'result-check';
      checkSpan.textContent = isApplied ? '\u2713' : '';

      row.appendChild(colorDot);
      row.appendChild(nameSpan);
      row.appendChild(checkSpan);

      row.addEventListener('click', function () {
        handleToggleLabel(result.category.displayName, isApplied);
      });

      dom.searchResults.appendChild(row);
    });

    // "Create new label" option
    if (!exactMatchExists && query.length > 0) {
      var createRow = document.createElement('div');
      createRow.className = 'search-result-row create-new';
      var totalIndex = results.length;
      createRow.setAttribute('data-index', String(totalIndex));
      if (totalIndex === state.focusedResultIndex) createRow.classList.add('focused');

      var plusIcon = document.createElement('span');
      plusIcon.className = 'create-icon';
      plusIcon.textContent = '+';

      var createText = document.createElement('span');
      createText.className = 'create-text';
      createText.textContent = 'Create \u201C' + query + '\u201D';

      createRow.appendChild(plusIcon);
      createRow.appendChild(createText);

      createRow.addEventListener('click', function () {
        openCreateDialog(query);
      });

      dom.searchResults.appendChild(createRow);
    }
  }

  function handleToggleLabel(displayName, isCurrentlyApplied) {
    var action = isCurrentlyApplied
      ? removeLabelFromItem(displayName)
      : addLabelToItem(displayName);

    action
      .then(function () { return loadItemCategories(); })
      .then(function () {
        renderAppliedLabels();
        renderAllLabels();
        renderSearchResults();
      })
      .catch(function (err) {
        showStatus('Error: ' + (err.message || err), 'error');
      });
  }

  function isLabelApplied(displayName) {
    for (var i = 0; i < state.itemCategories.length; i++) {
      if (state.itemCategories[i].displayName === displayName) return true;
    }
    return false;
  }

  // --- Rendering: All Labels ---

  function renderAllLabels() {
    dom.allLabelsList.innerHTML = '';

    state.masterCategories.forEach(function (cat) {
      var row = document.createElement('div');
      row.className = 'all-label-row';

      var colorDot = document.createElement('span');
      colorDot.className = 'color-dot';
      colorDot.style.backgroundColor = getColorHex(cat.color);

      var nameSpan = document.createElement('span');
      nameSpan.className = 'all-label-name';
      nameSpan.textContent = cat.displayName;

      var checkSpan = document.createElement('span');
      checkSpan.className = 'all-label-check';
      checkSpan.textContent = isLabelApplied(cat.displayName) ? '\u2713' : '';

      row.appendChild(colorDot);
      row.appendChild(nameSpan);
      row.appendChild(checkSpan);

      if (!state.isSharedMailbox) {
        var deleteBtn = document.createElement('button');
        deleteBtn.className = 'all-label-delete';
        deleteBtn.textContent = '\u00D7';
        deleteBtn.title = 'Delete label';
        deleteBtn.addEventListener('click', function (e) {
          e.stopPropagation();
          confirmDeleteLabel(cat.displayName);
        });
        row.appendChild(deleteBtn);

        row.addEventListener('click', function () {
          var applied = isLabelApplied(cat.displayName);
          handleToggleLabel(cat.displayName, applied);
        });
      }

      dom.allLabelsList.appendChild(row);
    });
  }

  function updateLabelCount() {
    dom.labelCount.textContent = String(state.masterCategories.length);
  }

  // --- Create Label Dialog ---

  function openCreateDialog(prefillName) {
    dom.newLabelName.value = prefillName || '';
    renderColorPicker();
    dom.createOverlay.classList.remove('hidden');
    dom.newLabelName.focus();
    dom.newLabelName.select();
  }

  function closeCreateDialog() {
    dom.createOverlay.classList.add('hidden');
    dom.newLabelName.value = '';
  }

  function renderColorPicker() {
    dom.colorPicker.innerHTML = '';
    var presets = Object.keys(CATEGORY_COLORS);

    presets.forEach(function (presetKey) {
      var info = CATEGORY_COLORS[presetKey];
      var swatch = document.createElement('button');
      swatch.type = 'button';
      swatch.className = 'color-swatch';
      swatch.style.backgroundColor = info.hex;
      swatch.title = info.name;
      swatch.setAttribute('data-preset', presetKey);

      swatch.addEventListener('click', function () {
        dom.colorPicker.querySelectorAll('.color-swatch').forEach(function (s) {
          s.classList.remove('selected');
        });
        swatch.classList.add('selected');
      });

      dom.colorPicker.appendChild(swatch);
    });

    var defaultSwatch = dom.colorPicker.querySelector('[data-preset="Preset7"]');
    if (defaultSwatch) defaultSwatch.classList.add('selected');
  }

  function handleCreateConfirm() {
    var name = dom.newLabelName.value.trim();
    if (!name) {
      showStatus('Label name cannot be empty', 'error');
      return;
    }

    if (FuzzySearch.hasExactMatch(name, state.masterCategories)) {
      showStatus('A label with this name already exists', 'error');
      return;
    }

    var selectedSwatch = dom.colorPicker.querySelector('.color-swatch.selected');
    var presetKey = selectedSwatch ? selectedSwatch.getAttribute('data-preset') : 'Preset7';
    var colorEnum = Office.MailboxEnums.CategoryColor[presetKey];

    addMasterCategory(name, colorEnum)
      .then(function () { return fetchAllApiCategories(); })
      .then(function () {
        buildMasterCategoriesFromOwn();
        return addLabelToItem(name);
      })
      .then(function () { return loadItemCategories(); })
      .then(function () {
        closeCreateDialog();
        dom.searchInput.value = '';
        state.searchQuery = '';
        dom.searchResults.innerHTML = '';
        renderAppliedLabels();
        renderAllLabels();
        updateLabelCount();
        showStatus('Label \u201C' + name + '\u201D created and applied', 'success');
      })
      .catch(function (err) {
        showStatus('Error creating label: ' + (err.message || err), 'error');
      });
  }

  // --- Delete Label Dialog ---

  function confirmDeleteLabel(displayName) {
    state.pendingDeleteLabel = displayName;
    dom.deleteMsg.textContent = 'Delete \u201C' + displayName + '\u201D? This removes it from all emails.';
    dom.deleteOverlay.classList.remove('hidden');
  }

  function closeDeleteDialog() {
    dom.deleteOverlay.classList.add('hidden');
    state.pendingDeleteLabel = null;
  }

  function handleDeleteConfirm() {
    var name = state.pendingDeleteLabel;
    if (!name) return;

    deleteMasterCategory(name)
      .then(function () { return fetchAllApiCategories(); })
      .then(function () {
        buildMasterCategoriesFromOwn();
        return loadItemCategories();
      })
      .then(function () {
        closeDeleteDialog();
        renderAppliedLabels();
        renderAllLabels();
        updateLabelCount();
        renderSearchResults();
        showStatus('Label \u201C' + name + '\u201D deleted', 'success');
      })
      .catch(function (err) {
        closeDeleteDialog();
        buildMasterCategoriesFromOwn();
        renderAppliedLabels();
        renderAllLabels();
        updateLabelCount();
        renderSearchResults();
        var msg = (err && err.message) ? err.message : String(err || 'Unknown error');
        showStatus('Error deleting label: ' + msg, 'error');
      });
  }

  // --- Import Dialog ---

  function openImportDialog() {
    dom.importList.innerHTML = '';
    var ownNames = loadOwnCategoryNames();
    var ownLower = ownNames.map(function (n) { return n.toLowerCase(); });

    // Sort API categories alphabetically
    var sorted = state.allApiCategories.slice().sort(function (a, b) {
      return a.displayName.localeCompare(b.displayName);
    });

    if (sorted.length === 0) {
      var emptyMsg = document.createElement('p');
      emptyMsg.className = 'empty-state';
      emptyMsg.textContent = 'No categories found in Outlook.';
      dom.importList.appendChild(emptyMsg);
      dom.importOverlay.classList.remove('hidden');
      return;
    }

    sorted.forEach(function (cat) {
      var row = document.createElement('label');
      row.className = 'import-row';

      var checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.value = cat.displayName;
      checkbox.checked = ownLower.indexOf(cat.displayName.toLowerCase()) !== -1;

      var colorDot = document.createElement('span');
      colorDot.className = 'color-dot';
      colorDot.style.backgroundColor = getColorHex(cat.color);

      var nameSpan = document.createElement('span');
      nameSpan.className = 'import-name';
      nameSpan.textContent = cat.displayName;

      row.appendChild(checkbox);
      row.appendChild(colorDot);
      row.appendChild(nameSpan);
      dom.importList.appendChild(row);
    });

    // Update select-all state
    updateSelectAllState();

    dom.importOverlay.classList.remove('hidden');
  }

  function closeImportDialog() {
    dom.importOverlay.classList.add('hidden');
  }

  function updateSelectAllState() {
    var checkboxes = dom.importList.querySelectorAll('input[type="checkbox"]');
    var allChecked = true;
    checkboxes.forEach(function (cb) {
      if (!cb.checked) allChecked = false;
    });
    dom.importSelectAll.checked = allChecked && checkboxes.length > 0;
  }

  function handleSelectAllToggle() {
    var checked = dom.importSelectAll.checked;
    var checkboxes = dom.importList.querySelectorAll('input[type="checkbox"]');
    checkboxes.forEach(function (cb) {
      cb.checked = checked;
    });
  }

  function handleImportConfirm() {
    var checkboxes = dom.importList.querySelectorAll('input[type="checkbox"]');
    var selectedNames = [];
    checkboxes.forEach(function (cb) {
      if (cb.checked) {
        selectedNames.push(cb.value);
      }
    });

    saveOwnCategoryNames(selectedNames);
    markImportDone();
    buildMasterCategoriesFromOwn();

    closeImportDialog();
    renderAllLabels();
    updateLabelCount();
    renderSearchResults();
    showStatus('Imported ' + selectedNames.length + ' label' + (selectedNames.length !== 1 ? 's' : ''), 'success');
  }

  // --- Keyboard Navigation for Search ---

  function getTotalResultCount() {
    var count = state.searchResults.length;
    var query = state.searchQuery.trim();
    if (query && !FuzzySearch.hasExactMatch(query, state.masterCategories)) {
      count++;
    }
    return count;
  }

  function handleSearchKeydown(e) {
    var total = getTotalResultCount();
    if (total === 0) return;

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      state.focusedResultIndex = Math.min(state.focusedResultIndex + 1, total - 1);
      updateFocusedResult();
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      state.focusedResultIndex = Math.max(state.focusedResultIndex - 1, -1);
      updateFocusedResult();
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (state.focusedResultIndex >= 0 && state.focusedResultIndex < total) {
        var focusedRow = dom.searchResults.querySelector('[data-index="' + state.focusedResultIndex + '"]');
        if (focusedRow) focusedRow.click();
      }
    } else if (e.key === 'Escape') {
      dom.searchInput.value = '';
      state.searchQuery = '';
      state.focusedResultIndex = -1;
      dom.searchResults.innerHTML = '';
      dom.searchInput.blur();
    }
  }

  function updateFocusedResult() {
    var rows = dom.searchResults.querySelectorAll('.search-result-row');
    rows.forEach(function (row) { row.classList.remove('focused'); });
    if (state.focusedResultIndex >= 0) {
      var target = dom.searchResults.querySelector('[data-index="' + state.focusedResultIndex + '"]');
      if (target) {
        target.classList.add('focused');
        target.scrollIntoView({ block: 'nearest' });
      }
    }
  }

  // --- Event Binding ---

  function bindEvents() {
    dom.searchInput.addEventListener('input', debounce(function () {
      state.searchQuery = dom.searchInput.value;
      state.focusedResultIndex = -1;
      performSearch();
    }, 150));

    dom.searchInput.addEventListener('keydown', handleSearchKeydown);

    document.addEventListener('click', function (e) {
      if (!dom.searchInput.contains(e.target) && !dom.searchResults.contains(e.target)) {
        dom.searchResults.innerHTML = '';
        state.focusedResultIndex = -1;
      }
    });

    dom.searchInput.addEventListener('focus', function () {
      if (state.searchQuery.trim()) {
        performSearch();
      }
    });

    dom.toggleAllBtn.addEventListener('click', function () {
      state.isAllLabelsExpanded = !state.isAllLabelsExpanded;
      if (state.isAllLabelsExpanded) {
        dom.allLabelsList.classList.remove('collapsed');
        dom.allLabelsList.style.maxHeight = dom.allLabelsList.scrollHeight + 'px';
        dom.toggleArrow.classList.add('expanded');
      } else {
        dom.allLabelsList.style.maxHeight = '0';
        dom.allLabelsList.classList.add('collapsed');
        dom.toggleArrow.classList.remove('expanded');
      }
    });

    dom.refreshBtn.addEventListener('click', function () { loadAllData(); });

    // Import button
    dom.importBtn.addEventListener('click', function () { openImportDialog(); });

    // Create dialog
    dom.createCancel.addEventListener('click', closeCreateDialog);
    dom.createConfirm.addEventListener('click', handleCreateConfirm);
    dom.createOverlay.addEventListener('click', function (e) {
      if (e.target === dom.createOverlay) closeCreateDialog();
    });
    dom.newLabelName.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') handleCreateConfirm();
      if (e.key === 'Escape') closeCreateDialog();
    });

    // Delete dialog
    dom.deleteCancel.addEventListener('click', closeDeleteDialog);
    dom.deleteConfirm.addEventListener('click', handleDeleteConfirm);
    dom.deleteOverlay.addEventListener('click', function (e) {
      if (e.target === dom.deleteOverlay) closeDeleteDialog();
    });

    // Import dialog
    dom.importCancel.addEventListener('click', closeImportDialog);
    dom.importConfirm.addEventListener('click', handleImportConfirm);
    dom.importOverlay.addEventListener('click', function (e) {
      if (e.target === dom.importOverlay) closeImportDialog();
    });
    dom.importSelectAll.addEventListener('change', handleSelectAllToggle);
    dom.importList.addEventListener('change', function (e) {
      if (e.target.type === 'checkbox') updateSelectAllState();
    });
  }

  // --- Initialization ---

  Office.onReady(function (info) {
    cacheDom();

    if (info.host !== Office.HostType.Outlook) return;

    if (!Office.context.requirements.isSetSupported('Mailbox', '1.8')) {
      showView('unsupported');
      return;
    }

    if (!Office.context.mailbox.item) {
      showView('no-item');
      return;
    }

    bindEvents();
    loadAllData();
  });
})();
