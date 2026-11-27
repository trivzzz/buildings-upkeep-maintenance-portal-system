import { initializeApp } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-app.js";
    import { getAuth, signInAnonymously, signInWithCustomToken } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-auth.js";
    import { getFirestore, collection, doc, addDoc, updateDoc, deleteDoc, onSnapshot } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-firestore.js";

    // Global application state
    const state = {
      supplies: [],
      events: [],
      checklists: [],
      dtr: [],
      activeTab: 'supplies',
      supplyPage: 1,
      editContext: null,
    };

    let db = null;
    let auth = null;
    let currentUser = null;
    let useCloud = false;
    const appId = typeof __app_id !== 'undefined' ? __app_id : 'office-db-system';

    // Date & Time Utilities
    function getTodayISO() {
      const now = new Date();
      const offset = now.getTimezoneOffset();
      const local = new Date(now.getTime() - offset * 60 * 1000);
      return local.toISOString().split('T')[0];
    }

    function getCurrentMonthISO() {
      return getTodayISO().slice(0, 7);
    }

    function escapeHtml(str) {
      if (str === null || str === undefined) return '';
      return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
    }

    function formatTime12(time24) {
      if (!time24) return '--:--';
      const [hStr, mStr] = time24.split(':');
      let h = parseInt(hStr, 10);
      if (isNaN(h)) return time24;
      const ampm = h >= 12 ? 'PM' : 'AM';
      h = h % 12 || 12;
      return `${h}:${mStr} ${ampm}`;
    }

    function readTimePeriods(prefix, group) {
      return {
        [`${group}_morning_from`]: document.getElementById(`${prefix}_${group}_morning_from`).value,
        [`${group}_morning_to`]: document.getElementById(`${prefix}_${group}_morning_to`).value,
        [`${group}_afternoon_from`]: document.getElementById(`${prefix}_${group}_afternoon_from`).value,
        [`${group}_afternoon_to`]: document.getElementById(`${prefix}_${group}_afternoon_to`).value
      };
    }

    function validateTimePeriods(periods, group, groupLabel, requireOne) {
      const ranges = ['morning', 'afternoon'];
      let hasRange = false;
      for (const period of ranges) {
        const from = periods[`${group}_${period}_from`];
        const to = periods[`${group}_${period}_to`];
        if (from || to) {
          hasRange = true;
          if (!from || !to) return `${groupLabel} ${period} needs both a from and to time.`;
          if (from >= to) return `${groupLabel} ${period} end time must be after its start time.`;
        }
      }
      if (requireOne && !hasRange) return `Enter at least one ${groupLabel.toLowerCase()} time period.`;
      return '';
    }

    function validateTimeRange(from, to, label, requireOne) {
      if (!from && !to) return requireOne ? `Enter a ${label.toLowerCase()} time range.` : '';
      if (!from || !to) return `${label} needs both a from and to time.`;
      if (from >= to) return `${label} end time must be after its start time.`;
      return '';
    }

    function formatTimePeriods(record, group) {
      if (group === 'event' && (record.event_whole_day || record.event_period === 'Whole Day')) {
        const from = record.event_day_from;
        const to = record.event_day_to;
        return from || to
          ? `Whole Day (${from ? formatTime12(from) : '--:--'} – ${to ? formatTime12(to) : '--:--'})`
          : 'Whole Day';
      }
      const ranges = ['morning', 'afternoon'].flatMap(period => {
        const from = record[`${group}_${period}_from`];
        const to = record[`${group}_${period}_to`];
        if (!from && !to) return [];
        return [`${period.charAt(0).toUpperCase()}${period.slice(1)}: ${from ? formatTime12(from) : '--:--'} – ${to ? formatTime12(to) : '--:--'}`];
      });
      return ranges.length ? ranges.join('; ') : formatTime12(record[`${group}_time`]);
    }

    window.toggleWholeDayEventTimes = function(prefix) {
      const wholeDay = document.getElementById(`${prefix}_event_whole_day`);
      const periodFields = document.getElementById(`${prefix}_event_period_fields`);
      const wholeDayFields = document.getElementById(`${prefix}_event_whole_day_fields`);
      const inputs = periodFields.querySelectorAll('input[type="time"]');
      periodFields.classList.toggle('hidden', wholeDay.checked);
      wholeDayFields.classList.toggle('hidden', !wholeDay.checked);
      inputs.forEach(input => {
        input.disabled = wholeDay.checked;
        if (wholeDay.checked) input.value = '';
      });
      wholeDayFields.querySelectorAll('input[type="time"]').forEach(input => {
        input.disabled = !wholeDay.checked;
        if (!wholeDay.checked) input.value = '';
      });
    };

    window.toggleEventSetupFields = function(prefix) {
      const includeSetup = document.getElementById(`${prefix}_include_setup`).checked;
      const dateFields = document.getElementById(`${prefix}_setup_date_fields`);
      const timeFields = document.getElementById(`${prefix}_setup_time_fields`);
      const dateInput = document.getElementById(`${prefix}_setup_date`);

      dateFields.classList.toggle('hidden', !includeSetup);
      timeFields.classList.toggle('hidden', !includeSetup);
      dateInput.disabled = !includeSetup;
      dateInput.required = includeSetup;
      timeFields.querySelectorAll('input[type="time"]').forEach(input => {
        input.disabled = !includeSetup;
        if (!includeSetup) input.value = '';
      });
      if (!includeSetup) dateInput.value = '';
    };

    function renderTimePeriodFields(prefix, group, record, color) {
      const tone = color === 'indigo'
        ? { border: 'border-indigo-200 bg-indigo-50/30', label: 'text-indigo-700', input: 'border-indigo-300' }
        : { border: 'border-slate-200', label: 'text-slate-700', input: 'border-slate-300' };
      return ['morning', 'afternoon'].map(period => {
        const title = period.charAt(0).toUpperCase() + period.slice(1);
        const fromId = `${prefix}_${group}_${period}_from`;
        const toId = `${prefix}_${group}_${period}_to`;
        return `
          <fieldset class="rounded-lg border ${tone.border} p-3">
            <legend class="px-1 text-xs font-semibold ${tone.label}">${title}</legend>
            <div class="grid grid-cols-2 gap-2">
              <label class="text-xs text-slate-600">From<input type="time" id="${fromId}" value="${escapeHtml(record[`${group}_${period}_from`] || '')}" class="mt-1 w-full rounded-lg border ${tone.input} px-2 py-2 text-sm" /></label>
              <label class="text-xs text-slate-600">To<input type="time" id="${toId}" value="${escapeHtml(record[`${group}_${period}_to`] || '')}" class="mt-1 w-full rounded-lg border ${tone.input} px-2 py-2 text-sm" /></label>
            </div>
          </fieldset>
        `;
      }).join('');
    }

    // Toast notification helper
    function showToast(message, type = 'success') {
      const container = document.getElementById('toastContainer');
      const toast = document.createElement('div');
      const bgClass = type === 'error' ? 'bg-rose-600' : type === 'info' ? 'bg-[#1e3a8a]' : 'bg-emerald-600';
      toast.className = `${bgClass} text-white px-4 py-2.5 rounded-xl shadow-lg text-xs sm:text-sm font-semibold flex items-center gap-2 animate-fade-in pointer-events-auto`;
      toast.innerHTML = `<span>${escapeHtml(message)}</span>`;
      container.appendChild(toast);
      setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transition = 'opacity 0.25s ease';
        setTimeout(() => toast.remove(), 260);
      }, 2800);
    }

    // Initialize default form dates
    function initDefaultDates() {
      const today = getTodayISO();
      const month = getCurrentMonthISO();
      document.getElementById('sup_entry_date').value = today;
      document.getElementById('sup_date_received').value = today;
      document.getElementById('ev_date').value = today;
      document.getElementById('ev_setup_date').value = today;
      document.getElementById('chk_month').value = month;
      document.getElementById('dtr_date').value = today;
    }

    // Update Sync Status UI
    function setSyncStatus(mode, uid) {
      const badge = document.getElementById('syncStatusBadge');
      const dot = document.getElementById('syncStatusDot');
      const text = document.getElementById('syncStatusText');
      const footerTerm = document.getElementById('footerTerminalId');

      if (mode === 'cloud') {
        badge.className = 'inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-emerald-500/20 text-emerald-200 border border-emerald-400/30';
        dot.className = 'w-2 h-2 rounded-full bg-emerald-400 animate-pulse';
        text.textContent = 'Live Cloud Sync (Multi-PC)';
      } else {
        badge.className = 'inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-400/20 text-blue-100 border border-blue-300/30';
        dot.className = 'w-2 h-2 rounded-full bg-blue-300';
        text.textContent = 'Office Session Ready';
      }
      footerTerm.textContent = `Terminal ID: ${uid}`;
    }

    // Initialize Shared Database (Firebase Firestore with Anonymous/Custom Token Auth)
    async function initDatabase() {
      try {
        if (typeof __firebase_config !== 'undefined' && __firebase_config && __firebase_config.trim() !== '') {
          const firebaseConfig = JSON.parse(__firebase_config);
          const app = initializeApp(firebaseConfig);
          auth = getAuth(app);
          db = getFirestore(app);

          // RULE 3: Authenticate FIRST and await it before any Firestore queries
          if (typeof __initial_auth_token !== 'undefined' && __initial_auth_token) {
            await signInWithCustomToken(auth, __initial_auth_token);
          } else {
            await signInAnonymously(auth);
          }

          currentUser = auth.currentUser;
          if (!currentUser) throw new Error('Authentication did not return a user.');

          useCloud = true;
          setSyncStatus('cloud', currentUser.uid);

          // RULE 1 & RULE 2: Strict /artifacts/{appId}/public/data/{collectionName} path, no orderBy/where
          const suppliesCol = collection(db, 'artifacts', appId, 'public', 'data', 'office_supplies');
          const eventsCol = collection(db, 'artifacts', appId, 'public', 'data', 'office_events');
          const checklistsCol = collection(db, 'artifacts', appId, 'public', 'data', 'office_checklists');
          const dtrCol = collection(db, 'artifacts', appId, 'public', 'data', 'office_dtr');

          onSnapshot(suppliesCol, (snap) => {
            state.supplies = snap.docs.map(d => ({ id: d.id, ...d.data() }));
            renderAll();
          }, (err) => {
            console.error('Supplies snapshot error:', err);
            showToast('Sync warning on Supplies table.', 'error');
          });

          onSnapshot(eventsCol, (snap) => {
            state.events = snap.docs.map(d => ({ id: d.id, ...d.data() }));
            renderAll();
          }, (err) => {
            console.error('Events snapshot error:', err);
            showToast('Sync warning on Events table.', 'error');
          });

          onSnapshot(checklistsCol, (snap) => {
            state.checklists = snap.docs.map(d => ({ id: d.id, ...d.data() }));
            renderAll();
          }, (err) => {
            console.error('Checklists snapshot error:', err);
            showToast('Sync warning on Checklists table.', 'error');
          });

          onSnapshot(dtrCol, (snap) => {
            state.dtr = snap.docs.map(d => ({ id: d.id, ...d.data() }));
            renderAll();
          }, (err) => {
            console.error('DTR snapshot error:', err);
            showToast('Sync warning on DTR table.', 'error');
          });
        } else {
          // Fallback to clean local in-memory mode if opened offline without cloud config
          useCloud = false;
          currentUser = { uid: 'office-local-' + Math.random().toString(36).slice(2, 8) };
          setSyncStatus('local', currentUser.uid);
          renderAll();
        }
      } catch (error) {
        console.error('Database init fallback:', error);
        useCloud = false;
        currentUser = { uid: 'office-local-' + Math.random().toString(36).slice(2, 8) };
        setSyncStatus('local', currentUser.uid);
        renderAll();
      }
    }

    // Unified CRUD Operations
    async function createRecord(colName, stateKey, data) {
      const payload = { ...data, createdAt: Date.now() };
      if (useCloud && db && currentUser) {
        const colRef = collection(db, 'artifacts', appId, 'public', 'data', colName);
        await addDoc(colRef, payload);
      } else {
        const localId = 'loc_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6);
        state[stateKey].unshift({ id: localId, ...payload });
        renderAll();
      }
    }

    async function updateRecord(colName, stateKey, id, updates) {
      if (useCloud && db && currentUser) {
        const docRef = doc(db, 'artifacts', appId, 'public', 'data', colName, id);
        await updateDoc(docRef, updates);
      } else {
        const idx = state[stateKey].findIndex(item => item.id === id);
        if (idx !== -1) {
          state[stateKey][idx] = { ...state[stateKey][idx], ...updates };
          renderAll();
        }
      }
    }

    async function deleteRecord(colName, stateKey, id) {
      if (useCloud && db && currentUser) {
        const docRef = doc(db, 'artifacts', appId, 'public', 'data', colName, id);
        await deleteDoc(docRef);
      } else {
        state[stateKey] = state[stateKey].filter(item => item.id !== id);
        renderAll();
      }
    }

    // ==================== TAB NAVIGATION ====================
    window.switchTab = function(tabName) {
      state.activeTab = tabName;
      document.querySelectorAll('.module-panel').forEach(el => el.classList.add('hidden'));
      document.getElementById(`module-${tabName}`).classList.remove('hidden');

      const activeBtn = document.getElementById(`tab-btn-${tabName}`);
      if (activeBtn) {
        activeBtn.className = 'tab-nav-btn flex items-center justify-center gap-2 px-4 py-2.5 rounded-lg text-sm font-semibold transition cursor-pointer bg-[#1e3a8a] text-white shadow-sm';
        const activeBadge = activeBtn.querySelector('span:last-child');
        if (activeBadge) activeBadge.className = 'px-2 py-0.5 text-xs rounded-full bg-white/20 text-white';
      }
    };

    // ==================== MODULE 1: SUPPLIES INVENTORY ====================
    window.updateSupplyFormPreview = function() {
      const prev = parseInt(document.getElementById('sup_prev_total').value || '0', 10);
      const cred = parseInt(document.getElementById('sup_credit').value || '0', 10);
      const deb = parseInt(document.getElementById('sup_debit').value || '0', 10);
      const total = Math.max(0, prev + cred - deb);
      document.getElementById('supplyFormPreviewTotal').textContent = total;
    };

    window.handleAddSupply = async function(e) {
      e.preventDefault();
      const entry_date = document.getElementById('sup_entry_date').value;
      const descriptions = document.getElementById('sup_name').value
        .split(/[\n,]+/)
        .map(description => description.trim())
        .filter(Boolean);
      const date_received = document.getElementById('sup_date_received').value;
      const supply_type = document.getElementById('sup_type').value.trim();
      const previous_total = parseInt(document.getElementById('sup_prev_total').value || '0', 10);
      const credit = parseInt(document.getElementById('sup_credit').value || '0', 10);
      const debit = parseInt(document.getElementById('sup_debit').value || '0', 10);
      const total_amount = Math.max(0, previous_total + credit - debit);

      if (descriptions.length === 0) {
        showToast('Enter at least one supply description.', 'error');
        return;
      }

      try {
        for (const supply_name of descriptions) {
          await createRecord('office_supplies', 'supplies', {
            entry_date,
            supply_name,
            date_received,
            supply_type,
            previous_total,
            credit,
            debit,
            total_amount
          });
        }
        document.getElementById('sup_name').value = '';
        document.getElementById('sup_prev_total').value = '0';
        document.getElementById('sup_credit').value = '0';
        document.getElementById('sup_debit').value = '0';
        updateSupplyFormPreview();
        showToast(descriptions.length === 1
          ? `Added "${descriptions[0]}" (${total_amount} left)`
          : `Added ${descriptions.length} supply descriptions (${total_amount} left each)`);
      } catch (err) {
        console.error(err);
        showToast('Failed to add supply item.', 'error');
      }
    };

    window.openAdjustStockModal = function(id) {
      const item = state.supplies.find(s => s.id === id);
      if (!item) return;
      document.getElementById('adj_supply_id').value = item.id;
      document.getElementById('adj_supply_name').textContent = `${item.supply_name} (${item.supply_type})`;
      document.getElementById('adj_current_total').textContent = item.total_amount;
      document.getElementById('adj_credit').value = '0';
      document.getElementById('adj_debit').value = '0';
      document.getElementById('adj_date_received').value = '';
      document.getElementById('adj_new_total_preview').textContent = item.total_amount;
      document.getElementById('adjustStockModal').classList.remove('hidden');
    };

    window.updateAdjustPreview = function() {
      const id = document.getElementById('adj_supply_id').value;
      const item = state.supplies.find(s => s.id === id);
      if (!item) return;
      const addCred = parseInt(document.getElementById('adj_credit').value || '0', 10);
      const subDeb = parseInt(document.getElementById('adj_debit').value || '0', 10);
      const newTotal = Math.max(0, Number(item.total_amount || 0) + addCred - subDeb);
      document.getElementById('adj_new_total_preview').textContent = newTotal;
    };

    window.submitAdjustStock = async function(e) {
      e.preventDefault();
      const id = document.getElementById('adj_supply_id').value;
      const item = state.supplies.find(s => s.id === id);
      if (!item) return;

      const addCred = parseInt(document.getElementById('adj_credit').value || '0', 10);
      const subDeb = parseInt(document.getElementById('adj_debit').value || '0', 10);
      const newDateRec = document.getElementById('adj_date_received').value;
      const prevTotal = Number(item.total_amount || 0);
      const newTotal = Math.max(0, prevTotal + addCred - subDeb);

      try {
        await updateRecord('office_supplies', 'supplies', id, {
          entry_date: getTodayISO(),
          date_received: newDateRec || item.date_received || '',
          previous_total: prevTotal,
          credit: addCred,
          debit: subDeb,
          total_amount: newTotal
        });
        closeModal('adjustStockModal');
        showToast(`Updated "${item.supply_name}" → ${newTotal} left`);
      } catch (err) {
        console.error(err);
        showToast('Failed to update stock.', 'error');
      }
    };

    window.resetSupplyPageAndRender = function() {
      state.supplyPage = 1;
      renderSupplies();
    };

    window.changeSupplyPage = function(delta) {
      state.supplyPage += delta;
      renderSupplies();
    };

    function getFilteredSupplies() {
      const q = (document.getElementById('supplySearchInput').value || '').toLowerCase().trim();
      const cat = document.getElementById('supplyCategoryFilter').value;
      const lowOnly = document.getElementById('supplyLowStockOnly').checked;

      // Sort newest first in memory
      const sorted = [...state.supplies].sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));

      return sorted.filter(item => {
        const matchesQuery = !q ||
          (item.supply_name || '').toLowerCase().includes(q) ||
          (item.supply_type || '').toLowerCase().includes(q) ||
          (item.entry_date || '').includes(q);
        const matchesCat = cat === 'ALL' || item.supply_type === cat;
        const matchesLow = !lowOnly || Number(item.total_amount) <= 5;
        return matchesQuery && matchesCat && matchesLow;
      });
    }

    function renderSupplies() {
      // Update category dropdown dynamically with any custom categories
      const catSelect = document.getElementById('supplyCategoryFilter');
      const currentCat = catSelect.value;
      const defaultCats = ['Office Supplies', 'Equipment', 'Cleaning Supplies', 'Harmful Chemicals'];
      const dynamicCats = defaultCats;
      catSelect.innerHTML = `<option value="ALL">All Supply Types</option>` +
        dynamicCats.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
      if (dynamicCats.includes(currentCat) || currentCat === 'ALL') {
        catSelect.value = currentCat;
      }

      const filtered = getFilteredSupplies();
      const pageSize = parseInt(document.getElementById('supplyPageSize').value || '100', 10);
      const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
      if (state.supplyPage > totalPages) state.supplyPage = totalPages;
      if (state.supplyPage < 1) state.supplyPage = 1;

      const startIdx = (state.supplyPage - 1) * pageSize;
      const paged = filtered.slice(startIdx, startIdx + pageSize);

      document.getElementById('suppliesShowingInfo').textContent =
        `Showing ${paged.length} of ${filtered.length} matching items (${state.supplies.length} total in database)`;
      document.getElementById('supplyPageIndicator').textContent = `Page ${state.supplyPage} of ${totalPages}`;
      document.getElementById('supplyPrevPageBtn').disabled = state.supplyPage <= 1;
      document.getElementById('supplyNextPageBtn').disabled = state.supplyPage >= totalPages;

      const tbody = document.getElementById('suppliesTableBody');
      if (paged.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" class="px-4 py-8 text-center text-slate-400">No supply items match your filter. Add a supply above.</td></tr>`;
        return;
      }

      tbody.innerHTML = paged.map(item => {
        const total = Number(item.total_amount || 0);
        let badgeHtml = `<span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-bold bg-emerald-100 text-emerald-800">${total} left</span>`;
        if (total === 0) {
          badgeHtml = `<span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-bold bg-rose-100 text-rose-800">0 (Out)</span>`;
        } else if (total <= 5) {
          badgeHtml = `<span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-bold bg-amber-100 text-amber-800">${total} (Low)</span>`;
        }

        return `
          <tr class="hover:bg-slate-50/90 transition">
            <td class="px-3.5 py-3 whitespace-nowrap text-slate-600">${escapeHtml(item.entry_date)}</td>
            <td class="px-3.5 py-3 font-semibold text-slate-900">${escapeHtml(item.supply_name)}</td>
            <td class="px-3.5 py-3 whitespace-nowrap text-slate-600">${escapeHtml(item.date_received || '—')}</td>
            <td class="px-3.5 py-3 whitespace-nowrap">
              <span class="px-2 py-0.5 rounded-md text-xs font-medium bg-blue-50 text-blue-700 border border-blue-100">${escapeHtml(item.supply_type)}</span>
            </td>
            <td class="px-3.5 py-3 text-right font-mono text-slate-600">${Number(item.previous_total || 0)}</td>
            <td class="px-3.5 py-3 text-right font-mono font-semibold text-emerald-600">+${Number(item.credit || 0)}</td>
            <td class="px-3.5 py-3 text-right font-mono font-semibold text-rose-600">-${Number(item.debit || 0)}</td>
            <td class="px-3.5 py-3 text-center whitespace-nowrap">${badgeHtml}</td>
            <td class="px-3.5 py-3 text-right whitespace-nowrap space-x-1">
              <button onclick="openAdjustStockModal('${item.id}')" class="px-2.5 py-1 rounded-md text-xs font-semibold bg-blue-50 text-[#2563eb] hover:bg-blue-100 border border-blue-200 transition cursor-pointer">
                Adjust Stock
              </button>
              <button onclick="openEditModal('supplies', '${item.id}')" class="px-2 py-1 rounded-md text-xs font-semibold text-slate-600 hover:bg-slate-100 border border-slate-200 transition cursor-pointer">
                Edit
              </button>
              <button onclick="confirmDelete('office_supplies', 'supplies', '${item.id}', '${escapeHtml(item.supply_name)}')" class="px-2 py-1 rounded-md text-xs font-semibold text-rose-600 hover:bg-rose-50 border border-rose-200 transition cursor-pointer">
                Delete
              </button>
            </td>
          </tr>
        `;
      }).join('');
    }

    // ==================== MODULE 2: EVENT MONITORING ====================
    window.handleAddEvent = async function(e) {
      e.preventDefault();
      const event_date = document.getElementById('ev_date').value;
      const event_name = document.getElementById('ev_name').value.trim();
      const place = document.getElementById('ev_place').value.trim();
      const event_whole_day = document.getElementById('ev_event_whole_day').checked;
      const eventPeriods = readTimePeriods('ev', 'event');
      const event_day_from = document.getElementById('ev_event_day_from').value;
      const event_day_to = document.getElementById('ev_event_day_to').value;
      const include_setup = document.getElementById('ev_include_setup').checked;
      const setup_date = include_setup ? document.getElementById('ev_setup_date').value : '';
      const setupPeriods = include_setup
        ? readTimePeriods('ev', 'setup')
        : { setup_morning_from: '', setup_morning_to: '', setup_afternoon_from: '', setup_afternoon_to: '' };
      const arrangements = document.getElementById('ev_arrangements').value.trim();

      const eventHasTimes = Object.values(eventPeriods).some(Boolean);
      const eventTimeError = event_whole_day
        ? (eventHasTimes
          ? 'Whole Day cannot be combined with morning or afternoon time ranges.'
          : validateTimeRange(event_day_from, event_day_to, 'Whole-day event', true))
        : validateTimePeriods(eventPeriods, 'event', 'Event', true);
      const setupTimeError = include_setup ? validateTimePeriods(setupPeriods, 'setup', 'Set-up', true) : '';
      if (eventTimeError || setupTimeError) {
        showToast(eventTimeError || setupTimeError, 'error');
        return;
      }

      try {
        await createRecord('office_events', 'events', {
          event_date,
          event_name,
          place,
          event_time: event_whole_day ? '' : eventPeriods.event_morning_from || eventPeriods.event_afternoon_from,
          event_whole_day,
          event_day_from: event_whole_day ? event_day_from : '',
          event_day_to: event_whole_day ? event_day_to : '',
          ...eventPeriods,
          setup_date,
          setup_time: include_setup ? setupPeriods.setup_morning_from || setupPeriods.setup_afternoon_from : '',
          ...setupPeriods,
          arrangements,
          status: 'Upcoming'
        });
        document.getElementById('ev_name').value = '';
        document.getElementById('ev_place').value = '';
        document.getElementById('ev_arrangements').value = '';
        showToast(`Event "${event_name}" scheduled!`);
      } catch (err) {
        console.error(err);
        showToast('Failed to schedule event.', 'error');
      }
    };

    window.cycleEventStatus = async function(id) {
      const ev = state.events.find(item => item.id === id);
      if (!ev) return;
      const order = ['Upcoming', 'Set-up Ready', 'Completed'];
      const nextStatus = order[(order.indexOf(ev.status) + 1) % order.length];
      try {
        await updateRecord('office_events', 'events', id, { status: nextStatus });
        showToast(`Event marked as "${nextStatus}"`);
      } catch (err) {
        console.error(err);
        showToast('Failed to update status.', 'error');
      }
    };

    window.renderEvents = function() {
      const q = (document.getElementById('eventSearchInput').value || '').toLowerCase().trim();
      const statusFilter = document.getElementById('eventStatusFilter').value;

      const sorted = [...state.events].sort((a, b) => (a.event_date || '').localeCompare(b.event_date || ''));
      const filtered = sorted.filter(ev => {
        const matchesQ = !q ||
          (ev.event_name || '').toLowerCase().includes(q) ||
          (ev.place || '').toLowerCase().includes(q) ||
          (ev.arrangements || '').toLowerCase().includes(q);
        const matchesStatus = statusFilter === 'ALL' ||
          (statusFilter === 'ACTIVE' && ev.status !== 'Completed') ||
          (statusFilter === 'Completed' && ev.status === 'Completed');
        return matchesQ && matchesStatus;
      });

      const tbody = document.getElementById('eventsTableBody');
      if (filtered.length === 0) {
        tbody.innerHTML = `<tr><td colspan="8" class="px-4 py-8 text-center text-slate-400">No events found. Schedule an event above.</td></tr>`;
        return;
      }

      tbody.innerHTML = filtered.map(ev => {
        let statusBadge = `<span class="px-2.5 py-1 rounded-full text-xs font-bold bg-amber-100 text-amber-800">Upcoming</span>`;
        if (ev.status === 'Set-up Ready') {
          statusBadge = `<span class="px-2.5 py-1 rounded-full text-xs font-bold bg-blue-100 text-blue-800">Set-up Ready</span>`;
        } else if (ev.status === 'Completed') {
          statusBadge = `<span class="px-2.5 py-1 rounded-full text-xs font-bold bg-emerald-100 text-emerald-800">Completed</span>`;
        }
        const hasSetupDetails = Boolean(
          ev.setup_date || ev.setup_time ||
          ev.setup_morning_from || ev.setup_morning_to ||
          ev.setup_afternoon_from || ev.setup_afternoon_to
        );
        const setupDetails = hasSetupDetails
          ? `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-xs font-semibold bg-indigo-50 text-indigo-800 border border-indigo-200">${escapeHtml(ev.setup_date || '—')} • ${escapeHtml(formatTimePeriods(ev, 'setup'))}</span>`
          : '<span class="text-slate-400">No set-up</span>';

        return `
          <tr class="hover:bg-slate-50/90 transition">
            <td class="px-3.5 py-3 whitespace-nowrap font-semibold text-slate-800">${escapeHtml(ev.event_date)}</td>
            <td class="px-3.5 py-3 font-bold text-slate-900">${escapeHtml(ev.event_name)}</td>
            <td class="px-3.5 py-3 text-slate-700">${escapeHtml(ev.place)}</td>
            <td class="px-3.5 py-3 whitespace-nowrap font-medium text-slate-800">${escapeHtml(formatTimePeriods(ev, 'event'))}</td>
            <td class="px-3.5 py-3 whitespace-nowrap">${setupDetails}</td>
            <td class="px-3.5 py-3 text-xs text-slate-600 max-w-xs">${escapeHtml(ev.arrangements || '—')}</td>
            <td class="px-3.5 py-3 text-center whitespace-nowrap">
              <button onclick="cycleEventStatus('${ev.id}')" title="Click to toggle status" class="cursor-pointer hover:opacity-85 transition">
                ${statusBadge}
              </button>
            </td>
            <td class="px-3.5 py-3 text-right whitespace-nowrap space-x-1">
              <button onclick="openEditModal('events', '${ev.id}')" class="px-2 py-1 rounded-md text-xs font-semibold text-slate-600 hover:bg-slate-100 border border-slate-200 transition cursor-pointer">
                Edit
              </button>
              <button onclick="confirmDelete('office_events', 'events', '${ev.id}', '${escapeHtml(ev.event_name)}')" class="px-2 py-1 rounded-md text-xs font-semibold text-rose-600 hover:bg-rose-50 border border-rose-200 transition cursor-pointer">
                Delete
              </button>
            </td>
          </tr>
        `;
      }).join('');
    };

    // ==================== MODULE 3: CHECKLIST MONITORING ====================
    window.handleAddChecklist = async function(e) {
      e.preventDefault();
      const dept_name = document.getElementById('chk_dept_name').value.trim();
      const month_year = document.getElementById('chk_month').value;
      const restroom_checked = document.getElementById('chk_restroom').checked;
      const classroom_checked = document.getElementById('chk_classroom').checked;
      const restroom_monthly_inv_checked = document.getElementById('chk_restroom_monthly_inv').checked;
      const classroom_monthly_inv_checked = document.getElementById('chk_classroom_monthly_inv').checked;
      const documents_checked = document.getElementById('chk_documents').checked;
      const remarks = document.getElementById('chk_remarks').value.trim();

      try {
        await createRecord('office_checklists', 'checklists', {
          dept_name,
          month_year,
          restroom_checked,
          classroom_checked,
          restroom_monthly_inv_checked,
          classroom_monthly_inv_checked,
          documents_checked,
          remarks
        });
        document.getElementById('chk_dept_name').value = '';
        document.getElementById('chk_remarks').value = '';
        document.getElementById('chk_restroom').checked = false;
        document.getElementById('chk_classroom').checked = false;
        document.getElementById('chk_restroom_monthly_inv').checked = false;
        document.getElementById('chk_classroom_monthly_inv').checked = false;
        document.getElementById('chk_documents').checked = false;
        showToast(`Checklist saved for ${dept_name}`);
      } catch (err) {
        console.error(err);
        showToast('Failed to save checklist.', 'error');
      }
    };

    function isRestroomMonthlyInventoryChecked(checklist) {
      return checklist.restroom_monthly_inv_checked ?? checklist.monthly_inv_checked ?? false;
    }

    function isClassroomMonthlyInventoryChecked(checklist) {
      return checklist.classroom_monthly_inv_checked ?? checklist.monthly_inv_checked ?? false;
    }

    function renderChecklistStatus(checked, label, color = 'blue') {
      const checkedColor = color === 'green'
        ? 'bg-emerald-600 border-emerald-600'
        : 'bg-[#2563eb] border-[#2563eb]';
      const statusColor = checked ? checkedColor : 'bg-white border-slate-300';
      const checkmark = checked
        ? '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2.5" aria-hidden="true"><path stroke-linecap="round" stroke-linejoin="round" d="m3 8 3.2 3.2L13 4.5" /></svg>'
        : '';
      return `<span role="img" aria-label="${escapeHtml(label)}: ${checked ? 'Complete' : 'Incomplete'}" class="inline-flex h-4 w-4 items-center justify-center rounded border ${statusColor} ${checked ? 'text-white' : 'text-transparent'}">${checkmark}</span>`;
    }

    window.renderChecklists = function() {
      // Populate month filter
      const monthSelect = document.getElementById('checklistMonthFilter');
      const currentMonthVal = monthSelect.value;
      const months = Array.from(new Set(state.checklists.map(c => c.month_year).filter(Boolean))).sort().reverse();
      monthSelect.innerHTML = `<option value="ALL">All Months</option>` +
        months.map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join('');
      if (months.includes(currentMonthVal) || currentMonthVal === 'ALL') {
        monthSelect.value = currentMonthVal;
      }

      const q = (document.getElementById('checklistSearchInput').value || '').toLowerCase().trim();
      const monthFilter = monthSelect.value;

      const sorted = [...state.checklists].sort((a, b) => (b.month_year || '').localeCompare(a.month_year || ''));
      const filtered = sorted.filter(chk => {
        const matchesQ = !q ||
          (chk.dept_name || '').toLowerCase().includes(q) ||
          (chk.remarks || '').toLowerCase().includes(q);
        const matchesMonth = monthFilter === 'ALL' || chk.month_year === monthFilter;
        return matchesQ && matchesMonth;
      });

      const tbody = document.getElementById('checklistsTableBody');
      if (filtered.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" class="px-4 py-8 text-center text-slate-400">No checklist records found. Add a department checklist above.</td></tr>`;
        return;
      }

      tbody.innerHTML = filtered.map(chk => {
        const docBadge = chk.documents_checked
          ? `<span class="px-2 py-0.5 rounded-full text-xs font-bold bg-emerald-100 text-emerald-800">Documents Checked</span>`
          : `<span class="px-2 py-0.5 rounded-full text-xs font-bold bg-amber-100 text-amber-800">Pending Verification</span>`;

        return `
          <tr class="hover:bg-slate-50/90 transition">
            <td class="px-3.5 py-3 font-bold text-slate-900">${escapeHtml(chk.dept_name)}</td>
            <td class="px-3.5 py-3 whitespace-nowrap font-medium text-slate-700">${escapeHtml(chk.month_year)}</td>
            <td class="px-3.5 py-3 text-center">
              ${renderChecklistStatus(chk.restroom_checked, 'Restroom checklist')}
            </td>
            <td class="px-3.5 py-3 text-center">
              ${renderChecklistStatus(chk.classroom_checked, 'Classroom checklist')}
            </td>
            <td class="px-3.5 py-3 text-center">
              ${renderChecklistStatus(isRestroomMonthlyInventoryChecked(chk), 'Restroom monthly inventory')}
            </td>
            <td class="px-3.5 py-3 text-center">
              ${renderChecklistStatus(isClassroomMonthlyInventoryChecked(chk), 'Classroom monthly inventory')}
            </td>
            <td class="px-3.5 py-3 text-center whitespace-nowrap">
              <span class="inline-flex items-center gap-2">
                ${renderChecklistStatus(chk.documents_checked, 'Document check', 'green')}
                ${docBadge}
              </span>
            </td>
            <td class="px-3.5 py-3 text-xs text-slate-600">${escapeHtml(chk.remarks || '—')}</td>
            <td class="px-3.5 py-3 text-right whitespace-nowrap space-x-1">
              <button onclick="openEditModal('checklists', '${chk.id}')" class="px-2 py-1 rounded-md text-xs font-semibold text-slate-600 hover:bg-slate-100 border border-slate-200 transition cursor-pointer">
                Edit
              </button>
              <button onclick="confirmDelete('office_checklists', 'checklists', '${chk.id}', '${escapeHtml(chk.dept_name)} (${escapeHtml(chk.month_year)})')" class="px-2 py-1 rounded-md text-xs font-semibold text-rose-600 hover:bg-rose-50 border border-rose-200 transition cursor-pointer">
                Delete
              </button>
            </td>
          </tr>
        `;
      }).join('');
    };

    // ==================== MODULE 4: JSA DTR MONITORING ====================
    window.autoCalculateDtrHours = function() {
      const tIn = document.getElementById('dtr_time_in').value;
      const tOut = document.getElementById('dtr_time_out').value;
      const deductLunch = document.getElementById('dtr_deduct_lunch').checked;
      if (!tIn || !tOut) return;

      const [h1, m1] = tIn.split(':').map(Number);
      const [h2, m2] = tOut.split(':').map(Number);
      let diff = (h2 + m2 / 60) - (h1 + m1 / 60);
      if (diff < 0) diff += 24;
      if (deductLunch && diff > 5) diff -= 1;
      document.getElementById('dtr_hours').value = Math.max(0, diff).toFixed(2);
    };

    window.handleDtrStatusChange = function() {
      const status = document.getElementById('dtr_status').value;
      if (status === 'Absent' || status === 'On Leave') {
        document.getElementById('dtr_time_in').value = '';
        document.getElementById('dtr_time_out').value = '';
        document.getElementById('dtr_hours').value = '0.00';
      }
    };

    window.handleAddDtr = async function(e) {
      e.preventDefault();
      const log_date = document.getElementById('dtr_date').value;
      const employee_name = document.getElementById('dtr_employee_name').value.trim();
      const time_in = document.getElementById('dtr_time_in').value;
      const time_out = document.getElementById('dtr_time_out').value;
      const hours_worked = parseFloat(document.getElementById('dtr_hours').value || '0');
      const status = document.getElementById('dtr_status').value;
      const absent_form_passed = document.getElementById('dtr_absent_form').checked;
      const remarks = document.getElementById('dtr_remarks').value.trim();

      try {
        await createRecord('office_dtr', 'dtr', {
          log_date,
          employee_name,
          time_in,
          time_out,
          hours_worked,
          status,
          absent_form_passed,
          remarks
        });
        document.getElementById('dtr_remarks').value = '';
        document.getElementById('dtr_absent_form').checked = false;
        showToast(`DTR logged for ${employee_name} (${hours_worked.toFixed(2)} hrs)`);
      } catch (err) {
        console.error(err);
        showToast('Failed to log DTR.', 'error');
      }
    };

    window.renderDtr = function() {
      // Populate Datalist & Filters
      const names = Array.from(new Set(state.dtr.map(d => d.employee_name).filter(Boolean))).sort();
      document.getElementById('jsaNamesDatalist').innerHTML =
        names.map(n => `<option value="${escapeHtml(n)}"></option>`).join('');

      const empSelect = document.getElementById('dtrEmployeeFilter');
      const curEmp = empSelect.value;
      empSelect.innerHTML = `<option value="ALL">All JSA Co-Workers</option>` +
        names.map(n => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join('');
      if (names.includes(curEmp) || curEmp === 'ALL') empSelect.value = curEmp;

      const monthSelect = document.getElementById('dtrMonthFilter');
      const curMonth = monthSelect.value;
      const months = Array.from(new Set(state.dtr.map(d => (d.log_date || '').slice(0, 7)).filter(Boolean))).sort().reverse();
      monthSelect.innerHTML = `<option value="ALL">All Months</option>` +
        months.map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join('');
      if (months.includes(curMonth) || curMonth === 'ALL') monthSelect.value = curMonth;

      const q = (document.getElementById('dtrSearchInput').value || '').toLowerCase().trim();
      const empFilter = empSelect.value;
      const monthFilter = monthSelect.value;

      const sorted = [...state.dtr].sort((a, b) => (b.log_date || '').localeCompare(a.log_date || ''));
      const filtered = sorted.filter(row => {
        const matchesEmp = empFilter === 'ALL' || row.employee_name === empFilter;
        const matchesMonth = monthFilter === 'ALL' || (row.log_date || '').startsWith(monthFilter);
        const matchesQ = !q ||
          (row.employee_name || '').toLowerCase().includes(q) ||
          (row.log_date || '').includes(q) ||
          (row.remarks || '').toLowerCase().includes(q) ||
          (row.status || '').toLowerCase().includes(q);
        return matchesEmp && matchesMonth && matchesQ;
      });

      // Compute Summary Strip Metrics
      let daysPresent = 0;
      let totalHours = 0;
      let absences = 0;
      let absentForms = 0;

      filtered.forEach(r => {
        totalHours += Number(r.hours_worked || 0);
        if (r.status === 'Absent' || r.status === 'On Leave') {
          absences++;
        } else {
          daysPresent++;
        }
        if (r.absent_form_passed) absentForms++;
      });

      document.getElementById('sumDtrDaysPresent').textContent = `${daysPresent} day${daysPresent === 1 ? '' : 's'}`;
      document.getElementById('sumDtrTotalHours').textContent = `${totalHours.toFixed(2)} hrs`;
      document.getElementById('sumDtrAbsences').textContent = `${absences} day${absences === 1 ? '' : 's'}`;
      document.getElementById('sumDtrAbsentForms').textContent = `${absentForms} passed`;
      document.getElementById('dtrShowingInfo').textContent = `Showing ${filtered.length} of ${state.dtr.length} DTR records`;

      const tbody = document.getElementById('dtrTableBody');
      if (filtered.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" class="px-4 py-8 text-center text-slate-400">No JSA DTR entries match your filter.</td></tr>`;
        return;
      }

      tbody.innerHTML = filtered.map(row => {
        let statusBadge = `<span class="px-2.5 py-0.5 rounded-full text-xs font-bold bg-emerald-100 text-emerald-800">Present</span>`;
        if (row.status === 'Absent') {
          statusBadge = `<span class="px-2.5 py-0.5 rounded-full text-xs font-bold bg-rose-100 text-rose-800">Absent</span>`;
        } else if (row.status === 'Late' || row.status === 'Half-Day') {
          statusBadge = `<span class="px-2.5 py-0.5 rounded-full text-xs font-bold bg-amber-100 text-amber-800">${escapeHtml(row.status)}</span>`;
        } else if (row.status === 'On Leave') {
          statusBadge = `<span class="px-2.5 py-0.5 rounded-full text-xs font-bold bg-blue-100 text-blue-800">On Leave</span>`;
        }

        const absentFormBadge = row.absent_form_passed
          ? `<span class="px-2.5 py-0.5 rounded-full text-xs font-bold bg-blue-100 text-blue-800">Yes - Passed</span>`
          : `<span class="text-xs text-slate-400">No</span>`;

        return `
          <tr class="hover:bg-slate-50/90 transition">
            <td class="px-3.5 py-3 whitespace-nowrap font-medium text-slate-700">${escapeHtml(row.log_date)}</td>
            <td class="px-3.5 py-3 font-bold text-slate-900">${escapeHtml(row.employee_name)}</td>
            <td class="px-3.5 py-3 text-center whitespace-nowrap text-slate-700">${formatTime12(row.time_in)}</td>
            <td class="px-3.5 py-3 text-center whitespace-nowrap text-slate-700">${formatTime12(row.time_out)}</td>
            <td class="px-3.5 py-3 text-right whitespace-nowrap font-mono font-bold text-[#1e3a8a]">${Number(row.hours_worked || 0).toFixed(2)} hrs</td>
            <td class="px-3.5 py-3 text-center whitespace-nowrap">${statusBadge}</td>
            <td class="px-3.5 py-3 text-center whitespace-nowrap">${absentFormBadge}</td>
            <td class="px-3.5 py-3 text-xs text-slate-600">${escapeHtml(row.remarks || '—')}</td>
            <td class="px-3.5 py-3 text-right whitespace-nowrap space-x-1">
              <button onclick="openEditModal('dtr', '${row.id}')" class="px-2 py-1 rounded-md text-xs font-semibold text-slate-600 hover:bg-slate-100 border border-slate-200 transition cursor-pointer">
                Edit
              </button>
              <button onclick="confirmDelete('office_dtr', 'dtr', '${row.id}', '${escapeHtml(row.employee_name)} (${escapeHtml(row.log_date)})')" class="px-2 py-1 rounded-md text-xs font-semibold text-rose-600 hover:bg-rose-50 border border-rose-200 transition cursor-pointer">
                Delete
              </button>
            </td>
          </tr>
        `;
      }).join('');
    };

    // ==================== EDIT MODAL & CONFIRMATION MODAL ====================
    window.openEditModal = function(moduleKey, id) {
      const item = state[moduleKey].find(x => x.id === id);
      if (!item) return;
      state.editContext = { moduleKey, id };
      const fieldsDiv = document.getElementById('editModalFields');
      const titleEl = document.getElementById('editModalTitle');

      if (moduleKey === 'supplies') {
        titleEl.textContent = 'Edit Supply Item';
        fieldsDiv.innerHTML = `
          <div class="grid grid-cols-2 gap-3">
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">Date</label>
              <input type="date" id="ed_entry_date" value="${escapeHtml(item.entry_date)}" required class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">Date Received</label>
              <input type="date" id="ed_date_received" value="${escapeHtml(item.date_received || '')}" class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" />
            </div>
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">Name of Supply</label>
            <textarea id="ed_supply_name" required rows="3" class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm">${escapeHtml(item.supply_name)}</textarea>
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">Type of Supply</label>
            <select id="ed_supply_type" required class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm bg-white">
              <option value="Office Supplies" ${item.supply_type === 'Office Supplies' ? 'selected' : ''}>Office Supplies</option>
              <option value="Equipment" ${item.supply_type === 'Equipment' ? 'selected' : ''}>Equipment</option>
              <option value="Cleaning Supplies" ${item.supply_type === 'Cleaning Supplies' ? 'selected' : ''}>Cleaning Supplies</option>
              <option value="Harmful Chemicals" ${item.supply_type === 'Harmful Chemicals' ? 'selected' : ''}>Harmful Chemicals</option>
              ${item.supply_type && !['Office Supplies','Equipment','Cleaning Supplies','Harmful Chemicals'].includes(item.supply_type) ? `<option value="${escapeHtml(item.supply_type)}" selected>${escapeHtml(item.supply_type)} (Legacy)</option>` : ''}
            </select>
          </div>
          <div class="grid grid-cols-3 gap-3">
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">Prev. Total</label>
              <input type="number" id="ed_prev_total" value="${Number(item.previous_total || 0)}" min="0" required class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-emerald-700 mb-1">Credit (+Added)</label>
              <input type="number" id="ed_credit" value="${Number(item.credit || 0)}" min="0" required class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-rose-700 mb-1">Debit (-Used)</label>
              <input type="number" id="ed_debit" value="${Number(item.debit || 0)}" min="0" required class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" />
            </div>
          </div>
        `;
      } else if (moduleKey === 'events') {
        titleEl.textContent = 'Edit Event Schedule';
        const eventHasStructuredTime = item.event_whole_day || item.event_morning_from || item.event_morning_to || item.event_afternoon_from || item.event_afternoon_to;
        const setupHasStructuredTime = item.setup_morning_from || item.setup_morning_to || item.setup_afternoon_from || item.setup_afternoon_to;
        const hasSetupDetails = Boolean(item.setup_date || item.setup_time || setupHasStructuredTime);
        const legacyEventTime = item.event_whole_day && !item.event_day_from && !item.event_day_to
          ? '<p class="mt-1 text-[11px] text-slate-500">Add the whole-day start and end times.</p>'
          : !eventHasStructuredTime && item.event_time
            ? `<p class="mt-1 text-[11px] text-slate-500">Saved time: ${escapeHtml(formatTime12(item.event_time))}. Leave the new ranges blank to keep it.</p>`
            : '';
        const legacySetupTime = !setupHasStructuredTime && item.setup_time
          ? `<p class="mt-1 text-[11px] text-slate-500">Saved time: ${escapeHtml(formatTime12(item.setup_time))}. Leave the new ranges blank to keep it.</p>`
          : '';
        fieldsDiv.innerHTML = `
          <div class="grid grid-cols-1 gap-3">
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">Date of Event</label>
              <input type="date" id="ed_event_date" value="${escapeHtml(item.event_date)}" required class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" />
            </div>
            <div>
              <div class="flex flex-wrap items-center justify-between gap-2">
                <label class="block text-xs font-semibold text-slate-700">Time of Event</label>
                <label class="inline-flex items-center gap-2 text-xs font-semibold text-slate-700">
                  <input type="checkbox" id="ed_event_whole_day" ${item.event_whole_day || item.event_period === 'Whole Day' ? 'checked' : ''} onchange="toggleWholeDayEventTimes('ed')" class="rounded text-[#2563eb]" />
                  Whole Day
                </label>
              </div>
              <div id="ed_event_period_fields" class="grid grid-cols-1 sm:grid-cols-2 gap-2 mt-2 ${item.event_whole_day || item.event_period === 'Whole Day' ? 'hidden' : ''}">
                ${renderTimePeriodFields('ed', 'event', item)}
              </div>
              <div id="ed_event_whole_day_fields" class="${item.event_whole_day || item.event_period === 'Whole Day' ? '' : 'hidden'} mt-2 max-w-md">
                <fieldset class="rounded-lg border border-slate-200 p-3">
                  <legend class="px-1 text-xs font-semibold text-slate-700">Whole-day time range</legend>
                  <div class="grid grid-cols-2 gap-3">
                    <label class="text-xs text-slate-600">From<input type="time" id="ed_event_day_from" value="${escapeHtml(item.event_day_from || '')}" class="mt-1 w-full rounded-lg border border-slate-300 px-2 py-2 text-sm" /></label>
                    <label class="text-xs text-slate-600">To<input type="time" id="ed_event_day_to" value="${escapeHtml(item.event_day_to || '')}" class="mt-1 w-full rounded-lg border border-slate-300 px-2 py-2 text-sm" /></label>
                  </div>
                </fieldset>
              </div>
              ${legacyEventTime}
            </div>
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">Name of Event</label>
            <input type="text" id="ed_event_name" value="${escapeHtml(item.event_name)}" required class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" />
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">Place / Venue / Building</label>
            <input type="text" id="ed_place" list="buildingSuggestions" value="${escapeHtml(item.place)}" required class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" />
          </div>
          <div class="grid grid-cols-1 gap-3">
            <div>
              <label class="inline-flex items-center gap-2 text-xs font-semibold text-indigo-700 cursor-pointer">
                <input type="checkbox" id="ed_include_setup" ${hasSetupDetails ? 'checked' : ''} onchange="toggleEventSetupFields('ed')" class="rounded text-[#2563eb]" />
                Include set-up details
              </label>
              <div id="ed_setup_date_fields" class="${hasSetupDetails ? '' : 'hidden'} mt-2">
                <label class="block text-xs font-semibold text-indigo-700 mb-1">Set-up Date *</label>
                <input type="date" id="ed_setup_date" value="${escapeHtml(item.setup_date || '')}" class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" />
              </div>
            </div>
            <div id="ed_setup_time_fields" class="${hasSetupDetails ? '' : 'hidden'}">
              <label class="block text-xs font-semibold text-indigo-700 mb-1">Set-up Time *</label>
              <div class="grid grid-cols-1 sm:grid-cols-2 gap-2 mt-2">
                ${renderTimePeriodFields('ed', 'setup', item, 'indigo')}
              </div>
              ${legacySetupTime}
            </div>
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">Arrangements</label>
            <textarea id="ed_arrangements" rows="2" class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm">${escapeHtml(item.arrangements || '')}</textarea>
          </div>
        `;
        window.toggleEventSetupFields('ed');
      } else if (moduleKey === 'checklists') {
        titleEl.textContent = 'Edit Building Checklist';
        fieldsDiv.innerHTML = `
          <div class="grid grid-cols-2 gap-3">
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">Name of Building</label>
              <input type="text" id="ed_dept_name" list="buildingSuggestions" value="${escapeHtml(item.dept_name)}" required class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">Month</label>
              <input type="month" id="ed_month_year" value="${escapeHtml(item.month_year)}" required class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" />
            </div>
          </div>
          <div class="grid grid-cols-2 gap-2 pt-1">
            <label class="inline-flex items-center gap-2 text-sm">
              <input type="checkbox" id="ed_restroom" ${item.restroom_checked ? 'checked' : ''} class="rounded text-[#2563eb]" />
              <span>Restroom Checklist</span>
            </label>
            <label class="inline-flex items-center gap-2 text-sm">
              <input type="checkbox" id="ed_classroom" ${item.classroom_checked ? 'checked' : ''} class="rounded text-[#2563eb]" />
              <span>Classroom Checklist</span>
            </label>
            <label class="inline-flex items-center gap-2 text-sm">
              <input type="checkbox" id="ed_restroom_monthly_inv" ${isRestroomMonthlyInventoryChecked(item) ? 'checked' : ''} class="rounded text-[#2563eb]" />
              <span>Restroom Monthly Inventory</span>
            </label>
            <label class="inline-flex items-center gap-2 text-sm">
              <input type="checkbox" id="ed_classroom_monthly_inv" ${isClassroomMonthlyInventoryChecked(item) ? 'checked' : ''} class="rounded text-[#2563eb]" />
              <span>Classroom Monthly Inventory</span>
            </label>
            <label class="inline-flex items-center gap-2 text-sm font-semibold text-[#1e3a8a]">
              <input type="checkbox" id="ed_documents" ${item.documents_checked ? 'checked' : ''} class="rounded text-[#2563eb]" />
              <span>Checking of Documents</span>
            </label>
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">Remarks / Checked By</label>
            <input type="text" id="ed_remarks" value="${escapeHtml(item.remarks || '')}" class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" />
          </div>
        `;
      } else if (moduleKey === 'dtr') {
        titleEl.textContent = 'Edit JSA DTR Entry';
        fieldsDiv.innerHTML = `
          <div class="grid grid-cols-2 gap-3">
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">Date</label>
              <input type="date" id="ed_log_date" value="${escapeHtml(item.log_date)}" required class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">JSA Co-Worker Name</label>
              <input type="text" id="ed_employee_name" value="${escapeHtml(item.employee_name)}" required class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" />
            </div>
          </div>
          <div class="grid grid-cols-3 gap-3">
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">Time-In</label>
              <input type="time" id="ed_time_in" value="${escapeHtml(item.time_in || '')}" class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">Time-Out</label>
              <input type="time" id="ed_time_out" value="${escapeHtml(item.time_out || '')}" class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">Hours</label>
              <input type="number" step="0.01" id="ed_hours_worked" value="${Number(item.hours_worked || 0)}" required class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" />
            </div>
          </div>
          <div class="grid grid-cols-2 gap-3 items-center">
            <div>
              <label class="block text-xs font-semibold text-slate-700 mb-1">Status</label>
              <select id="ed_status" class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm bg-white">
                <option value="Present" ${item.status === 'Present' ? 'selected' : ''}>Present</option>
                <option value="Late" ${item.status === 'Late' ? 'selected' : ''}>Late</option>
                <option value="Half-Day" ${item.status === 'Half-Day' ? 'selected' : ''}>Half-Day</option>
                <option value="Absent" ${item.status === 'Absent' ? 'selected' : ''}>Absent</option>
                <option value="On Leave" ${item.status === 'On Leave' ? 'selected' : ''}>On Leave</option>
              </select>
            </div>
            <label class="inline-flex items-center gap-2 text-sm font-semibold text-slate-700 pt-4">
              <input type="checkbox" id="ed_absent_form" ${item.absent_form_passed ? 'checked' : ''} class="rounded text-[#2563eb]" />
              <span>Absent Form Passed</span>
            </label>
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">Remarks</label>
            <input type="text" id="ed_dtr_remarks" value="${escapeHtml(item.remarks || '')}" class="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" />
          </div>
        `;
      }

      document.getElementById('editRecordModal').classList.remove('hidden');
    };

    window.submitEditModal = async function(e) {
      e.preventDefault();
      if (!state.editContext) return;
      const { moduleKey, id } = state.editContext;

      try {
        if (moduleKey === 'supplies') {
          const prev = parseInt(document.getElementById('ed_prev_total').value || '0', 10);
          const cred = parseInt(document.getElementById('ed_credit').value || '0', 10);
          const deb = parseInt(document.getElementById('ed_debit').value || '0', 10);
          await updateRecord('office_supplies', 'supplies', id, {
            entry_date: document.getElementById('ed_entry_date').value,
            date_received: document.getElementById('ed_date_received').value,
            supply_name: document.getElementById('ed_supply_name').value.trim(),
            supply_type: document.getElementById('ed_supply_type').value.trim(),
            previous_total: prev,
            credit: cred,
            debit: deb,
            total_amount: Math.max(0, prev + cred - deb)
          });
        } else if (moduleKey === 'events') {
          const currentEvent = state.events.find(item => item.id === id);
          const event_whole_day = document.getElementById('ed_event_whole_day').checked;
          const include_setup = document.getElementById('ed_include_setup').checked;
          let eventPeriods = readTimePeriods('ed', 'event');
          const event_day_from = document.getElementById('ed_event_day_from').value;
          const event_day_to = document.getElementById('ed_event_day_to').value;
          let setupPeriods = include_setup
            ? readTimePeriods('ed', 'setup')
            : { setup_morning_from: '', setup_morning_to: '', setup_afternoon_from: '', setup_afternoon_to: '' };
          const eventHasTimes = Object.values(eventPeriods).some(Boolean);
          const setupHasTimes = Object.values(setupPeriods).some(Boolean);

          if (event_whole_day && eventHasTimes) {
            showToast('Whole Day cannot be combined with morning or afternoon time ranges.', 'error');
            return;
          }
          const eventTimeError = event_whole_day
            ? validateTimeRange(event_day_from, event_day_to, 'Whole-day event', true)
            : validateTimePeriods(eventPeriods, 'event', 'Event', false);
          const setupTimeError = include_setup ? validateTimePeriods(setupPeriods, 'setup', 'Set-up', false) : '';
          if (eventTimeError || setupTimeError) {
            showToast(eventTimeError || setupTimeError, 'error');
            return;
          }

          const currentEventHasTime = currentEvent && (
            currentEvent.event_time ||
            ['event_morning_from', 'event_morning_to', 'event_afternoon_from', 'event_afternoon_to'].some(key => currentEvent[key])
          );
          const currentSetupHasTime = currentEvent && (
            currentEvent.setup_time ||
            ['setup_morning_from', 'setup_morning_to', 'setup_afternoon_from', 'setup_afternoon_to'].some(key => currentEvent[key])
          );
          if (!event_whole_day && !eventHasTimes && currentEventHasTime) {
            eventPeriods = Object.fromEntries(
              ['event_morning_from', 'event_morning_to', 'event_afternoon_from', 'event_afternoon_to']
                .map(key => [key, currentEvent[key] || ''])
            );
          } else if (!event_whole_day && !eventHasTimes) {
            showToast('Enter an event time period or choose Whole Day.', 'error');
            return;
          }
          if (include_setup && !setupHasTimes && currentSetupHasTime) {
            setupPeriods = Object.fromEntries(
              ['setup_morning_from', 'setup_morning_to', 'setup_afternoon_from', 'setup_afternoon_to']
                .map(key => [key, currentEvent[key] || ''])
            );
          } else if (include_setup && !setupHasTimes) {
            showToast('Enter at least one set-up time period.', 'error');
            return;
          }

          await updateRecord('office_events', 'events', id, {
            event_date: document.getElementById('ed_event_date').value,
            event_time: event_whole_day ? '' : eventPeriods.event_morning_from || eventPeriods.event_afternoon_from || (currentEvent && currentEvent.event_time) || '',
            event_whole_day,
            event_day_from: event_whole_day ? event_day_from : '',
            event_day_to: event_whole_day ? event_day_to : '',
            ...eventPeriods,
            event_name: document.getElementById('ed_event_name').value.trim(),
            place: document.getElementById('ed_place').value.trim(),
            setup_date: include_setup ? document.getElementById('ed_setup_date').value : '',
            setup_time: include_setup ? setupPeriods.setup_morning_from || setupPeriods.setup_afternoon_from || (currentEvent && currentEvent.setup_time) || '' : '',
            ...setupPeriods,
            arrangements: document.getElementById('ed_arrangements').value.trim()
          });
        } else if (moduleKey === 'checklists') {
          await updateRecord('office_checklists', 'checklists', id, {
            dept_name: document.getElementById('ed_dept_name').value.trim(),
            month_year: document.getElementById('ed_month_year').value,
            restroom_checked: document.getElementById('ed_restroom').checked,
            classroom_checked: document.getElementById('ed_classroom').checked,
            restroom_monthly_inv_checked: document.getElementById('ed_restroom_monthly_inv').checked,
            classroom_monthly_inv_checked: document.getElementById('ed_classroom_monthly_inv').checked,
            documents_checked: document.getElementById('ed_documents').checked,
            remarks: document.getElementById('ed_remarks').value.trim()
          });
        } else if (moduleKey === 'dtr') {
          await updateRecord('office_dtr', 'dtr', id, {
            log_date: document.getElementById('ed_log_date').value,
            employee_name: document.getElementById('ed_employee_name').value.trim(),
            time_in: document.getElementById('ed_time_in').value,
            time_out: document.getElementById('ed_time_out').value,
            hours_worked: parseFloat(document.getElementById('ed_hours_worked').value || '0'),
            status: document.getElementById('ed_status').value,
            absent_form_passed: document.getElementById('ed_absent_form').checked,
            remarks: document.getElementById('ed_dtr_remarks').value.trim()
          });
        }
        closeModal('editRecordModal');
        showToast('Record updated successfully!');
      } catch (err) {
        console.error(err);
        showToast('Failed to update record.', 'error');
      }
    };

    window.confirmDelete = function(colName, stateKey, id, label) {
      document.getElementById('confirmModalTitle').textContent = 'Delete Record?';
      document.getElementById('confirmModalMessage').textContent = `Are you sure you want to permanently delete "${label}"? This will update all connected office computers.`;
      const actionBtn = document.getElementById('confirmModalActionBtn');
      actionBtn.onclick = async () => {
        try {
          await deleteRecord(colName, stateKey, id);
          closeModal('confirmModal');
          showToast(`Deleted "${label}"`);
        } catch (err) {
          console.error(err);
          showToast('Failed to delete record.', 'error');
        }
      };
      document.getElementById('confirmModal').classList.remove('hidden');
    };

    window.closeModal = function(modalId) {
      document.getElementById(modalId).classList.add('hidden');
    };

    // ==================== KPI & GLOBAL RENDER ====================
    function renderKPIs() {
      // Supplies KPIs
      const supCount = state.supplies.length;
      const lowStock = state.supplies.filter(s => Number(s.total_amount) <= 5).length;
      const totalUnits = state.supplies.reduce((acc, s) => acc + Number(s.total_amount || 0), 0);
      document.getElementById('kpiSupplyCount').textContent = supCount;
      document.getElementById('kpiTotalUnits').textContent = `${totalUnits} total units left`;
      const lowBadge = document.getElementById('kpiLowStockBadge');
      lowBadge.textContent = `${lowStock} Low Stock`;
      lowBadge.className = lowStock > 0
        ? 'px-2 py-0.5 rounded-full font-semibold bg-amber-100 text-amber-800'
        : 'px-2 py-0.5 rounded-full font-semibold bg-emerald-100 text-emerald-800';
      const tabBadgeSupplies = document.getElementById('tabBadgeSupplies');
      if (tabBadgeSupplies) tabBadgeSupplies.textContent = supCount;

      // Events KPIs
      const upcoming = state.events
        .filter(e => e.status !== 'Completed')
        .sort((a, b) => (a.event_date || '').localeCompare(b.event_date || ''));
      document.getElementById('kpiUpcomingEvents').textContent = upcoming.length;
      document.getElementById('kpiNextEventText').textContent = upcoming.length > 0
        ? `Next: ${upcoming[0].event_name} (${upcoming[0].event_date})`
        : 'No upcoming events scheduled';
      const tabBadgeEvents = document.getElementById('tabBadgeEvents');
      if (tabBadgeEvents) tabBadgeEvents.textContent = state.events.length;

      // Checklists KPIs
      const pendingChk = state.checklists.filter(c => !c.documents_checked).length;
      const verifiedChk = state.checklists.filter(c => c.documents_checked).length;
      document.getElementById('kpiPendingChecklists').textContent = pendingChk;
      document.getElementById('kpiVerifiedChecklists').textContent = `${verifiedChk} Verified`;
      const tabBadgeChecklists = document.getElementById('tabBadgeChecklists');
      if (tabBadgeChecklists) tabBadgeChecklists.textContent = state.checklists.length;

      // DTR KPIs
      const totalHours = state.dtr.reduce((acc, r) => acc + Number(r.hours_worked || 0), 0);
      const uniqueStaff = new Set(state.dtr.map(r => r.employee_name).filter(Boolean)).size;
      const absentForms = state.dtr.filter(r => r.absent_form_passed).length;
      document.getElementById('kpiDtrHours').textContent = totalHours.toFixed(2);
      document.getElementById('kpiJsaStaffCount').textContent = `${uniqueStaff} Co-Worker${uniqueStaff === 1 ? '' : 's'}`;
      document.getElementById('kpiAbsentFormsCount').textContent = `${absentForms} absent form${absentForms === 1 ? '' : 's'}`;
      const tabBadgeDtr = document.getElementById('tabBadgeDtr');
      if (tabBadgeDtr) tabBadgeDtr.textContent = state.dtr.length;
    }

    function renderAll() {
      renderKPIs();
      renderSupplies();
      renderEvents();
      renderChecklists();
      renderDtr();
    }

    // ==================== CSV & JSON EXPORT HELPERS ====================
    function downloadFile(filename, content, mimeType) {
      const blob = new Blob([content], { type: mimeType });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }

    function toCSVRow(arr) {
      return arr.map(val => `"${String(val ?? '').replace(/"/g, '""')}"`).join(',');
    }

    window.exportSuppliesCSV = function() {
      const headers = ['Date', 'Name of Supply', 'Date Received', 'Type of Supply', 'Previous Total', 'Credit (+Added)', 'Debit (-Used)', 'Total Left'];
      const rows = state.supplies.map(s => [
        s.entry_date, s.supply_name, s.date_received, s.supply_type,
        s.previous_total, s.credit, s.debit, s.total_amount
      ]);
      const csv = [toCSVRow(headers), ...rows.map(toCSVRow)].join('\n');
      downloadFile(`Supplies_Inventory_${getTodayISO()}.csv`, csv, 'text/csv;charset=utf-8;');
      showToast('Exported Supplies Inventory CSV!');
    };

    window.exportEventsCSV = function() {
      const headers = ['Date of Event', 'Name of Event', 'Place', 'Time of Event', 'Set-up Date', 'Set-up Time', 'Arrangements', 'Status'];
      const rows = state.events.map(e => [
        e.event_date, e.event_name, e.place, formatTimePeriods(e, 'event'),
        e.setup_date, formatTimePeriods(e, 'setup'), e.arrangements, e.status
      ]);
      const csv = [toCSVRow(headers), ...rows.map(toCSVRow)].join('\n');
      downloadFile(`Event_Monitoring_${getTodayISO()}.csv`, csv, 'text/csv;charset=utf-8;');
      showToast('Exported Event Monitoring CSV!');
    };

    window.exportChecklistsCSV = function() {
      const headers = ['Building', 'Month', 'Restroom Checklist', 'Classroom Checklist', 'Restroom Monthly Inventory', 'Classroom Monthly Inventory', 'Checking of Documents', 'Remarks / Checked By'];
      const rows = state.checklists.map(c => [
        c.dept_name, c.month_year,
        c.restroom_checked ? 'Checked' : 'Pending',
        c.classroom_checked ? 'Checked' : 'Pending',
        isRestroomMonthlyInventoryChecked(c) ? 'Checked' : 'Pending',
        isClassroomMonthlyInventoryChecked(c) ? 'Checked' : 'Pending',
        c.documents_checked ? 'Checked' : 'Pending',
        c.remarks
      ]);
      const csv = [toCSVRow(headers), ...rows.map(toCSVRow)].join('\n');
      downloadFile(`Checklist_Monitoring_${getTodayISO()}.csv`, csv, 'text/csv;charset=utf-8;');
      showToast('Exported Checklist Monitoring CSV!');
    };

    window.exportDtrCSV = function() {
      const headers = ['Date', 'JSA Co-Worker Name', 'Time-In', 'Time-Out', 'Number of Hours', 'Status', 'Absent Form Passed', 'Remarks'];
      const rows = state.dtr.map(d => [
        d.log_date, d.employee_name, d.time_in, d.time_out,
        Number(d.hours_worked || 0).toFixed(2), d.status,
        d.absent_form_passed ? 'Yes' : 'No', d.remarks
      ]);
      const csv = [toCSVRow(headers), ...rows.map(toCSVRow)].join('\n');
      downloadFile(`JSA_DTR_Monitoring_${getTodayISO()}.csv`, csv, 'text/csv;charset=utf-8;');
      showToast('Exported JSA DTR CSV!');
    };

    window.exportAllCSV = function() {
      const sections = [
        '=== 1. SUPPLIES INVENTORY ===',
        toCSVRow(['Date', 'Name of Supply', 'Date Received', 'Type of Supply', 'Prev Total', 'Credit', 'Debit', 'Total Left']),
        ...state.supplies.map(s => toCSVRow([s.entry_date, s.supply_name, s.date_received, s.supply_type, s.previous_total, s.credit, s.debit, s.total_amount])),
        '',
        '=== 2. EVENT MONITORING ===',
        toCSVRow(['Date of Event', 'Name of Event', 'Place', 'Time of Event', 'Set-up Date', 'Set-up Time', 'Arrangements', 'Status']),
        ...state.events.map(e => toCSVRow([e.event_date, e.event_name, e.place, formatTimePeriods(e, 'event'), e.setup_date, formatTimePeriods(e, 'setup'), e.arrangements, e.status])),
        '',
        '=== 3. CHECKLIST MONITORING ===',
        toCSVRow(['Building', 'Month', 'Restroom Checklist', 'Classroom Checklist', 'Restroom Monthly Inventory', 'Classroom Monthly Inventory', 'Checking of Documents', 'Remarks / Checked By']),
        ...state.checklists.map(c => toCSVRow([c.dept_name, c.month_year, c.restroom_checked ? 'Yes' : 'No', c.classroom_checked ? 'Yes' : 'No', isRestroomMonthlyInventoryChecked(c) ? 'Yes' : 'No', isClassroomMonthlyInventoryChecked(c) ? 'Yes' : 'No', c.documents_checked ? 'Yes' : 'No', c.remarks])),
        '',
        '=== 4. JSA DTR MONITORING ===',
        toCSVRow(['Date', 'JSA Co-Worker Name', 'Time-In', 'Time-Out', 'Hours', 'Status', 'Absent Form Passed', 'Remarks']),
        ...state.dtr.map(d => toCSVRow([d.log_date, d.employee_name, d.time_in, d.time_out, Number(d.hours_worked || 0).toFixed(2), d.status, d.absent_form_passed ? 'Yes' : 'No', d.remarks]))
      ];
      downloadFile(`Office_Database_Complete_${getTodayISO()}.csv`, sections.join('\n'), 'text/csv;charset=utf-8;');
      showToast('Exported Complete Office Database CSV!');
    };

    window.backupJSON = function() {
      const backup = {
        exportedAt: new Date().toISOString(),
        supplies: state.supplies,
        events: state.events,
        checklists: state.checklists,
        dtr: state.dtr
      };
      downloadFile(`Office_Database_Backup_${getTodayISO()}.json`, JSON.stringify(backup, null, 2), 'application/json');
      showToast('Downloaded full JSON backup!');
    };

    // ==================== SAMPLE OFFICE DATA LOADER ====================
    window.loadSampleData = async function() {
      const today = getTodayISO();
      const month = getCurrentMonthISO();

      const sampleSupplies = [
        { entry_date: today, supply_name: 'Bond Paper A4 (80gsm Ream)', date_received: today, supply_type: 'Office Supplies', previous_total: 40, credit: 20, debit: 12, total_amount: 48 },
        { entry_date: today, supply_name: 'Black Ballpen (Box of 12)', date_received: today, supply_type: 'Office Supplies', previous_total: 10, credit: 5, debit: 11, total_amount: 4 },
        { entry_date: today, supply_name: 'Liquid Hand Soap (Gallon)', date_received: today, supply_type: 'Cleaning Supplies', previous_total: 8, credit: 4, debit: 3, total_amount: 9 },
        { entry_date: today, supply_name: 'Printer (Office Unit)', date_received: today, supply_type: 'Equipment', previous_total: 6, credit: 0, debit: 4, total_amount: 2 },
        { entry_date: today, supply_name: 'Whiteboard Marker (Assorted)', date_received: today, supply_type: 'Office Supplies', previous_total: 15, credit: 10, debit: 7, total_amount: 18 }
      ];

      const sampleEvents = [
        { event_date: today, event_name: 'Quarterly Department Coordination Meeting', place: 'Main Conference Room', event_time: '09:00', setup_date: today, setup_time: '07:30', arrangements: '25 chairs, conference mic setup, projector & HDMI, coffee & water station', status: 'Set-up Ready' },
        { event_date: today, event_name: 'JSA Orientation & Safety Briefing', place: 'Audio-Visual Room (AVR)', event_time: '14:00', setup_date: today, setup_time: '12:30', arrangements: '50 monoblock chairs, 2 wireless mics, attendance registration table', status: 'Upcoming' }
      ];

      const sampleChecklists = [
        { dept_name: 'Senior High School Building', month_year: month, restroom_checked: true, classroom_checked: true, restroom_monthly_inv_checked: true, classroom_monthly_inv_checked: true, documents_checked: true, remarks: 'All checklists verified and filed' },
        { dept_name: 'Junior High School Building', month_year: month, restroom_checked: true, classroom_checked: true, restroom_monthly_inv_checked: false, classroom_monthly_inv_checked: false, documents_checked: false, remarks: 'Awaiting monthly inventory sheet' },
        { dept_name: 'Elementary Building', month_year: month, restroom_checked: true, classroom_checked: true, restroom_monthly_inv_checked: true, classroom_monthly_inv_checked: false, documents_checked: false, remarks: 'Pending supervisor signature on documents' }
      ];

      const sampleDtr = [
        { log_date: today, employee_name: 'Juan Dela Cruz', time_in: '07:55', time_out: '17:00', hours_worked: 8.08, status: 'Present', absent_form_passed: false, remarks: 'Regular shift (1hr lunch deducted)' },
        { log_date: today, employee_name: 'Maria Clara Santos', time_in: '08:15', time_out: '17:00', hours_worked: 7.75, status: 'Late', absent_form_passed: false, remarks: '15 mins late due to traffic' },
        { log_date: today, employee_name: 'Jose Rizalino Reyes', time_in: '', time_out: '', hours_worked: 0.00, status: 'Absent', absent_form_passed: true, remarks: 'Medical check-up; Absent Form submitted' }
      ];

      try {
        showToast('Loading sample office records...', 'info');
        for (const s of sampleSupplies) await createRecord('office_supplies', 'supplies', s);
        for (const ev of sampleEvents) await createRecord('office_events', 'events', ev);
        for (const chk of sampleChecklists) await createRecord('office_checklists', 'checklists', chk);
        for (const d of sampleDtr) await createRecord('office_dtr', 'dtr', d);
        showToast('Sample office data loaded across all 4 modules!');
      } catch (err) {
        console.error(err);
        showToast('Error loading sample data.', 'error');
      }
    };

    // Boot Application
    window.addEventListener('DOMContentLoaded', () => {
      initDefaultDates();
      initDatabase();
    });