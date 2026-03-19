(() => {
	const APPS_SCRIPT_WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbyEeXu5602tSW2uth2qrhf9d60N4utS06MZNrQ0HHs6JWtjqMHGrASHPInBAHs8IIpA8g/exec';

	const state = {
		citas: [],
		filtered: [],
		selectedId: null,
		activeTab: 'resumen',
		quickFilter: 'pendientes',
		visibleCount: 30,
		isAuthenticated: false,
	};

	const LIST_BATCH_SIZE = 30;
	const AUTH_SESSION_KEY = 'gestion_psicologia_auth';

	const ui = {
		app: document.getElementById('app'),
		loginScreen: document.getElementById('loginScreen'),
		loginForm: document.getElementById('loginForm'),
		loginError: document.getElementById('loginError'),
		listContainer: document.getElementById('listContainer'),
		searchInput: document.getElementById('searchInput'),
		quickFilters: document.getElementById('quickFilters'),
		emptyState: document.getElementById('emptyState'),
		detailPanel: document.getElementById('detailPanel'),
		detailNombre: document.getElementById('detailNombre'),
		detailMeta: document.getElementById('detailMeta'),
		detailFechaHighlight: document.getElementById('detailFechaHighlight'),
		detailHoraHighlight: document.getElementById('detailHoraHighlight'),
		detailInstitutoMeta: document.getElementById('detailInstitutoMeta'),
		detailEstadoBadge: document.getElementById('detailEstadoBadge'),
		tabsNav: document.getElementById('tabsNav'),
		globalLoader: document.getElementById('globalLoader'),
		toastContainer: document.getElementById('toastContainer'),
		refreshBtn: document.getElementById('refreshBtn'),

		kpiTotal: document.getElementById('kpiTotal'),
		kpiRealizadas: document.getElementById('kpiRealizadas'),
		kpiNoAsistio: document.getElementById('kpiNoAsistio'),
		kpiCanceladas: document.getElementById('kpiCanceladas'),

		resumenCorreo: document.getElementById('resumenCorreo'),
		resumenCreado: document.getElementById('resumenCreado'),
		resumenTelefono: document.getElementById('resumenTelefono'),
		resumenEdadCiclo: document.getElementById('resumenEdadCiclo'),
		resumenCarrera: document.getElementById('resumenCarrera'),
		resumenConvive: document.getElementById('resumenConvive'),
		resumenMotivo: document.getElementById('resumenMotivo'),

		previewNumeroSesion: document.getElementById('previewNumeroSesion'),
		previewProfundidad: document.getElementById('previewProfundidad'),
		previewImpacto: document.getElementById('previewImpacto'),
		previewTipoDificultad: document.getElementById('previewTipoDificultad'),
		previewCompromiso: document.getElementById('previewCompromiso'),
		tipoDificultadSelect: document.getElementById('tipoDificultadSelect'),
		tipoDificultadOtro: document.getElementById('tipoDificultadOtro'),

		notasForm: document.getElementById('notasForm'),
		reminderForm: document.getElementById('reminderForm'),

		personalModal: document.getElementById('personalModal'),
		personalForm: document.getElementById('personalForm'),
		openEditPersonalBtn: document.getElementById('openEditPersonalBtn'),
		closePersonalModalBtn: document.getElementById('closePersonalModalBtn'),
		cancelPersonalBtn: document.getElementById('cancelPersonalBtn'),
	};

	document.addEventListener('DOMContentLoaded', init);

	function init() {
		bindEvents();
		renderQuickFilters();

		if (hasActiveAuthSession_()) {
			completeLoginUI_();
			loadCitas();
			return;
		}

		showLoginUI_();
	}

	function bindEvents() {
		ui.loginForm.addEventListener('submit', handleLogin);
		ui.searchInput.addEventListener('input', handleSearch);
		ui.quickFilters.addEventListener('click', handleQuickFilterClick);
		ui.tabsNav.addEventListener('click', handleTabClick);
		ui.notasForm.addEventListener('submit', handleSaveNotas);
		ui.reminderForm.addEventListener('submit', handleSendReminder);
		ui.refreshBtn.addEventListener('click', loadCitas);
		ui.tipoDificultadSelect.addEventListener('change', handleTipoDificultadChange);

		ui.openEditPersonalBtn.addEventListener('click', openPersonalModal);
		ui.closePersonalModalBtn.addEventListener('click', closePersonalModal);
		ui.cancelPersonalBtn.addEventListener('click', closePersonalModal);
		ui.personalForm.addEventListener('submit', handleSavePersonalInfo);
	}

	function handleLogin(e) {
		e.preventDefault();
		const formData = getFormData(ui.loginForm);
		const submitBtn = document.getElementById('loginBtn');

		ui.loginError.classList.add('hidden');
		ui.loginError.textContent = '';
		setButtonLoading(submitBtn, true);

		runServer('validateLogin', formData.usuario, formData.contrasena)
			.then(() => {
				persistAuthSession_();
				completeLoginUI_();
				loadCitas();
			})
			.catch((err) => {
				ui.loginError.textContent = err.message || 'No se pudo iniciar sesión.';
				ui.loginError.classList.remove('hidden');
			})
			.finally(() => setButtonLoading(submitBtn, false));
	}

	function loadCitas() {
		if (!state.isAuthenticated) return;

		setLoading(true);
		runServer('getCitas')
			.then((citas) => {
				state.citas = Array.isArray(citas) ? sortCitas(citas) : [];
				state.visibleCount = LIST_BATCH_SIZE;
				applyFilters();

				if (state.selectedId) {
					const stillExists = state.citas.some((c) => c.id === state.selectedId);
					if (!stillExists) state.selectedId = null;
				}

				renderKpis(state.citas);
				renderList();
				renderDetail();
			})
			.catch((err) => showToast(err.message || 'Error al cargar citas', 'error'))
			.finally(() => setLoading(false));
	}

	function handleSearch(e) {
		void e;
		state.visibleCount = LIST_BATCH_SIZE;
		applyFilters();
	}

	function handleQuickFilterClick(e) {
		const btn = e.target.closest('[data-filter]');
		if (!btn) return;

		state.quickFilter = btn.dataset.filter || 'pendientes';
		state.visibleCount = LIST_BATCH_SIZE;
		renderQuickFilters();
		applyFilters();
	}

	function renderQuickFilters() {
		ui.quickFilters.querySelectorAll('.pill-btn').forEach((btn) => {
			btn.classList.toggle('active-pill', btn.dataset.filter === state.quickFilter);
		});
	}

	function applyFilters() {
		const query = normalize(ui.searchInput.value);
		const filteredByQuick = state.citas.filter((cita) => matchesQuickFilter(cita, state.quickFilter));

		if (!query) {
			state.filtered = filteredByQuick;
			renderList();
			return;
		}

		state.filtered = filteredByQuick.filter((cita) => {
			const stack = [
				cita.reservadoPor,
				cita.correo,
				cita.telefono,
				cita.instituto,
				cita.estado,
				cita.fecha,
				cita.hora,
			]
				.join(' ')
				.toLowerCase();
			return stack.includes(query);
		});
		renderList();
	}

	function renderKpis(citas) {
		const data = {
			total: citas.length,
			realizadas: 0,
			noAsistio: 0,
			canceladas: 0,
		};

		citas.forEach((cita) => {
			const estado = normalize(cita.estado);
			if (estado === 'realizada') data.realizadas += 1;
			else if (estado === 'no asistió' || estado === 'no asistio') data.noAsistio += 1;
			else if (estado === 'cancelada') data.canceladas += 1;
		});

		ui.kpiTotal.textContent = data.total;
		ui.kpiRealizadas.textContent = data.realizadas;
		ui.kpiNoAsistio.textContent = data.noAsistio;
		ui.kpiCanceladas.textContent = data.canceladas;
	}

	function renderList() {
		if (!state.filtered.length) {
			ui.listContainer.innerHTML = '<div class="p-4 text-sm text-gray-500">No se encontraron citas.</div>';
			return;
		}

		const visibleItems = state.filtered.slice(0, state.visibleCount);
		let lastGroupLabel = '';

		const html = visibleItems
			.map((cita) => {
				const active = state.selectedId === cita.id;
				const instituteTag = getInstituteBadge(cita.instituto);
				const groupLabel = getDateGroupLabel(cita.fecha);
				const fechaHora = `${formatDateDisplay(cita.fecha)} - ${formatTimeDisplay(cita.hora)}`;
				const showGroup = groupLabel !== lastGroupLabel;
				lastGroupLabel = groupLabel;

				return `
					${showGroup ? `<div class="px-2 pt-3 pb-1 text-[11px] uppercase tracking-wide text-gray-400 font-semibold">${escapeHtml(groupLabel)}</div>` : ''}
					<button data-id="${escapeAttr(cita.id)}" class="w-full text-left p-3 mb-2 rounded-xl border transition ${
						active
							? 'border-brand-300 bg-brand-50 shadow-sm'
							: 'border-gray-100 bg-white hover:bg-gray-50'
					}">
						<div class="flex justify-between items-start gap-2">
							<h3 class="font-medium text-gray-900">${escapeHtml(cita.reservadoPor || 'Sin nombre')}</h3>
							<span class="${getStatusClass(cita.estado)} text-xs px-2 py-1 rounded-full font-medium">${escapeHtml(cita.estado || 'Programada')}</span>
						</div>
						<p class="text-xs text-gray-500 mt-1">${escapeHtml(fechaHora)}</p>
						<div class="mt-2 flex items-center gap-2">
							${instituteTag}
						</div>
					</button>
				`;
			})
			.join('');

		const hasMore = state.filtered.length > state.visibleCount;
		const loadMoreButton = hasMore
			? `
				<button id="loadMoreBtn" class="w-full mt-1 mb-2 px-4 py-2.5 rounded-xl border border-gray-200 text-sm text-gray-600 hover:text-gray-800 hover:bg-gray-50 transition">
					Cargar más citas...
				</button>
			`
			: '';

		ui.listContainer.innerHTML = html + loadMoreButton;

		ui.listContainer.querySelectorAll('button[data-id]').forEach((btn) => {
			btn.addEventListener('click', () => {
				state.selectedId = btn.dataset.id;
				renderList();
				renderDetail();
			});
		});

		const loadMoreBtn = document.getElementById('loadMoreBtn');
		if (loadMoreBtn) {
			loadMoreBtn.addEventListener('click', () => {
				state.visibleCount += LIST_BATCH_SIZE;
				renderList();
			});
		}
	}

	function renderDetail() {
		const cita = getSelectedCita();
		if (!cita) {
			ui.emptyState.classList.remove('hidden');
			ui.detailPanel.classList.add('hidden');
			return;
		}

		ui.emptyState.classList.add('hidden');
		ui.detailPanel.classList.remove('hidden');

		const fechaVista = formatDateDisplay(cita.fecha);
		const horaVista = formatTimeDisplay(cita.hora);
		const creadoVista = formatCreatedDisplay(cita.creado);
		ui.detailNombre.textContent = cita.reservadoPor || 'Sin nombre';
		ui.detailFechaHighlight.textContent = fechaVista;
		ui.detailHoraHighlight.textContent = horaVista;
		ui.detailInstitutoMeta.textContent = cita.instituto || 'Sin instituto';
		ui.detailEstadoBadge.className = `${getStatusClass(cita.estado)} text-xs px-2.5 py-1 rounded-full font-medium`;
		ui.detailEstadoBadge.textContent = cita.estado || 'Programada';

		ui.resumenCreado.textContent = `Cita registrada en el sistema ${creadoVista}`;
		ui.resumenCorreo.textContent = cita.correo || '—';
		ui.resumenTelefono.textContent = cita.telefono || '—';
		ui.resumenEdadCiclo.textContent = `${cita.edad || '—'} años / Ciclo ${cita.ciclo || '—'}`;
		ui.resumenCarrera.textContent = cita.carrera || '—';
		ui.resumenConvive.textContent = cita.conviveCon || '—';
		ui.resumenMotivo.textContent = cita.motivoConsulta || '—';

		const sessionNumber = extractSessionNumber(cita.numeroSesion);
		const normalizedProfundidad = normalizeScaleValue(cita.profundidadProblema);
		const normalizedImpacto = normalizeScaleValue(cita.impactoExitoProfesional);
		const normalizedCompromiso = normalizeScaleValue(cita.nivelCompromiso);
		const tipoDificultadConfig = mapTipoDificultadValue(cita.tipoDificultad);

		hydrateForm(ui.notasForm, {
			numeroSesionNumero: sessionNumber,
			estado: cita.estado || 'Programada',
			problemaHallado: cita.problemaHallado,
			profundidadProblemaSelect: normalizedProfundidad,
			impactoExitoProfesionalSelect: normalizedImpacto,
			tipoDificultadSelect: tipoDificultadConfig.selectValue,
			tipoDificultadOtro: tipoDificultadConfig.otherValue,
			nivelCompromisoSelect: normalizedCompromiso,
			observaciones: cita.observaciones,
		});

		handleTipoDificultadChange();
		renderNotasCurrentPreview(cita);

		hydrateForm(ui.reminderForm, {
			mensaje: '',
			meetLink: '',
		});

		hydrateForm(ui.personalForm, {
			reservadoPor: cita.reservadoPor,
			correo: cita.correo,
			telefono: cita.telefono,
		});

		switchTab(state.activeTab);
	}

	function handleTabClick(e) {
		const btn = e.target.closest('[data-tab]');
		if (!btn) return;
		switchTab(btn.dataset.tab);
	}

	function switchTab(tab) {
		state.activeTab = tab;
		document.querySelectorAll('.tab-btn').forEach((btn) => {
			btn.classList.toggle('active-tab', btn.dataset.tab === tab);
		});
		document.querySelectorAll('.tab-content').forEach((panel) => {
			panel.classList.toggle('hidden', panel.id !== `tab-${tab}`);
		});
	}

	function handleSaveNotas(e) {
		e.preventDefault();
		const cita = getSelectedCita();
		if (!cita) return;

		const formData = getFormData(ui.notasForm);
		const numeroSesionValue = formData.numeroSesionNumero
			? `Completada sesión ${formData.numeroSesionNumero}`
			: '';
		const tipoDificultadValue = formData.tipoDificultadSelect === 'Otros'
			? (formData.tipoDificultadOtro || 'Otros')
			: (formData.tipoDificultadSelect || '');

		const payload = {
			numeroSesion: numeroSesionValue,
			estado: formData.estado,
			problemaHallado: formData.problemaHallado,
			profundidadProblema: formData.profundidadProblemaSelect || 'No aplica',
			impactoExitoProfesional: formData.impactoExitoProfesionalSelect || 'No aplica',
			tipoDificultad: tipoDificultadValue,
			nivelCompromiso: formData.nivelCompromisoSelect || 'No aplica',
			observaciones: formData.observaciones,
		};
		setButtonLoading(e.submitter, true);

		runServer('updateCita', cita.id, payload)
			.then((res) => {
				updateLocalCita(res.cita);
				showToast('Cita actualizada correctamente', 'success');
			})
			.catch((err) => showToast(err.message || 'No se pudo actualizar', 'error'))
			.finally(() => setButtonLoading(e.submitter, false));
	}

	function handleSavePersonalInfo(e) {
		e.preventDefault();
		const cita = getSelectedCita();
		if (!cita) return;

		const payload = getFormData(ui.personalForm);
		const submitBtn = document.getElementById('savePersonalBtn');
		setButtonLoading(submitBtn, true);

		runServer('updateCita', cita.id, payload)
			.then((res) => {
				updateLocalCita(res.cita);
				closePersonalModal();
				showToast('Información personal actualizada', 'success');
			})
			.catch((err) => showToast(err.message || 'No se pudo guardar', 'error'))
			.finally(() => setButtonLoading(submitBtn, false));
	}

	function handleSendReminder(e) {
		e.preventDefault();
		const cita = getSelectedCita();
		if (!cita) return;

		const { mensaje, meetLink } = getFormData(ui.reminderForm);
		if (!cita.correo) {
			showToast('La cita no tiene correo registrado', 'error');
			return;
		}

		const payload = {
			estudiante: cita.reservadoPor || '',
			correo: cita.correo || '',
			mensaje: mensaje || '',
			meet: meetLink || '',
			fechaCita: formatDateDisplay(cita.fecha),
			horaCita: formatTimeDisplay(cita.hora),
			instituto: cita.instituto || '',
			fechaEnvio: getNowPayloadTimestamp(),
		};

		setButtonLoading(e.submitter, true);
		runServer('sendReminder', payload)
			.then((res) => {
				if (res?.channel === 'email-fallback') {
					showToast('n8n no disponible. Recordatorio enviado por correo', 'info');
					return;
				}
				showToast('Recordatorio enviado a n8n correctamente', 'success');
			})
			.catch((err) => showToast(err.message || 'No se pudo enviar el correo', 'error'))
			.finally(() => setButtonLoading(e.submitter, false));
	}

	function openPersonalModal() {
		ui.personalModal.classList.remove('hidden');
	}

	function closePersonalModal() {
		ui.personalModal.classList.add('hidden');
	}

	function getSelectedCita() {
		return state.citas.find((c) => c.id === state.selectedId) || null;
	}

	function updateLocalCita(updated) {
		if (!updated || !updated.id) return;
		state.citas = sortCitas(state.citas.map((c) => (c.id === updated.id ? updated : c)));
		state.visibleCount = Math.max(state.visibleCount, LIST_BATCH_SIZE);
		applyFilters();

		renderKpis(state.citas);
		renderDetail();
	}

	function getStatusClass(status) {
		const st = normalize(status);
		if (st === 'realizada') return 'bg-emerald-100 text-emerald-700';
		if (st === 'no asistió' || st === 'no asistio') return 'bg-rose-100 text-rose-700';
		if (st === 'cancelada') return 'bg-amber-100 text-amber-700';
		if (st === 'reprogramada') return 'bg-blue-100 text-blue-700';
		return 'bg-yellow-100 text-yellow-700';
	}

	function getInstituteBadge(instituto) {
		const text = escapeHtml(instituto || 'Sin instituto');
		const isPurple = normalize(instituto).includes('adni');
		const cls = isPurple
			? 'bg-purple-100 text-purple-700'
			: 'bg-orange-100 text-orange-700';
		return `<span class="${cls} text-[11px] font-medium px-2 py-1 rounded-full">${text}</span>`;
	}

	function hydrateForm(form, data) {
		if (!form || !data) return;
		Object.entries(data).forEach(([key, value]) => {
			const input = form.elements.namedItem(key);
			if (input) input.value = value ?? '';
		});
	}

	function handleTipoDificultadChange() {
		const isOtros = (ui.tipoDificultadSelect.value || '') === 'Otros';
		ui.tipoDificultadOtro.classList.toggle('hidden', !isOtros);
		if (!isOtros) ui.tipoDificultadOtro.value = '';
	}

	function renderNotasCurrentPreview(cita) {
		ui.previewNumeroSesion.textContent = cita.numeroSesion || '—';
		ui.previewProfundidad.textContent = cita.profundidadProblema || '—';
		ui.previewImpacto.textContent = cita.impactoExitoProfesional || '—';
		ui.previewTipoDificultad.textContent = cita.tipoDificultad || '—';
		ui.previewCompromiso.textContent = cita.nivelCompromiso || '—';
	}

	function getFormData(form) {
		const formData = new FormData(form);
		const out = {};
		formData.forEach((value, key) => {
			out[key] = typeof value === 'string' ? value.trim() : value;
		});
		return out;
	}

	function setLoading(show) {
		ui.globalLoader.classList.toggle('hidden', !show);
	}

	function setButtonLoading(btn, loading) {
		if (!btn) return;
		btn.disabled = loading;
		if (loading) {
			btn.dataset.originalText = btn.textContent;
			btn.textContent = 'Procesando...';
		} else if (btn.dataset.originalText) {
			btn.textContent = btn.dataset.originalText;
		}
	}

	function runServer(functionName, ...args) {
		return new Promise((resolve, reject) => {
			const hasGoogleRun =
				typeof google !== 'undefined' &&
				google &&
				google.script &&
				google.script.run;

			// Modo 1: frontend embebido en Apps Script
			if (hasGoogleRun && !APPS_SCRIPT_WEBAPP_URL) {
				google.script.run
					.withSuccessHandler(resolve)
					.withFailureHandler((error) => {
						reject(error instanceof Error ? error : new Error(error?.message || 'Error desconocido.'));
					})[functionName](...args);
				return;
			}

			// Modo 2: frontend externo (repo separado) consumiendo URL Web App
			if (!APPS_SCRIPT_WEBAPP_URL) {
				reject(
					new Error(
						'Configura APPS_SCRIPT_WEBAPP_URL en main.js con tu URL /exec de Apps Script.'
					)
				);
				return;
			}

			const body = new URLSearchParams({
				action: functionName,
				payload: JSON.stringify(args),
			});

			fetch(APPS_SCRIPT_WEBAPP_URL, {
				method: 'POST',
				body,
			})
				.then(async (res) => {
					const raw = await res.text();
					let json;
					try {
						json = JSON.parse(raw);
					} catch (e) {
						throw new Error('Respuesta inválida del backend de Apps Script.');
					}

					if (!json.ok) {
						throw new Error(json.error || 'Error en backend');
					}

					return json.result;
				})
				.then(resolve)
				.catch((err) => reject(err instanceof Error ? err : new Error('Error de red.')));
		});
	}

	function showToast(message, type = 'info') {
		const toast = document.createElement('div');
		const classes = {
			success: 'bg-emerald-600',
			error: 'bg-rose-600',
			info: 'bg-gray-900',
		};

		toast.className = `text-white text-sm px-4 py-3 rounded-lg shadow-lg ${classes[type] || classes.info}`;
		toast.textContent = message;
		ui.toastContainer.appendChild(toast);

		setTimeout(() => {
			toast.classList.add('opacity-0', 'translate-y-2', 'transition');
			setTimeout(() => toast.remove(), 250);
		}, 2600);
	}

	function matchesQuickFilter(cita, quickFilter) {
		if (quickFilter === 'todas') return true;

		const citaDate = parseDateOnly(cita.fecha);
		const today = getTodayAtMidnight();

		if (quickFilter === 'hoy') {
			return citaDate && citaDate.getTime() === today.getTime();
		}

		// pendientes (por defecto)
		const estado = normalize(cita.estado);
		return estado === '' || estado === 'programada';
	}

	function sortCitas(citas) {
		return [...citas].sort((a, b) => {
			const da = getDateTimeValue(a.fecha, a.hora);
			const db = getDateTimeValue(b.fecha, b.hora);
			return db - da;
		});
	}

	function getDateTimeValue(fechaStr, horaStr) {
		const date = parseDateOnly(fechaStr);
		if (!date) return Number.MAX_SAFE_INTEGER;

		const { hours24, minutes } = parseTimeParts(horaStr);
		const hh = Number.isFinite(hours24) ? hours24 : 0;
		const mm = Number.isFinite(minutes) ? minutes : 0;

		date.setHours(hh, mm, 0, 0);
		return date.getTime();
	}

	function formatDateDisplay(fechaStr) {
		const date = parseDateOnly(fechaStr);
		if (!date) return '—';

		const day = String(date.getDate()).padStart(2, '0');
		const month = getMonthAbbr(date.getMonth());
		const year = date.getFullYear();
		return `${day} ${month} ${year}`;
	}

	function formatCreatedDisplay(createdStr) {
		if (!createdStr) return '—';
		const raw = String(createdStr).trim();
		if (!raw) return '—';

		const parsed = new Date(raw.replace(' ', 'T'));
		if (Number.isNaN(parsed.getTime())) {
			return `el ${raw}`;
		}

		return `el ${formatDateDisplay(formatDateToIso_(parsed))} a las ${formatTimeDisplay(formatTimeToHHMM_(parsed))}`;
	}

	function formatDateToIso_(date) {
		const y = date.getFullYear();
		const m = String(date.getMonth() + 1).padStart(2, '0');
		const d = String(date.getDate()).padStart(2, '0');
		return `${y}-${m}-${d}`;
	}

	function formatTimeToHHMM_(date) {
		const hh = String(date.getHours()).padStart(2, '0');
		const mm = String(date.getMinutes()).padStart(2, '0');
		return `${hh}:${mm}`;
	}

	function extractSessionNumber(numeroSesionRaw) {
		const raw = String(numeroSesionRaw || '').trim();
		if (!raw) return '';
		const match = raw.match(/(\d{1,2})/);
		if (!match) return '';
		const n = Number(match[1]);
		if (!Number.isFinite(n) || n < 1 || n > 10) return '';
		return String(n);
	}

	function normalizeScaleValue(value) {
		const v = normalize(value);
		if (v === 'bajo') return 'Bajo';
		if (v === 'medio') return 'Medio';
		if (v === 'alto') return 'Alto';
		return 'No aplica';
	}

	function mapTipoDificultadValue(value) {
		const options = [
			'No aplica',
			'Vocacional',
			'Social comunicativa',
			'Afectiva emocional',
			'Gerencia del tiempo',
			'Ansiedad',
			'Estres',
			'Familiar',
		];

		const raw = String(value || '').trim();
		if (!raw) {
			return { selectValue: 'No aplica', otherValue: '' };
		}

		const exact = options.find((opt) => normalize(opt) === normalize(raw));
		if (exact) {
			return { selectValue: exact, otherValue: '' };
		}

		return { selectValue: 'Otros', otherValue: raw };
	}

	function formatTimeDisplay(horaStr) {
		const { hours24, minutes } = parseTimeParts(horaStr);
		if (!Number.isFinite(hours24) || !Number.isFinite(minutes)) {
			return String(horaStr || '—').trim() || '—';
		}

		const period = hours24 >= 12 ? 'PM' : 'AM';
		const hour12 = hours24 % 12 === 0 ? 12 : hours24 % 12;
		const min2 = String(minutes).padStart(2, '0');
		return `${hour12}:${min2} ${period}`;
	}

	function parseTimeParts(horaStr) {
		const raw = String(horaStr || '').trim();
		if (!raw) return { hours24: NaN, minutes: NaN };

		const match = raw.match(/^(\d{1,2})\s*:\s*(\d{2})\s*([AaPp][Mm])?$/);
		if (!match) return { hours24: NaN, minutes: NaN };

		let hours = Number(match[1]);
		const minutes = Number(match[2]);
		const suffix = (match[3] || '').toUpperCase();

		if (!Number.isFinite(hours) || !Number.isFinite(minutes)) {
			return { hours24: NaN, minutes: NaN };
		}

		if (suffix === 'AM') {
			hours = hours === 12 ? 0 : hours;
		} else if (suffix === 'PM') {
			hours = hours === 12 ? 12 : hours + 12;
		}

		if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) {
			return { hours24: NaN, minutes: NaN };
		}

		return { hours24: hours, minutes };
	}

	function getMonthAbbr(monthIndex) {
		const months = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
		return months[monthIndex] || '—';
	}

	function getNowPayloadTimestamp() {
		const now = new Date();
		const y = now.getFullYear();
		const m = String(now.getMonth() + 1).padStart(2, '0');
		const d = String(now.getDate()).padStart(2, '0');
		const hh = String(now.getHours()).padStart(2, '0');
		const mm = String(now.getMinutes()).padStart(2, '0');
		const ss = String(now.getSeconds()).padStart(2, '0');
		return `${y}-${m}-${d} ${hh}:${mm}:${ss}`;
	}

	function parseDateOnly(value) {
		if (!value) return null;
		const raw = String(value).trim();

		if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
			const d = new Date(`${raw}T00:00:00`);
			return Number.isNaN(d.getTime()) ? null : d;
		}

		const fallback = new Date(raw);
		if (Number.isNaN(fallback.getTime())) return null;
		fallback.setHours(0, 0, 0, 0);
		return fallback;
	}

	function getTodayAtMidnight() {
		const today = new Date();
		today.setHours(0, 0, 0, 0);
		return today;
	}

	function getDateGroupLabel(fechaStr) {
		const d = parseDateOnly(fechaStr);
		if (!d) return 'Sin fecha';

		const today = getTodayAtMidnight();
		const yesterday = new Date(today);
		yesterday.setDate(yesterday.getDate() - 1);

		if (d.getTime() === today.getTime()) return 'Hoy';
		if (d.getTime() === yesterday.getTime()) return 'Ayer';

		return d.toLocaleDateString('es-PE', {
			month: 'long',
			year: 'numeric',
		});
	}

	function normalize(v) {
		return String(v || '').toLowerCase().trim();
	}

	function escapeHtml(str) {
		return String(str || '')
			.replace(/&/g, '&amp;')
			.replace(/</g, '&lt;')
			.replace(/>/g, '&gt;')
			.replace(/"/g, '&quot;')
			.replace(/'/g, '&#39;');
	}

	function escapeAttr(str) {
		return escapeHtml(str).replace(/`/g, '&#96;');
	}

	function showLoginUI_() {
		state.isAuthenticated = false;
		ui.app.classList.add('hidden');
		ui.loginScreen.classList.remove('hidden');
	}

	function completeLoginUI_() {
		state.isAuthenticated = true;
		ui.loginScreen.classList.add('hidden');
		ui.app.classList.remove('hidden');
	}

	function persistAuthSession_() {
		sessionStorage.setItem(AUTH_SESSION_KEY, '1');
	}

	function hasActiveAuthSession_() {
		return sessionStorage.getItem(AUTH_SESSION_KEY) === '1';
	}
})();
