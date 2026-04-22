(() => {
	const APPS_SCRIPT_WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbzX6M5CfWHB6Ez27hFzTHC-xzsTbU-g7DPcsZjhiBkVS_tXpFACxn-XfGkUneIFx8R5lA/exec';
	const DELETE_CITA_WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbx0Yxqxhk-_15pSZAFTR_omHFzGWX4kjzikyi4QgOJNJnL8J2Dh8NFhZWlbM3uDwqAL/exec';

	const state = {
		citas: [],
		filtered: [],
		selectedId: null,
		activeTab: 'resumen',
		quickFilter: 'pendientes',
		listInstituteFilters: {
			jvn: false,
			iempresa: false,
		},
		calendarInstituteFilters: {
			jvn: false,
			iempresa: false,
		},
		leftView: 'list',
		calendarCursor: new Date(),
		calendarSelectedDate: '',
		visibleCount: 30,
		isAuthenticated: false,
	};

	const LIST_BATCH_SIZE = 30;
	const AUTH_SESSION_KEY = 'gestion_psicologia_auth';
	const MEET_LINKS_BY_INSTITUTE = {
		jvn: 'https://meet.google.com/uft-mcjq-pdv',
		ie: 'https://meet.google.com/ouj-kben-cjo',
	};

	const INSTITUTE_FILTER_CONFIG = {
		jvn: { label: 'Neumann', fullLabel: 'Neumann' },
		iempresa: { label: 'IEmpresa', fullLabel: 'IEmpresa' },
	};

	const ui = {
		app: document.getElementById('app'),
		appHeader: document.getElementById('appHeader'),
		mainContent: document.getElementById('mainContent'),
		workspacePanels: document.getElementById('workspacePanels'),
		loginScreen: document.getElementById('loginScreen'),
		loginForm: document.getElementById('loginForm'),
		loginError: document.getElementById('loginError'),
		listContainer: document.getElementById('listContainer'),
		calendarContainer: document.getElementById('calendarContainer'),
		searchBlock: document.getElementById('searchBlock'),
		searchInput: document.getElementById('searchInput'),
		quickFiltersMain: document.getElementById('quickFiltersMain'),
		quickFiltersInstitutes: document.getElementById('quickFiltersInstitutes'),
		viewToggle: document.getElementById('viewToggle'),
		viewListBtn: document.getElementById('viewListBtn'),
		viewCalendarBtn: document.getElementById('viewCalendarBtn'),
		calendarMonthLabel: document.getElementById('calendarMonthLabel'),
		calendarWeekdays: document.getElementById('calendarWeekdays'),
		calendarGrid: document.getElementById('calendarGrid'),
		calendarDayList: document.getElementById('calendarDayList'),
		calendarPrevBtn: document.getElementById('calendarPrevBtn'),
		calendarNextBtn: document.getElementById('calendarNextBtn'),
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
		logoutBtn: document.getElementById('logoutBtn'),

		kpiTotal: document.getElementById('kpiTotal'),
		kpiRealizadas: document.getElementById('kpiRealizadas'),
		kpiNoAsistio: document.getElementById('kpiNoAsistio'),
		kpiCanceladas: document.getElementById('kpiCanceladas'),
		kpiReprogramadas: document.getElementById('kpiReprogramadas'),

		kpiTotalJVN: document.getElementById('kpiTotalJVN'),
		kpiTotalIE: document.getElementById('kpiTotalIE'),
		kpiRealizadasJVN: document.getElementById('kpiRealizadasJVN'),
		kpiRealizadasIE: document.getElementById('kpiRealizadasIE'),
		kpiNoAsistioJVN: document.getElementById('kpiNoAsistioJVN'),
		kpiNoAsistioIE: document.getElementById('kpiNoAsistioIE'),
		kpiCanceladasJVN: document.getElementById('kpiCanceladasJVN'),
		kpiCanceladasIE: document.getElementById('kpiCanceladasIE'),
		kpiReprogramadasJVN: document.getElementById('kpiReprogramadasJVN'),
		kpiReprogramadasIE: document.getElementById('kpiReprogramadasIE'),

		resumenCorreo: document.getElementById('resumenCorreo'),
		resumenCreado: document.getElementById('resumenCreado'),
		resumenTelefono: document.getElementById('resumenTelefono'),
		resumenDni: document.getElementById('resumenDni'),
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
		anularCitaBtn: document.getElementById('anularCitaBtn'),
		reminderForm: document.getElementById('reminderForm'),
		sendReminderBtn: document.getElementById('sendReminderBtn'),
		reminderModalidadBadge: document.getElementById('reminderModalidadBadge'),
		reminderMeetLabel: document.getElementById('reminderMeetLabel'),
		reminderMeetLinkPreview: document.getElementById('reminderMeetLinkPreview'),
		reminderStateHelp: document.getElementById('reminderStateHelp'),
		reminderMeetQuickAccess: document.getElementById('reminderMeetQuickAccess'),

		personalModal: document.getElementById('personalModal'),
		personalForm: document.getElementById('personalForm'),
		openEditPersonalBtn: document.getElementById('openEditPersonalBtn'),
		closePersonalModalBtn: document.getElementById('closePersonalModalBtn'),
		cancelPersonalBtn: document.getElementById('cancelPersonalBtn'),
	};

	document.addEventListener('DOMContentLoaded', init);

	function init() {
		state.calendarCursor = getStartOfMonth(new Date());
		state.calendarSelectedDate = getDefaultBusinessDayIso(new Date());
		bindEvents();
		renderQuickFilters();
		renderLeftViewToggle();
		renderLeftControlsVisibility();
		adjustWorkspacePanelsHeight();

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
		ui.quickFiltersMain.addEventListener('click', handleQuickFilterClick);
		ui.quickFiltersInstitutes.addEventListener('click', handleQuickFilterClick);
		ui.viewToggle.addEventListener('click', handleViewToggleClick);
		ui.calendarPrevBtn.addEventListener('click', handleCalendarPrevMonth);
		ui.calendarNextBtn.addEventListener('click', handleCalendarNextMonth);
		ui.calendarGrid.addEventListener('click', handleCalendarDayClick);
		ui.calendarDayList.addEventListener('click', handleCalendarAgendaClick);
		ui.tabsNav.addEventListener('click', handleTabClick);
		ui.notasForm.addEventListener('submit', handleSaveNotas);
		ui.anularCitaBtn.addEventListener('click', handleDeleteCita);
		ui.reminderForm.addEventListener('submit', handleSendReminder);
		ui.refreshBtn.addEventListener('click', loadCitas);
		ui.logoutBtn.addEventListener('click', handleLogout);
		ui.tipoDificultadSelect.addEventListener('change', handleTipoDificultadChange);

		ui.openEditPersonalBtn.addEventListener('click', openPersonalModal);
		ui.closePersonalModalBtn.addEventListener('click', closePersonalModal);
		ui.cancelPersonalBtn.addEventListener('click', closePersonalModal);
		ui.personalForm.addEventListener('submit', handleSavePersonalInfo);
		window.addEventListener('resize', adjustWorkspacePanelsHeight);
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

	function handleLogout() {
		clearAuthSession_();
		state.citas = [];
		state.filtered = [];
		state.selectedId = null;
		showLoginUI_();
		showToast('Sesión cerrada correctamente', 'info');
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
				renderLeftPane();
				renderDetail();
			})
			.catch((err) => showToast(err.message || 'Error al cargar citas', 'error'))
			.finally(() => setLoading(false));

		requestAnimationFrame(adjustWorkspacePanelsHeight);
	}

	function handleSearch(e) {
		void e;
		state.visibleCount = LIST_BATCH_SIZE;
		applyFilters();
	}

	function handleQuickFilterClick(e) {
		const btn = e.target.closest('[data-filter]');
		if (!btn) return;
		const filter = btn.dataset.filter || 'pendientes';

		if (filter === 'jvn' || filter === 'iempresa') {
			const targetFilters = state.leftView === 'calendar'
				? state.calendarInstituteFilters
				: state.listInstituteFilters;
			const wasActive = Boolean(targetFilters[filter]);

			if (wasActive) {
				targetFilters.jvn = false;
				targetFilters.iempresa = false;
			} else if (filter === 'jvn') {
				targetFilters.jvn = true;
				targetFilters.iempresa = false;
			} else {
				targetFilters.iempresa = true;
				targetFilters.jvn = false;
			}
			state.visibleCount = LIST_BATCH_SIZE;
			renderQuickFilters();
			applyFilters();
			return;
		}

		state.quickFilter = filter;
		state.visibleCount = LIST_BATCH_SIZE;
		renderQuickFilters();
		applyFilters();
	}

	function handleViewToggleClick(e) {
		const btn = e.target.closest('[data-view]');
		if (!btn) return;

		switchLeftView(btn.dataset.view || 'list');
	}

	function switchLeftView(view) {
		if (view !== 'list' && view !== 'calendar') return;
		state.leftView = view;
		renderLeftViewToggle();
		renderLeftControlsVisibility();
		renderQuickFilters();
		applyFilters();
	}

	function renderLeftControlsVisibility() {
		const isCalendar = state.leftView === 'calendar';
		ui.searchBlock.classList.toggle('hidden', isCalendar);
		ui.quickFiltersMain.classList.toggle('hidden', isCalendar);
	}

	function renderQuickFilters() {
		renderMainQuickFilters();
		renderInstituteQuickFilters();
	}

	function renderMainQuickFilters() {
		if (!ui.quickFiltersMain) return;
		ui.quickFiltersMain.querySelectorAll('.pill-btn').forEach((btn) => {
			btn.classList.toggle('active-pill', btn.dataset.filter === state.quickFilter);
		});
	}

	function renderInstituteQuickFilters() {
		if (!ui.quickFiltersInstitutes) return;
		const filtersByView = state.leftView === 'calendar'
			? state.calendarInstituteFilters
			: state.listInstituteFilters;

		const chips = ['jvn', 'iempresa']
			.map((filter) => {
				const active = Boolean(filtersByView[filter]);
				const config = INSTITUTE_FILTER_CONFIG[filter];
				const label = config?.label || filter.toUpperCase();
				const colorClass = filter === 'jvn'
					? (active
						? 'bg-purple-600 text-white border-purple-600 hover:bg-purple-700'
						: 'bg-gray-100 text-gray-600 border-gray-200 hover:bg-gray-200')
					: (active
						? 'bg-orange-500 text-white border-orange-500 hover:bg-orange-600'
						: 'bg-gray-100 text-gray-600 border-gray-200 hover:bg-gray-200');
				return `
					<button data-filter="${filter}" class="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium border cursor-pointer transition ${colorClass}">
						<span>${escapeHtml(label)}</span>
						${active ? '<span class="inline-flex items-center justify-center w-4 h-4 rounded-full bg-white/20 text-[10px] font-bold leading-none">×</span>' : ''}
					</button>
				`;
			})
			.join('');

		ui.quickFiltersInstitutes.innerHTML = `
			<span class="text-[10px] text-gray-400 font-medium mr-1">Filtros:</span>
			${chips}
		`;
	}

	function renderLeftViewToggle() {
		if (!ui.viewToggle) return;
		ui.viewListBtn.classList.toggle('active-pill', state.leftView === 'list');
		ui.viewCalendarBtn.classList.toggle('active-pill', state.leftView === 'calendar');
	}

	function renderLeftPane() {
		if (state.leftView === 'calendar') {
			ui.listContainer.classList.add('hidden');
			ui.calendarContainer.classList.remove('hidden');
			renderCalendar();
			return;
		}

		ui.calendarContainer.classList.add('hidden');
		ui.listContainer.classList.remove('hidden');
		renderList();
	}

	function applyFilters() {
		const isCalendar = state.leftView === 'calendar';
		const query = isCalendar ? '' : normalize(ui.searchInput.value);
		const effectiveQuickFilter = isCalendar ? 'todas' : state.quickFilter;
		const filteredByQuick = state.citas.filter((cita) => matchesQuickFilter(cita, effectiveQuickFilter));
		const instituteFiltersByView = isCalendar ? state.calendarInstituteFilters : state.listInstituteFilters;
		const activeInstitutes = Object.entries(instituteFiltersByView)
			.filter(([, enabled]) => enabled)
			.map(([key]) => key);

		const filteredByInstitute = !activeInstitutes.length
			? filteredByQuick
			: filteredByQuick.filter((cita) => {
				const group = getInstituteGroup(cita.instituto);
				if (group === 'ie') return activeInstitutes.includes('iempresa');
				return activeInstitutes.includes(group);
			});

		if (!query) {
			state.filtered = filteredByInstitute;
			renderLeftPane();
			return;
		}

		state.filtered = filteredByInstitute.filter((cita) => {
			const stack = normalize([
				cita.reservadoPor,
				cita.correo,
				cita.telefono,
				cita.instituto,
				cita.estado,
				cita.fecha,
				cita.hora,
			]
				.join(' '));
			return stack.includes(query);
		});
		renderLeftPane();
	}

	function renderKpis(citas) {
		const data = {
			total: citas.length,
			totalJVN: 0,
			totalIE: 0,
			realizadas: 0,
			realizadasJVN: 0,
			realizadasIE: 0,
			noAsistio: 0,
			noAsistioJVN: 0,
			noAsistioIE: 0,
			canceladas: 0,
			canceladasJVN: 0,
			canceladasIE: 0,
			reprogramadas: 0,
			reprogramadasJVN: 0,
			reprogramadasIE: 0,
		};

		citas.forEach((cita) => {
			const estado = normalize(cita.estado);
			const group = getInstituteGroup(cita.instituto);

			if (group === 'jvn') data.totalJVN += 1;
			if (group === 'ie') data.totalIE += 1;

			if (estado === 'realizada') {
				data.realizadas += 1;
				if (group === 'jvn') data.realizadasJVN += 1;
				if (group === 'ie') data.realizadasIE += 1;
			} else if (estado === 'no asistio') {
				data.noAsistio += 1;
				if (group === 'jvn') data.noAsistioJVN += 1;
				if (group === 'ie') data.noAsistioIE += 1;
			} else if (estado === 'cancelada') {
				data.canceladas += 1;
				if (group === 'jvn') data.canceladasJVN += 1;
				if (group === 'ie') data.canceladasIE += 1;
			} else if (estado === 'reprogramada') {
				data.reprogramadas += 1;
				if (group === 'jvn') data.reprogramadasJVN += 1;
				if (group === 'ie') data.reprogramadasIE += 1;
			}
		});

		ui.kpiTotal.textContent = data.total;
		ui.kpiRealizadas.textContent = data.realizadas;
		ui.kpiNoAsistio.textContent = data.noAsistio;
		ui.kpiCanceladas.textContent = data.canceladas;
		ui.kpiReprogramadas.textContent = data.reprogramadas;

		ui.kpiTotalJVN.textContent = data.totalJVN;
		ui.kpiTotalIE.textContent = data.totalIE;
		ui.kpiRealizadasJVN.textContent = data.realizadasJVN;
		ui.kpiRealizadasIE.textContent = data.realizadasIE;
		ui.kpiNoAsistioJVN.textContent = data.noAsistioJVN;
		ui.kpiNoAsistioIE.textContent = data.noAsistioIE;
		ui.kpiCanceladasJVN.textContent = data.canceladasJVN;
		ui.kpiCanceladasIE.textContent = data.canceladasIE;
		ui.kpiReprogramadasJVN.textContent = data.reprogramadasJVN;
		ui.kpiReprogramadasIE.textContent = data.reprogramadasIE;
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
				const modalidadLabel = getModalidadLabel(cita);
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
						<p class="text-xs text-gray-500 mt-1">${escapeHtml(fechaHora)}${modalidadLabel ? ` · <span class="font-semibold text-gray-700">${escapeHtml(modalidadLabel)}</span>` : ''}</p>
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

	function renderCalendar() {
		if (!ui.calendarGrid || !ui.calendarMonthLabel) return;

		const monthStart = getStartOfMonth(state.calendarCursor || new Date());
		state.calendarCursor = monthStart;

		const year = monthStart.getFullYear();
		const month = monthStart.getMonth();
		const daysInMonth = new Date(year, month + 1, 0).getDate();
		const firstDayOffset = (() => {
			const weekday = (monthStart.getDay() + 6) % 7;
			return weekday >= 5 ? 0 : weekday;
		})();

		ui.calendarMonthLabel.textContent = capitalizeFirst(monthStart.toLocaleDateString('es-PE', {
			month: 'long',
			year: 'numeric',
		}));
		ui.calendarMonthLabel.className = 'text-base font-semibold text-gray-900 tracking-tight';

		const citasByDate = groupCitasByDate(state.filtered);
		const todayIso = formatDateToIso_(new Date());

		const weekdays = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie'];
		const weekdayHeaderHtml = weekdays
			.map((d) => `<div class="text-center py-1 font-medium">${d}</div>`)
			.join('');

		const cells = [];
		for (let i = 0; i < firstDayOffset; i += 1) {
			cells.push('<div class="min-h-[72px] rounded-lg border border-transparent"></div>');
		}

		for (let day = 1; day <= daysInMonth; day += 1) {
			const weekdayIndex = new Date(year, month, day).getDay();
			if (weekdayIndex === 0 || weekdayIndex === 6) continue;

			const iso = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
			const citasDelDia = citasByDate.get(iso) || [];
			const hasCitas = citasDelDia.length > 0;
			const isToday = iso === todayIso;
			const isSelected = state.calendarSelectedDate === iso;
			const dayNumberClass = isToday
				? 'inline-flex items-center justify-center w-7 h-7 rounded-full bg-blue-100 text-blue-700 font-semibold'
				: 'text-[#374151] font-medium';
			const indicatorHtml = hasCitas ? buildCalendarCitaIndicator(citasDelDia) : '';

			cells.push(`
				<button
					type="button"
					data-date="${iso}"
					class="min-h-[74px] rounded-xl border text-left px-2 pt-0 pb-1.5 transition relative overflow-hidden ${
						hasCitas
							? 'border-gray-200 bg-white hover:bg-gray-50 hover:border-gray-300'
							: 'border-gray-200 bg-gray-50 text-gray-400 hover:bg-gray-100'
					} ${isSelected ? 'border-blue-300 bg-blue-50/50 ring-1 ring-blue-200' : ''}"
				>
					<div class="flex items-start justify-between gap-2 -mt-0.5">
						<span class="text-sm ${dayNumberClass}">${String(day).padStart(2, '0')}</span>
						${isToday ? '<span class="text-[10px] font-semibold text-blue-700">Hoy</span>' : ''}
					</div>
					${indicatorHtml}
				</button>
			`);
		}

		const remainder = cells.length % 5;
		if (remainder > 0) {
			for (let i = remainder; i < 5; i += 1) {
				cells.push('<div class="min-h-[72px] rounded-lg border border-transparent"></div>');
			}
		}

		const emptyHelp = !state.filtered.length
			? '<div class="col-span-5 mt-1 text-xs text-gray-500">No hay citas para los filtros actuales.</div>'
			: '';

		ui.calendarWeekdays.innerHTML = weekdayHeaderHtml;
		ui.calendarGrid.innerHTML = cells.join('') + emptyHelp;
		renderCalendarDayList();
	}

	function buildCalendarCitaIndicator(citasDelDia) {
		const maxDots = 4;
		const dots = (citasDelDia || []).slice(0, maxDots);
		const extraCount = Math.max((citasDelDia || []).length - maxDots, 0);
		const tooltip = buildCalendarTooltip(citasDelDia);

		const dotsHtml = dots
			.map((cita) => {
				const group = getInstituteGroup(cita?.instituto);
				const dotClass = group === 'jvn'
					? 'bg-purple-500'
					: group === 'ie'
						? 'bg-orange-500'
						: 'bg-gray-400';
				return `<span class="inline-block w-1.5 h-1.5 rounded-full ${dotClass}"></span>`;
			})
			.join('');

		return `
			<span title="${escapeAttr(tooltip)}" class="absolute bottom-1.5 left-2 inline-flex items-center gap-1 px-1.5 py-1 rounded-full bg-gray-100 text-[10px] text-gray-600">
				${dotsHtml}
				${extraCount > 0 ? `<span class="text-[10px] font-medium text-gray-500">+${extraCount}</span>` : ''}
			</span>
		`;
	}

	function buildCalendarTooltip(citasDelDia) {
		const total = (citasDelDia || []).length;
		const counts = { jvn: 0, iempresa: 0, other: 0 };

		(citasDelDia || []).forEach((cita) => {
			const group = getInstituteGroup(cita?.instituto);
			if (group === 'jvn') counts.jvn += 1;
			else if (group === 'ie') counts.iempresa += 1;
			else counts.other += 1;
		});

		const parts = [];
		if (counts.jvn) parts.push(`${counts.jvn} Neumann`);
		if (counts.iempresa) parts.push(`${counts.iempresa} IEmpresa`);
		if (counts.other) parts.push(`${counts.other} Otros`);

		return `${total} citas${parts.length ? ` – ${parts.join(', ')}` : ''}`;
	}

	function handleCalendarPrevMonth() {
		const base = getStartOfMonth(state.calendarCursor || new Date());
		state.calendarCursor = new Date(base.getFullYear(), base.getMonth() - 1, 1);
		renderCalendar();
	}

	function handleCalendarNextMonth() {
		const base = getStartOfMonth(state.calendarCursor || new Date());
		state.calendarCursor = new Date(base.getFullYear(), base.getMonth() + 1, 1);
		renderCalendar();
	}

	function handleCalendarDayClick(e) {
		const btn = e.target.closest('[data-date]');
		if (!btn) return;

		const iso = btn.dataset.date || '';
		if (!iso) return;

		state.calendarSelectedDate = iso;
		renderCalendar();
	}

	function renderCalendarDayList() {
		if (!ui.calendarDayList) return;

		const selectedIso = state.calendarSelectedDate;
		if (!selectedIso) {
			ui.calendarDayList.innerHTML = '<p class="text-xs text-gray-500">Selecciona un día para ver sus citas.</p>';
			return;
		}

		const citasDelDia = state.filtered
			.filter((cita) => normalizeDateValue(cita.fecha) === selectedIso)
			.sort((a, b) => getDateTimeValue(a.fecha, a.hora) - getDateTimeValue(b.fecha, b.hora));

		if (!citasDelDia.length) {
			ui.calendarDayList.innerHTML = `
				<div class="text-xs text-gray-500">
					Sin citas para ${escapeHtml(formatDateDisplay(selectedIso))}.
				</div>
			`;
			return;
		}

		const itemsHtml = citasDelDia
			.map((cita) => {
				const active = state.selectedId === cita.id;
				const modalidad = getModalidadLabel(cita);
				const instituteLabel = getInstituteDisplayLabel(cita.instituto);
				const instituteGroup = getInstituteGroup(cita.instituto);
				const instituteTagClass = instituteGroup === 'jvn'
					? 'bg-purple-100 text-purple-700'
					: instituteGroup === 'ie'
						? 'bg-orange-100 text-orange-700'
						: 'bg-gray-100 text-gray-700';
				return `
					<button
						type="button"
						data-agenda-id="${escapeAttr(cita.id)}"
						class="w-full text-left rounded-lg border px-2.5 py-2 mb-2 transition ${active ? 'border-brand-300 bg-brand-50' : 'border-gray-200 hover:bg-gray-50'}"
					>
						<div class="flex items-center justify-between gap-2">
							<p class="text-xs font-medium text-gray-900 truncate">${escapeHtml(cita.reservadoPor || 'Sin nombre')}</p>
						</div>
						<div class="mt-1 flex items-center gap-2 text-[11px] text-gray-500">
							<span>${escapeHtml(formatTimeDisplay(cita.hora))}</span>
							${modalidad ? `<span class="font-semibold text-gray-700">${escapeHtml(modalidad)}</span>` : ''}
							<span class="ml-auto inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-medium ${instituteTagClass}">${escapeHtml(instituteLabel)}</span>
						</div>
					</button>
				`;
			})
			.join('');

		ui.calendarDayList.innerHTML = `
			<div class="flex items-center justify-between mb-2">
				<p class="text-xs font-semibold text-gray-700">${escapeHtml(formatDateDisplay(selectedIso))}</p>
				<p class="text-[11px] text-gray-500">${citasDelDia.length} registros</p>
			</div>
			${itemsHtml}
		`;
	}

	function handleCalendarAgendaClick(e) {
		const btn = e.target.closest('[data-agenda-id]');
		if (!btn) return;

		state.selectedId = btn.dataset.agendaId;
		renderCalendarDayList();
		renderDetail();
	}

	function groupCitasByDate(citas) {
		const map = new Map();
		(citas || []).forEach((cita) => {
			const iso = normalizeDateValue(cita.fecha);
			if (!iso) return;
			if (!map.has(iso)) map.set(iso, []);
			map.get(iso).push(cita);
		});
		return map;
	}

	function normalizeDateValue(value) {
		const d = parseDateOnly(value);
		if (!d) return '';
		return formatDateToIso_(d);
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
		ui.resumenDni.textContent = cita.dni || cita.DNI || '—';
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
			linkOpcional: '',
		});
		renderReminderPanel(cita);

		hydrateForm(ui.personalForm, {
			reservadoPor: cita.reservadoPor,
			correo: cita.correo,
			telefono: cita.telefono,
			dni: cita.dni || cita.DNI,
			edad: cita.edad,
			ciclo: cita.ciclo,
			carrera: cita.carrera,
			conviveCon: cita.conviveCon,
			motivoConsulta: cita.motivoConsulta,
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
		payload.DNI = payload.dni;
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

	function handleDeleteCita() {
		const cita = getSelectedCita();
		if (!cita) return;

		const confirmed = window.confirm(
			`¿Seguro que deseas anular esta cita?\n\nAlumno: ${cita.reservadoPor || 'Sin nombre'}\nID: ${cita.id}\n\nEsta acción eliminará el registro.`
		);
		if (!confirmed) return;

		const deleteUrl = `${DELETE_CITA_WEBAPP_URL}?action=delete&id=${encodeURIComponent(cita.id)}`;
		setButtonLoading(ui.anularCitaBtn, true);

		fetch(deleteUrl, {
			method: 'GET',
			mode: 'no-cors',
			cache: 'no-store',
		})
			.then(() => runServer('getCitas'))
			.then((citasActualizadas) => {
				state.citas = Array.isArray(citasActualizadas) ? sortCitas(citasActualizadas) : [];
				state.visibleCount = LIST_BATCH_SIZE;

				const stillExists = state.citas.some((item) => item.id === cita.id);
				if (!stillExists) {
					state.selectedId = null;
					showToast('Cita anulada correctamente', 'success');
				} else {
					showToast('Solicitud enviada. Si aún aparece, pulsa Actualizar.', 'info');
				}

				applyFilters();
				renderKpis(state.citas);
				renderDetail();
			})
			.catch(() => {
				showToast('La cita pudo haberse anulado, pero el navegador no pudo confirmarlo. Pulsa Actualizar.', 'info');
			})
			.finally(() => setButtonLoading(ui.anularCitaBtn, false));
	}

	function handleSendReminder(e) {
		e.preventDefault();
		const cita = getSelectedCita();
		if (!cita) return;
		if (!canSendReminderByEstado(cita)) {
			showToast('Solo puedes enviar recordatorio cuando el estado está Programada o vacío.', 'info');
			return;
		}

		const { mensaje, linkOpcional } = getFormData(ui.reminderForm);
		if (!cita.correo) {
			showToast('La cita no tiene correo registrado', 'error');
			return;
		}

		const modalidadType = getModalidadType(cita);
		const isVirtual = modalidadType === 'virtual';
		const meetLink = isVirtual ? getMeetLinkByCita(cita) : '';

		const payload = {
			estudiante: cita.reservadoPor || '',
			correo: cita.correo || '',
			mensaje: mensaje || '',
			linkOpcional: linkOpcional || '',
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

	function adjustWorkspacePanelsHeight() {
		if (!ui.workspacePanels) return;

		const isDesktop = window.matchMedia('(min-width: 1280px)').matches;
		if (!isDesktop || !state.isAuthenticated) {
			ui.workspacePanels.style.height = '';
			return;
		}

		const rect = ui.workspacePanels.getBoundingClientRect();
		const viewportHeight = window.innerHeight || document.documentElement.clientHeight;
		const bottomGap = 12;
		const available = Math.floor(viewportHeight - rect.top - bottomGap);
		const targetHeight = Math.max(280, available);

		ui.workspacePanels.style.height = `${targetHeight}px`;
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
		const text = escapeHtml(getInstituteDisplayLabel(instituto));
		const group = getInstituteGroup(instituto);
		const cls = group === 'jvn'
			? 'bg-purple-100 text-purple-700'
			: group === 'ie'
				? 'bg-orange-100 text-orange-700'
				: 'bg-gray-100 text-gray-700';
		return `<span class="${cls} text-[11px] font-medium px-2 py-1 rounded-full">${text}</span>`;
	}

	function getModalidadLabel(cita) {
		const modalidadType = getModalidadType(cita);
		if (modalidadType === 'virtual') {
			return 'Virtual';
		}

		if (modalidadType === 'presencial') {
			return 'Presencial';
		}

		return '';
	}

	function getInstituteDisplayLabel(instituto) {
		const group = getInstituteGroup(instituto);
		if (group === 'jvn') return INSTITUTE_FILTER_CONFIG.jvn.fullLabel;
		if (group === 'ie') return INSTITUTE_FILTER_CONFIG.iempresa.fullLabel;
		return String(instituto || 'Sin instituto').trim() || 'Sin instituto';
	}

	function getInstituteGroup(instituto) {
		const inst = normalize(instituto);
		if (!inst) return 'other';

		if (
			inst.includes('jvn') ||
			inst.includes('neumann') ||
			inst.includes('jhonn') ||
			inst.includes('john') ||
			inst.includes('vonn') ||
			inst.includes('von neumann')
		) {
			return 'jvn';
		}

		if (inst.includes('iempresa') || inst.includes('instituto de la empresa') || inst.includes('empresa')) {
			return 'ie';
		}

		return 'other';
	}

	function getMeetLinkByCita(cita) {
		const group = getInstituteGroup(cita?.instituto);
		return MEET_LINKS_BY_INSTITUTE[group] || '';
	}

	function getModalidadType(cita) {
		const raw = cita?.modalidad ?? cita?.Modalidad ?? '';
		const value = normalize(raw);
		if (value === 'virtual') return 'virtual';
		if (value === 'presencial') return 'presencial';
		return 'sin-definir';
	}

	function canSendReminderByEstado(cita) {
		const estado = normalize(cita?.estado);
		return estado === '' || estado === 'programada';
	}

	function renderReminderPanel(cita) {
		if (!cita) return;

		const modalidadType = getModalidadType(cita);
		const meetLink = getMeetLinkByCita(cita);
		const canSend = canSendReminderByEstado(cita);

		if (modalidadType === 'virtual') {
			ui.reminderModalidadBadge.className = 'text-[11px] px-2 py-1 rounded-full font-medium bg-blue-100 text-blue-700';
			ui.reminderModalidadBadge.textContent = 'Modalidad: Virtual';
			ui.reminderMeetLabel.textContent = 'Este enlace se enviará automáticamente.';
			ui.reminderMeetLinkPreview.textContent = meetLink || 'Sin enlace configurado para este instituto.';

			if (meetLink) {
				ui.reminderMeetQuickAccess.href = meetLink;
				ui.reminderMeetQuickAccess.classList.remove('hidden');
			} else {
				ui.reminderMeetQuickAccess.classList.add('hidden');
			}
		} else if (modalidadType === 'presencial') {
			ui.reminderModalidadBadge.className = 'text-[11px] px-2 py-1 rounded-full font-medium bg-emerald-100 text-emerald-700';
			ui.reminderModalidadBadge.textContent = 'Modalidad: Presencial';
			ui.reminderMeetLabel.textContent = 'Esta cita es Presencial; no se enviará enlace de Meet.';
			ui.reminderMeetLinkPreview.textContent = '';
			ui.reminderMeetQuickAccess.classList.add('hidden');
		} else {
			ui.reminderModalidadBadge.className = 'text-[11px] px-2 py-1 rounded-full font-medium bg-gray-200 text-gray-700';
			ui.reminderModalidadBadge.textContent = 'Modalidad: Sin definir';
			ui.reminderMeetLabel.textContent = 'Defina si la cita es Virtual o Presencial para habilitar el acceso directo.';
			ui.reminderMeetLinkPreview.textContent = meetLink || 'Sin enlace configurado para este instituto.';
			ui.reminderMeetQuickAccess.classList.add('hidden');
		}

		ui.sendReminderBtn.disabled = !canSend;
		ui.reminderStateHelp.textContent = canSend
			? ''
			: 'Solo se puede enviar cuando aun está en "Programada".';
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

	function getStartOfMonth(value) {
		const d = value instanceof Date ? new Date(value) : new Date();
		d.setDate(1);
		d.setHours(0, 0, 0, 0);
		return d;
	}

	function getDefaultBusinessDayIso(value) {
		const date = value instanceof Date ? new Date(value) : new Date();
		const weekday = date.getDay();

		if (weekday === 6) {
			date.setDate(date.getDate() + 2);
		} else if (weekday === 0) {
			date.setDate(date.getDate() + 1);
		}

		return formatDateToIso_(date);
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
		return String(v || '')
			.normalize('NFD')
			.replace(/[\u0300-\u036f]/g, '')
			.toLowerCase()
			.trim();
	}

	function capitalizeFirst(value) {
		const text = String(value || '');
		if (!text) return '';
		return text.charAt(0).toUpperCase() + text.slice(1);
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
		requestAnimationFrame(adjustWorkspacePanelsHeight);
	}

	function persistAuthSession_() {
		sessionStorage.setItem(AUTH_SESSION_KEY, '1');
	}

	function clearAuthSession_() {
		sessionStorage.removeItem(AUTH_SESSION_KEY);
	}

	function hasActiveAuthSession_() {
		return sessionStorage.getItem(AUTH_SESSION_KEY) === '1';
	}
})();
