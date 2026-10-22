"""Desktop presentation only. Browser and extraction work stays in SearchWorker."""
import customtkinter as ctk

BG = '#0B1220'
PANEL = '#111D2E'
CARD = '#17263A'
BORDER = '#2A3C53'
TEXT = '#EEF4FC'
MUTED = '#A4B6CC'
ACCENT = '#25B9A7'

STAGES = {'import': 'File import', 'browser': 'Browser', 'link_resolution': 'Link resolution',
          'linkedin_profile': 'Profile', 'linkedin_company': 'Company page', 'maps_lookup': 'Maps',
          'public_company': 'Revenue / SIC / NAICS', 'email': 'Business email · 15s budget',
          'excel_save': 'Excel save', 'benchmark': 'Offline benchmark',
          'search_results': 'Sales search results', 'smart_validation': 'SMART validation'}


def reset_stages(app):
    for widgets in app.stage_widgets.values():
        widgets['bar'].stop()
        widgets['bar'].configure(mode='determinate')
        widgets['bar'].set(0)
        widgets['label'].configure(text='Waiting')
        widgets['running'] = False
        widgets['count'] = 0


def update_stage(app, name, state, detail=''):
    name = {'linkedin_profile_panel': 'linkedin_profile'}.get(name, name)
    widgets = getattr(app, 'stage_widgets', {}).get(name)
    if not widgets:
        return
    bar = widgets['bar']
    if state == 'running':
        if not widgets['running']:
            bar.configure(mode='indeterminate')
            bar.start()
        widgets['running'] = True
        text = detail or 'Working…'
    else:
        bar.stop()
        bar.configure(mode='determinate')
        bar.set(1 if state == 'done' else 0)
        widgets['running'] = False
        if state == 'done':
            widgets['count'] += 1
        text = (f"Finished · {widgets['count']} calls" if state == 'done' else state.capitalize())
        if detail:
            text += ' · ' + detail
    widgets['label'].configure(text=text)


def build_ui(app, version):
    app.configure(fg_color=BG)
    app.grid_columnconfigure(0, weight=0)
    app.grid_columnconfigure(1, weight=1)
    app.grid_rowconfigure(0, weight=1)

    def label(parent, text, size=13, bold=False, color=TEXT, **kwargs):
        return ctk.CTkLabel(parent, text=text, text_color=color,
                            font=ctk.CTkFont(family='Segoe UI', size=size,
                                            weight='bold' if bold else 'normal'), **kwargs)

    def card(parent, title):
        frame = ctk.CTkFrame(parent, fg_color=CARD, corner_radius=12,
                             border_width=1, border_color=BORDER)
        frame.pack(fill='x', pady=(0, 12))
        label(frame, title, 14, True).pack(anchor='w', padx=14, pady=(12, 9))
        return frame

    def button(parent, text, command, primary=False, **kwargs):
        return ctk.CTkButton(parent, text=text, command=command, height=36,
                            corner_radius=8, fg_color=ACCENT if primary else '#253B55',
                            hover_color='#199B8C' if primary else '#34516F',
                            text_color=BG if primary else TEXT,
                            font=ctk.CTkFont(family='Segoe UI', size=13, weight='bold'), **kwargs)

    def entry(parent, **kwargs):
        return ctk.CTkEntry(parent, height=36, fg_color=BG, border_color=BORDER,
                           text_color=TEXT, corner_radius=7, **kwargs)

    def check(parent, text, selected=False):
        control = ctk.CTkCheckBox(parent, text=text, text_color=TEXT, fg_color=ACCENT,
                                 hover_color='#199B8C', border_color=MUTED,
                                 checkbox_width=19, checkbox_height=19,
                                 font=ctk.CTkFont(family='Segoe UI', size=12))
        control.pack(anchor='w', padx=14, pady=(0, 12))
        if selected:
            control.select()
        return control

    app.settings_shell = left = ctk.CTkFrame(app, width=366, fg_color=PANEL, corner_radius=0)
    left.grid(row=0, column=0, sticky='nsew')
    left.grid_propagate(False)
    left.grid_columnconfigure(0, weight=1)
    left.grid_rowconfigure(1, weight=1)
    brand = ctk.CTkFrame(left, fg_color='transparent')
    brand.grid(row=0, column=0, sticky='ew', padx=18, pady=(18, 6))
    label(brand, 'LEAD EXTRACTOR', 23, True).pack(anchor='w')
    label(brand, f'LinkedIn research workspace   /   v{version}', 12, color=MUTED).pack(anchor='w')

    app.settings_tabs = tabs = ctk.CTkTabview(left, fg_color=PANEL, corner_radius=8,
        segmented_button_fg_color=BG, segmented_button_selected_color='#256477',
        segmented_button_selected_hover_color='#29798D',
        segmented_button_unselected_color=BG, text_color=TEXT)
    tabs.grid(row=1, column=0, sticky='nsew', padx=8, pady=(0, 8))
    areas = {}
    for name in ('Setup', 'Options', 'Tools'):
        panel = tabs.add(name)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(0, weight=1)
        scroll = ctk.CTkScrollableFrame(panel, fg_color='transparent', corner_radius=0,
                                      scrollbar_button_color=BORDER)
        scroll.grid(row=0, column=0, sticky='nsew')
        areas[name] = scroll

    setup = areas['Setup']
    connection = card(setup, '01   Connect your browser')
    app.browser_status = label(connection, 'Browser: Not opened', 12, color=MUTED,
                               anchor='w', wraplength=275, justify='left')
    app.browser_status.pack(fill='x', padx=14, pady=(0, 8))
    button(connection, 'Open Chrome / Sign In', app._open_browser).pack(fill='x', padx=12, pady=(0, 12))

    source = card(setup, '02   Choose your source')
    app.process_mode = ctk.CTkSegmentedButton(source, values=['Sales Navigator', 'Link Scraper'],
        command=app._on_process_mode_change, height=34, fg_color=BG,
        selected_color='#256477', selected_hover_color='#29798D',
        unselected_color=BG, text_color=TEXT)
    app.process_mode.pack(fill='x', padx=12, pady=(0, 10))
    app.process_mode.set('Link Scraper')
    app.sales_input_frame = ctk.CTkFrame(source, fg_color='transparent')
    app.search_url_label = label(app.sales_input_frame, 'Sales Navigator search URL', 12, anchor='w')
    app.search_url_label.pack(fill='x', pady=(0, 5))
    app.search_url_entry = entry(app.sales_input_frame, placeholder_text='Paste a Sales Navigator search URL')
    app.search_url_entry.pack(fill='x')
    app.link_input_frame = ctk.CTkFrame(source, fg_color='transparent')
    app.link_file_button = button(app.link_input_frame, 'Choose Excel / CSV Link File', app._choose_link_input)
    app.link_file_button.pack(fill='x', pady=(0, 6))
    app.link_file_label = label(app.link_input_frame, 'No link file selected', 12, color=MUTED,
                               anchor='w', justify='left', wraplength=270)
    app.link_file_label.pack(fill='x')
    app.link_count_label = label(app.link_input_frame, 'All sheets and columns are read.', 12,
                                 color=ACCENT, anchor='w', wraplength=270, justify='left')
    app.link_count_label.pack(fill='x', pady=(6, 0))
    label(app.link_input_frame, 'Mix normal profiles and Sales Navigator leads\nin one file. Links are resolved automatically.',
          12, color=MUTED, justify='left', anchor='w', wraplength=270).pack(fill='x', pady=(8, 0))

    output = card(setup, '03   Save your results')
    app.output_label = label(output, 'Save location not selected', 12, color=MUTED,
                            anchor='w', justify='left', wraplength=275)
    app.output_label.pack(fill='x', padx=14, pady=(0, 8))
    button(output, 'Choose Save Location', app._choose_output).pack(fill='x', padx=12, pady=(0, 10))
    row = ctk.CTkFrame(output, fg_color='transparent')
    row.pack(fill='x', padx=14, pady=(0, 12))
    label(row, 'Target saved leads', 12).pack(side='left')
    app.max_records_entry = entry(row, width=95)
    app.max_records_entry.pack(side='right')
    app.max_records_entry.insert(0, '25')
    app.all_links_option = check(output, 'All imported links (Link Scraper only)')
    label(output, 'Target limits saved leads, not file size.\nLinkedIn keeps its existing pace.', 12,
          color=MUTED, justify='left', wraplength=275).pack(anchor='w', padx=14, pady=(0, 10))
    app.worksheet_button = button(output, 'Worksheet · Add rows / columns', app._open_worksheet)
    app.worksheet_button.pack(fill='x', padx=12, pady=(0, 12))

    options = areas['Options']
    mode = card(options, 'Extraction & browser')
    app.extraction_mode = ctk.CTkSegmentedButton(mode, values=['FAST', 'SMART'], height=32,
        fg_color=BG, selected_color='#256477', selected_hover_color='#29798D', unselected_color=BG)
    app.extraction_mode.pack(fill='x', padx=12, pady=(0, 8))
    app.extraction_mode.set('FAST')
    label(mode, 'FAST uses page data and public sources.\nSMART also uses optional local Qwen validation.',
          12, color=MUTED, justify='left', wraplength=275).pack(anchor='w', padx=14, pady=(0, 12))
    app.headless_option = check(mode, 'Run browsers headless')
    app.parallel_enrichment = check(mode, 'Two concurrent enrichment sessions', True)
    label(mode, 'Sign in through Open Chrome first.\nLinkedIn uses one session with its existing waits.',
          12, color=MUTED, justify='left', wraplength=275).pack(anchor='w', padx=14, pady=(0, 12))
    quality = card(options, 'Location, email & saved data')
    app.local_office_option = check(quality, 'Find company office near prospect', True)
    app.refresh_profiles_option = check(quality, 'Refresh saved LinkedIn profiles')
    label(quality, 'Refresh also rechecks saved Sales Navigator link mappings.', 12,
          color=MUTED, justify='left', wraplength=275).pack(anchor='w', padx=14, pady=(0, 12))
    app.email_domain_checks = check(quality, 'Check email domains (DNS/MX)')
    label(quality, 'Optional: sends domains to Google DNS.\nDoes not verify individual mailboxes.', 12,
          color=MUTED, justify='left', wraplength=275).pack(anchor='w', padx=14, pady=(0, 12))

    tools = areas['Tools']
    benchmark = card(tools, 'Review data quality')
    label(benchmark, 'Offline sample size (1–200)', 12).pack(anchor='w', padx=14, pady=(0, 5))
    app.benchmark_sample_entry = entry(benchmark)
    app.benchmark_sample_entry.pack(fill='x', padx=12, pady=(0, 8))
    app.benchmark_sample_entry.insert(0, '150')
    app.benchmark_button = button(benchmark, 'Build Offline Benchmark', app._build_offline_benchmark)
    app.benchmark_button.pack(fill='x', padx=12, pady=(0, 8))
    label(benchmark, 'Uses a saved workbook. No website requests.', 12, color=MUTED,
          wraplength=275, justify='left').pack(anchor='w', padx=14, pady=(0, 12))
    system = card(tools, 'Connection tools')
    app.llm_status = label(system, 'LLM: Checking...', 12, color=MUTED, wraplength=275, justify='left')
    app.llm_status.pack(anchor='w', padx=14, pady=(0, 10))
    button(system, 'Clear Saved Browser Session', app._clear_session).pack(fill='x', padx=12, pady=(0, 12))
    guide = card(tools, 'Where to find your results')
    label(guide, 'Leads: extracted profile and company data.\n\nLink Resolution: original links, public URLs, duplicates and unresolved links.\n\nEvidence sheets: sources and supporting details.',
          12, color=MUTED, wraplength=275, justify='left').pack(anchor='w', padx=14, pady=(0, 12))

    app.action_card = action = ctk.CTkFrame(left, fg_color=CARD, corner_radius=12)
    action.grid(row=2, column=0, sticky='ew', padx=16, pady=(0, 16))
    app.start_btn = button(action, 'START LINK SCRAPER', app._start_processing, primary=True)
    app.start_btn.configure(height=42)
    app.start_btn.pack(fill='x', padx=12, pady=(12, 8))
    row = ctk.CTkFrame(action, fg_color='transparent')
    row.pack(fill='x', padx=12, pady=(0, 12))
    row.grid_columnconfigure((0, 1, 2), weight=1, uniform='action')
    app.run_controls = {}
    for col, (text, callback) in enumerate((('Pause', app._pause), ('Resume', app._resume), ('Stop', app._stop))):
        control = button(row, text, callback, width=70)
        control.grid(row=0, column=col, sticky='ew', padx=3)
        control.configure(state='disabled')
        app.run_controls[text] = control

    right = ctk.CTkFrame(app, fg_color=BG, corner_radius=0)
    right.grid(row=0, column=1, sticky='nsew', padx=22, pady=20)
    right.grid_columnconfigure(0, weight=1)
    right.grid_rowconfigure(4, weight=1)
    header = ctk.CTkFrame(right, fg_color='transparent')
    header.grid(row=0, column=0, sticky='ew', pady=(0, 14))
    label(header, 'Your extraction workspace', 25, True).pack(anchor='w')
    label(header, 'Import links. Enrich profiles. Review the evidence.', 13, color=MUTED).pack(anchor='w')
    status = ctk.CTkFrame(right, fg_color=PANEL, corner_radius=10)
    status.grid(row=1, column=0, sticky='ew', pady=(0, 12))
    status.grid_columnconfigure(1, weight=1)
    label(status, '●', 18, color=ACCENT).grid(row=0, column=0, padx=(14, 8), pady=10)
    app.status_label = label(status, 'Ready · choose a link file to begin', 13, anchor='w',
                             wraplength=480, justify='left')
    app.status_label.grid(row=0, column=1, sticky='ew', padx=(0, 12), pady=10)
    counts = ctk.CTkFrame(right, fg_color='transparent')
    counts.grid(row=2, column=0, sticky='ew', pady=(0, 12))
    counts.grid_columnconfigure((0, 1, 2, 3, 4), weight=1, uniform='counter')
    app.counter_labels = {}
    for col, key in enumerate(('Processed', 'Saved', 'Duplicates', 'Failed', 'Remaining')):
        box = ctk.CTkFrame(counts, fg_color=PANEL, corner_radius=10)
        box.grid(row=0, column=col, sticky='ew', padx=(0 if col == 0 else 4, 0 if col == 4 else 4))
        value = label(box, '0', 27, True, ACCENT if key == 'Saved' else '#F6BE72' if key == 'Failed' else TEXT)
        value.pack(pady=(12, 0))
        label(box, key, 11, color=MUTED).pack(pady=(0, 12))
        app.counter_labels[key] = value
    progress_frame = ctk.CTkFrame(right, fg_color='transparent')
    progress_frame.grid(row=3, column=0, sticky='ew', pady=(0, 14))
    app.harvest_label = label(progress_frame, 'Leads saved: 0', 12, color=MUTED, anchor='w')
    app.harvest_label.pack(fill='x', pady=(0, 5))
    app.progress = ctk.CTkProgressBar(progress_frame, height=7, fg_color=CARD, progress_color=ACCENT)
    app.progress.pack(fill='x')
    app.progress.set(0)
    app.workspace_tabs = workspace = ctk.CTkTabview(right, fg_color=PANEL,
        segmented_button_selected_color='#256477', segmented_button_unselected_color=BG)
    workspace.grid(row=4, column=0, sticky='nsew')
    stage_tab = workspace.add('Process progress')
    stage_tab.grid_columnconfigure(0, weight=1)
    stage_tab.grid_rowconfigure(0, weight=1)
    stages = ctk.CTkScrollableFrame(stage_tab, fg_color='transparent')
    stages.grid(row=0, column=0, sticky='nsew')
    stages.grid_columnconfigure((0, 1), weight=1, uniform='stage')
    app.stage_widgets = {}
    for index, (name, title) in enumerate(STAGES.items()):
        box = ctk.CTkFrame(stages, fg_color=CARD)
        box.grid(row=index // 2, column=index % 2, sticky='ew', padx=5, pady=5)
        label(box, title, 12, True).pack(anchor='w', padx=10, pady=(8, 4))
        bar = ctk.CTkProgressBar(box, height=5, progress_color=ACCENT)
        bar.pack(fill='x', padx=10, pady=(0, 4))
        status_label = label(box, 'Waiting', 11, color=MUTED, anchor='w', wraplength=240, justify='left')
        status_label.pack(fill='x', padx=10, pady=(0, 8))
        app.stage_widgets[name] = {'bar': bar, 'label': status_label, 'running': False, 'count': 0}
    reset_stages(app)
    preview_frame = workspace.add('Latest lead')
    label(preview_frame, 'Latest extracted lead', 14, True).pack(anchor='w', padx=16, pady=(12, 5))
    app.preview = ctk.CTkTextbox(preview_frame, height=140, fg_color=PANEL, text_color=TEXT,
                                font=ctk.CTkFont(family='Segoe UI', size=13), wrap='word')
    app.preview.pack(fill='both', expand=True, padx=10, pady=(0, 10))
    app.preview.insert('1.0', 'Your latest result will appear here.\n\nChoose an Excel or CSV file containing normal LinkedIn profiles,\nSales Navigator leads, or both.')
    app.preview.configure(state='disabled')
    log = workspace.add('Activity')
    log.grid_columnconfigure(0, weight=1)
    log.grid_rowconfigure(1, weight=1)
    app.workspace_note = label(log, 'Activity', 14, True, anchor='w')
    app.workspace_note.grid(row=0, column=0, sticky='ew', padx=16, pady=(12, 5))
    app.log_box = ctk.CTkTextbox(log, fg_color=PANEL, text_color=MUTED,
                                font=ctk.CTkFont(family='Consolas', size=12), wrap='word')
    app.log_box.grid(row=1, column=0, sticky='nsew', padx=10, pady=(0, 10))
    app.log_box.configure(state='disabled')
    app._on_process_mode_change(app.process_mode.get())
