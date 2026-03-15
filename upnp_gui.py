import ipaddress
import socket
import threading
import tkinter as tk
from tkinter import font as tkfont
from tkinter import messagebox, scrolledtext, ttk
from typing import Any

import upnpy


class UPnPToolGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("UPnP ポート開放ツール")
        self.root.geometry("980x760")
        self.root.minsize(860, 680)
        self.root.configure(bg="#eef3f9")

        self.upnp = upnpy.UPnP()
        self.service = None
        self.device = None
        self.service_candidates: list[dict[str, Any]] = []
        self.mapping_rows: dict[str, dict[str, str]] = {}
        self.busy = False

        self.status_var = tk.StringVar(value="UPnP デバイスを検出中...")
        self.external_ip_var = tk.StringVar(value="未取得")
        self.router_var = tk.StringVar()
        self.router_info_var = tk.StringVar(value="未選択")

        self.setup_styles()
        self.create_widgets()
        self.auto_fill_ip()
        self.start_discovery()

    def setup_styles(self) -> None:
        self.palette = {
            "app_bg": "#eef3f9",
            "card_bg": "#ffffff",
            "border": "#d8e1ee",
            "text": "#122033",
            "muted": "#5b6b82",
            "accent": "#2563eb",
            "accent_hover": "#1d4ed8",
            "danger": "#dc2626",
            "success": "#15803d",
            "log_bg": "#0f172a",
            "log_fg": "#e5eefc",
        }
        self.font_family = "Yu Gothic UI"

        try:
            self.root.tk.call("tk", "scaling", 1.1)
        except tk.TclError:
            pass

        for font_name, size, weight in (
            ("TkDefaultFont", 10, "normal"),
            ("TkTextFont", 10, "normal"),
            ("TkHeadingFont", 10, "bold"),
            ("TkMenuFont", 10, "normal"),
            ("TkCaptionFont", 10, "normal"),
        ):
            try:
                named_font = tkfont.nametofont(font_name)
                named_font.configure(family=self.font_family, size=size, weight=weight)
            except tk.TclError:
                continue

        self.root.option_add("*Font", f"{self.font_family} 10")
        self.root.option_add("*TCombobox*Listbox.font", f"{self.font_family} 10")

        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", background=self.palette["app_bg"], foreground=self.palette["text"])
        style.configure("App.TFrame", background=self.palette["app_bg"])
        style.configure(
            "Card.TFrame",
            background=self.palette["card_bg"],
            relief="solid",
            borderwidth=1,
            bordercolor=self.palette["border"],
        )
        style.configure(
            "Card.TLabelframe",
            background=self.palette["card_bg"],
            relief="solid",
            borderwidth=1,
            bordercolor=self.palette["border"],
            padding=14,
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=self.palette["card_bg"],
            foreground=self.palette["text"],
            font=(self.font_family, 11, "bold"),
        )
        style.configure(
            "Title.TLabel",
            background=self.palette["app_bg"],
            foreground=self.palette["text"],
            font=(self.font_family, 20, "bold"),
        )
        style.configure(
            "Subtitle.TLabel",
            background=self.palette["app_bg"],
            foreground=self.palette["muted"],
            font=(self.font_family, 10),
        )
        style.configure(
            "FieldLabel.TLabel",
            background=self.palette["card_bg"],
            foreground=self.palette["muted"],
            font=(self.font_family, 10, "bold"),
        )
        style.configure(
            "ValueLabel.TLabel",
            background=self.palette["card_bg"],
            foreground=self.palette["text"],
            font=(self.font_family, 11, "bold"),
        )
        style.configure(
            "SectionText.TLabel",
            background=self.palette["card_bg"],
            foreground=self.palette["muted"],
            font=(self.font_family, 10),
        )
        style.configure("TButton", padding=(14, 10), font=(self.font_family, 10, "bold"))
        style.configure(
            "Primary.TButton",
            background=self.palette["accent"],
            foreground="#ffffff",
            bordercolor=self.palette["accent"],
        )
        style.map(
            "Primary.TButton",
            background=[("active", self.palette["accent_hover"]), ("disabled", "#c7d2fe")],
            foreground=[("disabled", "#f8fafc")],
        )
        style.configure(
            "Secondary.TButton",
            background="#f8fbff",
            foreground=self.palette["text"],
            bordercolor=self.palette["border"],
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#e8f1ff"), ("disabled", "#f1f5f9")],
            foreground=[("disabled", "#94a3b8")],
        )
        style.configure(
            "Danger.TButton",
            background="#fee2e2",
            foreground=self.palette["danger"],
            bordercolor="#fecaca",
        )
        style.map(
            "Danger.TButton",
            background=[("active", "#fecaca"), ("disabled", "#fef2f2")],
            foreground=[("disabled", "#fca5a5")],
        )
        style.configure(
            "TEntry",
            fieldbackground="#ffffff",
            bordercolor=self.palette["border"],
            lightcolor=self.palette["border"],
            darkcolor=self.palette["border"],
            padding=(10, 8),
        )
        style.configure(
            "TCombobox",
            fieldbackground="#ffffff",
            bordercolor=self.palette["border"],
            lightcolor=self.palette["border"],
            darkcolor=self.palette["border"],
            arrowsize=16,
            padding=(8, 8),
        )
        style.configure(
            "Treeview",
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground=self.palette["text"],
            rowheight=30,
            font=(self.font_family, 10),
        )
        style.configure(
            "Treeview.Heading",
            background="#e8eef8",
            foreground=self.palette["text"],
            font=(self.font_family, 10, "bold"),
            relief="flat",
            padding=(8, 8),
        )
        style.map("Treeview", background=[("selected", "#dbeafe")], foreground=[("selected", self.palette["text"])])
        style.map("Treeview.Heading", background=[("active", "#dbeafe")])

    def create_widgets(self) -> None:
        main_frame = ttk.Frame(self.root, padding=18, style="App.TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)
        main_frame.rowconfigure(6, weight=1)

        header_frame = ttk.Frame(main_frame, style="App.TFrame")
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header_frame.columnconfigure(0, weight=1)

        ttk.Label(header_frame, text="UPnP ポート開放ツール", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header_frame,
            text="ルーター検出、ポート開放、既存マッピング確認を 1 画面で行えます。",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        status_frame = ttk.Frame(main_frame, padding=14, style="Card.TFrame")
        status_frame.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        status_frame.columnconfigure(0, weight=1)

        self.status_label = tk.Label(
            status_frame,
            textvariable=self.status_var,
            bg="#dbeafe",
            fg=self.palette["accent"],
            padx=14,
            pady=8,
            font=(self.font_family, 10, "bold"),
            anchor="w",
        )
        self.status_label.grid(row=0, column=0, sticky="w")

        self.refresh_devices_btn = ttk.Button(
            status_frame,
            text="ルーターを再検索",
            command=self.start_discovery,
            style="Secondary.TButton",
        )
        self.refresh_devices_btn.grid(row=0, column=1, sticky="e")

        router_frame = ttk.LabelFrame(main_frame, text="ルーター情報", style="Card.TLabelframe")
        router_frame.grid(row=2, column=0, pady=(0, 12), sticky="ew")
        router_frame.columnconfigure(1, weight=1)

        ttk.Label(router_frame, text="対象ルーター", style="FieldLabel.TLabel").grid(row=0, column=0, sticky="w")
        self.router_combo = ttk.Combobox(router_frame, textvariable=self.router_var, state="disabled")
        self.router_combo.grid(row=0, column=1, padx=(10, 0), pady=(0, 10), sticky="ew")
        self.router_combo.bind("<<ComboboxSelected>>", self.on_router_selected)

        ttk.Label(router_frame, text="検出したサービス", style="FieldLabel.TLabel").grid(row=1, column=0, sticky="w")
        ttk.Label(router_frame, textvariable=self.router_info_var, style="ValueLabel.TLabel").grid(
            row=1,
            column=1,
            padx=(10, 0),
            pady=(0, 10),
            sticky="w",
        )

        ttk.Label(router_frame, text="グローバル IP", style="FieldLabel.TLabel").grid(row=2, column=0, sticky="w")
        self.external_ip_label = ttk.Label(router_frame, textvariable=self.external_ip_var, style="ValueLabel.TLabel")
        self.external_ip_label.grid(row=2, column=1, padx=(10, 0), sticky="w")

        config_frame = ttk.LabelFrame(main_frame, text="ポート設定", style="Card.TLabelframe")
        config_frame.grid(row=3, column=0, pady=(0, 12), sticky="ew")
        config_frame.columnconfigure(1, weight=1)
        config_frame.columnconfigure(3, weight=1)

        ttk.Label(
            config_frame,
            text="一覧から選んだ内容をそのまま反映できるので、追加も削除も迷いにくくしています。",
            style="SectionText.TLabel",
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))

        ttk.Label(config_frame, text="外部ポート", style="FieldLabel.TLabel").grid(row=1, column=0, sticky="w")
        self.ext_port = ttk.Entry(config_frame)
        self.ext_port.insert(0, "8080")
        self.ext_port.grid(row=1, column=1, padx=(10, 18), pady=(0, 10), sticky="ew")

        ttk.Label(config_frame, text="内部ポート", style="FieldLabel.TLabel").grid(row=1, column=2, sticky="w")
        self.int_port = ttk.Entry(config_frame)
        self.int_port.insert(0, "8080")
        self.int_port.grid(row=1, column=3, padx=(10, 0), pady=(0, 10), sticky="ew")

        ttk.Label(config_frame, text="プロトコル", style="FieldLabel.TLabel").grid(row=2, column=0, sticky="w")
        self.protocol = ttk.Combobox(config_frame, values=["TCP", "UDP"], state="readonly")
        self.protocol.current(0)
        self.protocol.grid(row=2, column=1, padx=(10, 18), pady=(0, 10), sticky="ew")

        ttk.Label(config_frame, text="内部 IP", style="FieldLabel.TLabel").grid(row=2, column=2, sticky="w")
        self.local_ip = ttk.Entry(config_frame)
        self.local_ip.grid(row=2, column=3, padx=(10, 0), pady=(0, 10), sticky="ew")

        ttk.Label(config_frame, text="説明", style="FieldLabel.TLabel").grid(row=3, column=0, sticky="w")
        self.desc = ttk.Entry(config_frame)
        self.desc.insert(0, "UPnP Tool")
        self.desc.grid(row=3, column=1, columnspan=3, padx=(10, 0), sticky="ew")

        action_frame = ttk.Frame(main_frame, style="App.TFrame")
        action_frame.grid(row=4, column=0, pady=(0, 12), sticky="ew")
        action_frame.columnconfigure(0, weight=1)
        action_frame.columnconfigure(1, weight=1)
        action_frame.columnconfigure(2, weight=1)

        self.open_btn = ttk.Button(
            action_frame,
            text="ポートを開放 / 更新",
            command=self.start_add_mapping,
            state="disabled",
            style="Primary.TButton",
        )
        self.open_btn.grid(row=0, column=0, padx=(0, 8), sticky="ew")

        self.refresh_mappings_btn = ttk.Button(
            action_frame,
            text="一覧を更新",
            command=self.start_refresh_mappings,
            state="disabled",
            style="Secondary.TButton",
        )
        self.refresh_mappings_btn.grid(row=0, column=1, padx=8, sticky="ew")

        self.close_btn = ttk.Button(
            action_frame,
            text="フォーム内容で削除",
            command=self.start_delete_mapping,
            state="disabled",
            style="Danger.TButton",
        )
        self.close_btn.grid(row=0, column=2, padx=(8, 0), sticky="ew")

        mapping_frame = ttk.LabelFrame(main_frame, text="現在のポートマッピング", style="Card.TLabelframe")
        mapping_frame.grid(row=5, column=0, pady=(0, 12), sticky="nsew")
        mapping_frame.columnconfigure(0, weight=1)
        mapping_frame.rowconfigure(1, weight=1)

        ttk.Label(
            mapping_frame,
            text="一覧を選択するとフォームに反映されます。削除前の確認にも使えます。",
            style="SectionText.TLabel",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        columns = ("ext", "proto", "client", "int", "enabled", "lease", "desc")
        self.mapping_tree = ttk.Treeview(mapping_frame, columns=columns, show="headings", height=12)
        self.mapping_tree.heading("ext", text="外部")
        self.mapping_tree.heading("proto", text="プロトコル")
        self.mapping_tree.heading("client", text="内部 IP")
        self.mapping_tree.heading("int", text="内部")
        self.mapping_tree.heading("enabled", text="有効")
        self.mapping_tree.heading("lease", text="リース")
        self.mapping_tree.heading("desc", text="説明")

        self.mapping_tree.column("ext", width=90, anchor=tk.CENTER)
        self.mapping_tree.column("proto", width=110, anchor=tk.CENTER)
        self.mapping_tree.column("client", width=170, anchor=tk.CENTER)
        self.mapping_tree.column("int", width=90, anchor=tk.CENTER)
        self.mapping_tree.column("enabled", width=70, anchor=tk.CENTER)
        self.mapping_tree.column("lease", width=90, anchor=tk.CENTER)
        self.mapping_tree.column("desc", width=320, anchor=tk.W)
        self.mapping_tree.grid(row=1, column=0, sticky="nsew")
        self.mapping_tree.bind("<<TreeviewSelect>>", self.on_mapping_selected)

        tree_y_scroll = ttk.Scrollbar(mapping_frame, orient=tk.VERTICAL, command=self.mapping_tree.yview)
        tree_y_scroll.grid(row=1, column=1, sticky="ns")
        tree_x_scroll = ttk.Scrollbar(mapping_frame, orient=tk.HORIZONTAL, command=self.mapping_tree.xview)
        tree_x_scroll.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        self.mapping_tree.configure(yscrollcommand=tree_y_scroll.set, xscrollcommand=tree_x_scroll.set)

        log_frame = ttk.LabelFrame(main_frame, text="ログ", style="Card.TLabelframe")
        log_frame.grid(row=6, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)

        log_header = ttk.Frame(log_frame, style="Card.TFrame")
        log_header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        log_header.columnconfigure(0, weight=1)

        ttk.Label(
            log_header,
            text="検出結果や失敗理由をここに表示します。",
            style="SectionText.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Button(log_header, text="ログをクリア", command=self.clear_log, style="Secondary.TButton").grid(
            row=0,
            column=1,
            sticky="e",
        )

        self.log_area = scrolledtext.ScrolledText(
            log_frame,
            height=8,
            state="disabled",
            wrap=tk.WORD,
            bg=self.palette["log_bg"],
            fg=self.palette["log_fg"],
            insertbackground="#ffffff",
            relief="flat",
            borderwidth=0,
            padx=12,
            pady=12,
            font=("Consolas", 10),
        )
        self.log_area.grid(row=1, column=0, sticky="nsew")

    def clear_log(self) -> None:
        self.log_area.configure(state="normal")
        self.log_area.delete("1.0", tk.END)
        self.log_area.configure(state="disabled")

    def set_busy(self, busy: bool) -> None:
        self.busy = busy
        self.root.after(0, self._update_control_states)

    def _update_control_states(self) -> None:
        has_service = self.service is not None
        base_state = "disabled" if self.busy else "normal"

        self.refresh_devices_btn.config(state=base_state)
        self.open_btn.config(state="normal" if has_service and not self.busy else "disabled")
        self.close_btn.config(state="normal" if has_service and not self.busy else "disabled")
        self.refresh_mappings_btn.config(state="normal" if has_service and not self.busy else "disabled")

        if not self.service_candidates:
            self.router_combo.config(state="disabled")
        elif self.busy:
            self.router_combo.config(state="disabled")
        else:
            self.router_combo.config(state="readonly")

    def set_status(self, text: str, color: str = "#1d4ed8") -> None:
        status_backgrounds = {
            "#2563eb": "#dbeafe",
            "#1d4ed8": "#dbeafe",
            "#15803d": "#dcfce7",
            "#b91c1c": "#fee2e2",
        }

        def apply() -> None:
            self.status_var.set(text)
            self.status_label.config(
                foreground=color,
                background=status_backgrounds.get(color, "#e2e8f0"),
            )

        self.root.after(0, apply)

    def set_external_ip(self, value: str) -> None:
        self.root.after(0, lambda: self.external_ip_var.set(value))

    def set_router_info(self, value: str) -> None:
        self.root.after(0, lambda: self.router_info_var.set(value))

    def log(self, message: str) -> None:
        self.root.after(0, self._append_log, message)

    def _append_log(self, message: str) -> None:
        self.log_area.configure(state="normal")
        self.log_area.insert(tk.END, f"> {message}\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state="disabled")

    def show_error(self, title: str, message: str) -> None:
        self.root.after(0, lambda: messagebox.showerror(title, message))

    def show_info(self, title: str, message: str) -> None:
        self.root.after(0, lambda: messagebox.showinfo(title, message))

    def auto_fill_ip(self) -> None:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
                sock.connect(("8.8.8.8", 80))
                ip = sock.getsockname()[0]
            self.local_ip.delete(0, tk.END)
            self.local_ip.insert(0, ip)
        except OSError:
            self.log("内部 IP の自動取得に失敗しました。必要なら手動で入力してください。")

    def start_worker(self, target, *args) -> None:
        self.set_busy(True)
        threading.Thread(target=self._run_worker, args=(target, *args), daemon=True).start()

    def _run_worker(self, target, *args) -> None:
        try:
            target(*args)
        finally:
            self.set_busy(False)

    def start_discovery(self) -> None:
        self.start_worker(self.discover_devices)

    def discover_devices(self) -> None:
        self.set_status("UPnP デバイスを検索中...", "#2563eb")
        self.set_router_info("検索中")
        self.log("UPnP デバイスの検索を開始しました。")

        try:
            devices = self.upnp.discover()
        except Exception as exc:
            self.service = None
            self.device = None
            self.service_candidates = []
            self.update_router_candidates([])
            self.clear_mappings()
            self.set_external_ip("未取得")
            self.set_router_info("検索失敗")
            self.set_status("デバイス検索に失敗しました", "#b91c1c")
            self.log(f"検索エラー: {exc}")
            return

        candidates = self.find_service_candidates(devices)
        self.service_candidates = candidates
        self.update_router_candidates(candidates)

        if not candidates:
            self.service = None
            self.device = None
            self.clear_mappings()
            self.set_external_ip("未取得")
            self.set_router_info("対応サービスなし")
            self.set_status("対応する UPnP ルーターが見つかりません", "#b91c1c")
            self.log("ポートマッピングに使える WANIP/WANPPP サービスが見つかりませんでした。")
            return

        self.log(f"{len(candidates)} 件の候補が見つかりました。")
        for candidate in candidates:
            self.log(f"候補: {candidate['label']}")

        self.activate_candidate(0)

    def find_service_candidates(self, devices: list[Any]) -> list[dict[str, Any]]:
        candidates: list[dict[str, Any]] = []
        priority_keys = ("wanipconn", "wanipconnection", "wanpppconn", "wanpppconnection")

        for device in devices:
            friendly_name = getattr(device, "friendly_name", "Unknown device")
            try:
                services = device.get_services()
            except Exception as exc:
                self.log(f"{friendly_name} のサービス取得に失敗しました: {exc}")
                continue

            for service in services:
                service_id = getattr(service, "id", "") or getattr(service, "type_", "")
                lowered = str(service_id).lower()
                score = 999

                for index, key in enumerate(priority_keys):
                    if key in lowered:
                        score = index
                        break

                if score == 999 and "connection" not in lowered:
                    continue

                label = f"{friendly_name} | {service_id}"
                candidates.append(
                    {
                        "device": device,
                        "service": service,
                        "label": label,
                        "score": score,
                    }
                )

        candidates.sort(key=lambda item: (item["score"], item["label"]))
        return candidates

    def update_router_candidates(self, candidates: list[dict[str, Any]]) -> None:
        def apply() -> None:
            values = [candidate["label"] for candidate in candidates]
            self.router_combo["values"] = values
            if values:
                self.router_var.set(values[0])
                self.router_combo.current(0)
            else:
                self.router_var.set("")
                self.router_combo.set("")
            self._update_control_states()

        self.root.after(0, apply)

    def on_router_selected(self, _event=None) -> None:
        index = self.router_combo.current()
        if index < 0:
            return
        self.start_worker(self.activate_candidate, index)

    def activate_candidate(self, index: int) -> None:
        if index < 0 or index >= len(self.service_candidates):
            return

        candidate = self.service_candidates[index]
        self.device = candidate["device"]
        self.service = candidate["service"]

        friendly_name = getattr(self.device, "friendly_name", "Unknown device")
        self.set_status(f"接続中: {friendly_name}", "#2563eb")
        self.set_router_info(candidate["label"])
        self.log(f"ルーターを選択しました: {candidate['label']}")

        external_ip = "取得失敗"
        try:
            response = self.service.GetExternalIPAddress()
            external_ip = response.get("NewExternalIPAddress", "取得失敗")
        except Exception as exc:
            self.log(f"グローバル IP の取得に失敗しました: {exc}")

        self.set_external_ip(external_ip)
        self.set_status(f"接続済み: {friendly_name} ({external_ip})", "#15803d")
        self.refresh_mappings()

    def clear_mappings(self) -> None:
        def apply() -> None:
            self.mapping_rows.clear()
            for item_id in self.mapping_tree.get_children():
                self.mapping_tree.delete(item_id)

        self.root.after(0, apply)

    def start_refresh_mappings(self) -> None:
        if self.service is None:
            self.show_error("未接続", "先に利用するルーターを選択してください。")
            return
        self.start_worker(self.refresh_mappings)

    def refresh_mappings(self) -> None:
        if self.service is None:
            return

        self.log("ポートマッピング一覧を取得しています。")
        mappings = self.fetch_mappings()
        self.update_mapping_tree(mappings)
        self.log(f"ポートマッピングを {len(mappings)} 件読み込みました。")

    def fetch_mappings(self) -> list[dict[str, str]]:
        mappings: list[dict[str, str]] = []
        index = 0

        while True:
            try:
                entry = self.service.GetGenericPortMappingEntry(NewPortMappingIndex=index)
            except Exception:
                break

            mappings.append(
                {
                    "external_port": str(entry.get("NewExternalPort", "")),
                    "protocol": str(entry.get("NewProtocol", "")),
                    "internal_client": str(entry.get("NewInternalClient", "")),
                    "internal_port": str(entry.get("NewInternalPort", "")),
                    "enabled": "Yes" if str(entry.get("NewEnabled", "")) in {"1", "True", "true"} else "No",
                    "lease_duration": str(entry.get("NewLeaseDuration", "")),
                    "description": str(entry.get("NewPortMappingDescription", "")),
                }
            )
            index += 1

        mappings.sort(key=lambda item: (int(item["external_port"] or 0), item["protocol"]))
        return mappings

    def update_mapping_tree(self, mappings: list[dict[str, str]]) -> None:
        def apply() -> None:
            self.mapping_rows.clear()
            for item_id in self.mapping_tree.get_children():
                self.mapping_tree.delete(item_id)

            for index, mapping in enumerate(mappings):
                item_id = str(index)
                self.mapping_rows[item_id] = mapping
                self.mapping_tree.insert(
                    "",
                    tk.END,
                    iid=item_id,
                    values=(
                        mapping["external_port"],
                        mapping["protocol"],
                        mapping["internal_client"],
                        mapping["internal_port"],
                        mapping["enabled"],
                        mapping["lease_duration"],
                        mapping["description"],
                    ),
                )

        self.root.after(0, apply)

    def on_mapping_selected(self, _event=None) -> None:
        selection = self.mapping_tree.selection()
        if not selection:
            return

        mapping = self.mapping_rows.get(selection[0])
        if not mapping:
            return

        self.ext_port.delete(0, tk.END)
        self.ext_port.insert(0, mapping["external_port"])

        self.int_port.delete(0, tk.END)
        self.int_port.insert(0, mapping["internal_port"])

        self.protocol.set(mapping["protocol"] or "TCP")

        self.local_ip.delete(0, tk.END)
        self.local_ip.insert(0, mapping["internal_client"])

        self.desc.delete(0, tk.END)
        self.desc.insert(0, mapping["description"])

    def parse_port(self, value: str, label: str) -> int:
        try:
            port = int(value)
        except ValueError as exc:
            raise ValueError(f"{label}は数値で入力してください。") from exc

        if not 1 <= port <= 65535:
            raise ValueError(f"{label}は 1 から 65535 の範囲で入力してください。")
        return port

    def validate_ip(self, value: str) -> str:
        try:
            return str(ipaddress.IPv4Address(value))
        except ipaddress.AddressValueError as exc:
            raise ValueError("内部 IP は有効な IPv4 アドレスを入力してください。") from exc

    def read_form(self) -> dict[str, Any]:
        external_port = self.parse_port(self.ext_port.get().strip(), "外部ポート")
        internal_port = self.parse_port(self.int_port.get().strip(), "内部ポート")
        internal_ip = self.validate_ip(self.local_ip.get().strip())
        protocol = self.protocol.get().strip().upper()
        description = self.desc.get().strip() or "UPnP Tool"

        if protocol not in {"TCP", "UDP"}:
            raise ValueError("プロトコルは TCP または UDP を選択してください。")

        return {
            "external_port": external_port,
            "internal_port": internal_port,
            "internal_ip": internal_ip,
            "protocol": protocol,
            "description": description,
        }

    def read_mapping_key(self) -> dict[str, Any]:
        external_port = self.parse_port(self.ext_port.get().strip(), "外部ポート")
        protocol = self.protocol.get().strip().upper()

        if protocol not in {"TCP", "UDP"}:
            raise ValueError("プロトコルは TCP または UDP を選択してください。")

        return {
            "external_port": external_port,
            "protocol": protocol,
        }

    def start_add_mapping(self) -> None:
        if self.service is None:
            self.show_error("未接続", "先に利用するルーターを選択してください。")
            return

        try:
            params = self.read_form()
        except ValueError as exc:
            self.show_error("入力エラー", str(exc))
            return

        self.start_worker(self.add_mapping, params)

    def add_mapping(self, params: dict[str, Any]) -> None:
        if self.service is None:
            return

        external_port = params["external_port"]
        internal_port = params["internal_port"]
        internal_ip = params["internal_ip"]
        protocol = params["protocol"]
        description = params["description"]

        self.log(f"ポート {external_port}/{protocol} を {internal_ip}:{internal_port} に設定します。")

        try:
            existing = self.get_specific_mapping(external_port, protocol)
            if existing:
                current_ip = str(existing.get("NewInternalClient", ""))
                current_port = str(existing.get("NewInternalPort", ""))
                same_target = current_ip == internal_ip and current_port == str(internal_port)

                if same_target:
                    self.log("既存設定が見つかったため、いったん削除してから更新します。")
                    self.service.DeletePortMapping(
                        NewRemoteHost="",
                        NewExternalPort=external_port,
                        NewProtocol=protocol,
                    )
                else:
                    raise RuntimeError(
                        f"外部ポート {external_port}/{protocol} は既に "
                        f"{current_ip}:{current_port} に割り当てられています。"
                    )

            self.service.AddPortMapping(
                NewRemoteHost="",
                NewExternalPort=external_port,
                NewProtocol=protocol,
                NewInternalPort=internal_port,
                NewInternalClient=internal_ip,
                NewEnabled=1,
                NewPortMappingDescription=description,
                NewLeaseDuration=0,
            )
        except Exception as exc:
            self.log(f"ポート開放に失敗しました: {exc}")
            self.show_error("ポート開放に失敗", str(exc))
            return

        self.log(f"ポート {external_port}/{protocol} を開放しました。")
        self.show_info("完了", f"ポート {external_port}/{protocol} を開放しました。")
        self.refresh_mappings()

    def get_specific_mapping(self, external_port: int, protocol: str) -> dict[str, Any] | None:
        try:
            return self.service.GetSpecificPortMappingEntry(
                NewRemoteHost="",
                NewExternalPort=external_port,
                NewProtocol=protocol,
            )
        except Exception:
            return None

    def start_delete_mapping(self) -> None:
        if self.service is None:
            self.show_error("未接続", "先に利用するルーターを選択してください。")
            return

        try:
            params = self.read_mapping_key()
        except ValueError as exc:
            self.show_error("入力エラー", str(exc))
            return

        self.start_worker(self.delete_mapping, params)

    def delete_mapping(self, params: dict[str, Any]) -> None:
        if self.service is None:
            return

        external_port = params["external_port"]
        protocol = params["protocol"]
        self.log(f"ポート {external_port}/{protocol} の削除を開始します。")

        try:
            existing = self.get_specific_mapping(external_port, protocol)
            if not existing:
                raise RuntimeError(f"ポート {external_port}/{protocol} は見つかりませんでした。")

            self.service.DeletePortMapping(
                NewRemoteHost="",
                NewExternalPort=external_port,
                NewProtocol=protocol,
            )
        except Exception as exc:
            self.log(f"ポート削除に失敗しました: {exc}")
            self.show_error("ポート削除に失敗", str(exc))
            return

        self.log(f"ポート {external_port}/{protocol} を削除しました。")
        self.show_info("完了", f"ポート {external_port}/{protocol} を削除しました。")
        self.refresh_mappings()


if __name__ == "__main__":
    root = tk.Tk()
    app = UPnPToolGUI(root)
    root.mainloop()
