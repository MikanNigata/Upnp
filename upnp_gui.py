import ipaddress
import socket
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk
from typing import Any

import upnpy


class UPnPToolGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("UPnP ポート開放ツール")
        self.root.geometry("760x720")
        self.root.minsize(720, 620)

        self.upnp = upnpy.UPnP()
        self.service = None
        self.device = None
        self.service_candidates: list[dict[str, Any]] = []
        self.mapping_rows: dict[str, dict[str, str]] = {}
        self.busy = False

        self.status_var = tk.StringVar(value="UPnP デバイスを検出中...")
        self.external_ip_var = tk.StringVar(value="未取得")
        self.router_var = tk.StringVar()

        self.create_widgets()
        self.auto_fill_ip()
        self.start_discovery()

    def create_widgets(self) -> None:
        main_frame = ttk.Frame(self.root, padding=12)
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)

        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=0, column=0, sticky="ew")
        status_frame.columnconfigure(0, weight=1)

        self.status_label = ttk.Label(
            status_frame,
            textvariable=self.status_var,
            foreground="#1d4ed8",
            font=("", 10, "bold"),
        )
        self.status_label.grid(row=0, column=0, sticky="w")

        self.refresh_devices_btn = ttk.Button(
            status_frame,
            text="ルーター再検索",
            command=self.start_discovery,
        )
        self.refresh_devices_btn.grid(row=0, column=1, sticky="e")

        router_frame = ttk.LabelFrame(main_frame, text="ルーター情報", padding=10)
        router_frame.grid(row=1, column=0, pady=(10, 8), sticky="ew")
        router_frame.columnconfigure(1, weight=1)

        ttk.Label(router_frame, text="対象ルーター").grid(row=0, column=0, sticky="w")
        self.router_combo = ttk.Combobox(
            router_frame,
            textvariable=self.router_var,
            state="disabled",
        )
        self.router_combo.grid(row=0, column=1, padx=6, pady=3, sticky="ew")
        self.router_combo.bind("<<ComboboxSelected>>", self.on_router_selected)

        ttk.Label(router_frame, text="グローバル IP").grid(row=1, column=0, sticky="w")
        self.external_ip_label = ttk.Label(router_frame, textvariable=self.external_ip_var)
        self.external_ip_label.grid(row=1, column=1, padx=6, pady=3, sticky="w")

        config_frame = ttk.LabelFrame(main_frame, text="ポート設定", padding=10)
        config_frame.grid(row=2, column=0, pady=8, sticky="ew")
        config_frame.columnconfigure(1, weight=1)
        config_frame.columnconfigure(3, weight=1)

        ttk.Label(config_frame, text="外部ポート").grid(row=0, column=0, sticky="w")
        self.ext_port = ttk.Entry(config_frame)
        self.ext_port.insert(0, "8080")
        self.ext_port.grid(row=0, column=1, padx=6, pady=3, sticky="ew")

        ttk.Label(config_frame, text="内部ポート").grid(row=0, column=2, sticky="w")
        self.int_port = ttk.Entry(config_frame)
        self.int_port.insert(0, "8080")
        self.int_port.grid(row=0, column=3, padx=6, pady=3, sticky="ew")

        ttk.Label(config_frame, text="プロトコル").grid(row=1, column=0, sticky="w")
        self.protocol = ttk.Combobox(config_frame, values=["TCP", "UDP"], state="readonly")
        self.protocol.current(0)
        self.protocol.grid(row=1, column=1, padx=6, pady=3, sticky="ew")

        ttk.Label(config_frame, text="内部 IP").grid(row=1, column=2, sticky="w")
        self.local_ip = ttk.Entry(config_frame)
        self.local_ip.grid(row=1, column=3, padx=6, pady=3, sticky="ew")

        ttk.Label(config_frame, text="説明").grid(row=2, column=0, sticky="w")
        self.desc = ttk.Entry(config_frame)
        self.desc.insert(0, "UPnP Tool")
        self.desc.grid(row=2, column=1, columnspan=3, padx=6, pady=3, sticky="ew")

        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=3, column=0, pady=8, sticky="ew")

        self.open_btn = ttk.Button(
            action_frame,
            text="ポートを開放 / 更新",
            command=self.start_add_mapping,
            state="disabled",
        )
        self.open_btn.pack(side=tk.LEFT, padx=(0, 6))

        self.close_btn = ttk.Button(
            action_frame,
            text="フォーム内容で削除",
            command=self.start_delete_mapping,
            state="disabled",
        )
        self.close_btn.pack(side=tk.LEFT, padx=6)

        self.refresh_mappings_btn = ttk.Button(
            action_frame,
            text="一覧を更新",
            command=self.start_refresh_mappings,
            state="disabled",
        )
        self.refresh_mappings_btn.pack(side=tk.LEFT, padx=6)

        mapping_frame = ttk.LabelFrame(main_frame, text="現在のポートマッピング", padding=10)
        mapping_frame.grid(row=4, column=0, pady=(8, 8), sticky="nsew")
        mapping_frame.columnconfigure(0, weight=1)
        mapping_frame.rowconfigure(0, weight=1)

        columns = ("ext", "proto", "client", "int", "enabled", "lease", "desc")
        self.mapping_tree = ttk.Treeview(
            mapping_frame,
            columns=columns,
            show="headings",
            height=12,
        )
        self.mapping_tree.heading("ext", text="外部")
        self.mapping_tree.heading("proto", text="プロトコル")
        self.mapping_tree.heading("client", text="内部 IP")
        self.mapping_tree.heading("int", text="内部")
        self.mapping_tree.heading("enabled", text="有効")
        self.mapping_tree.heading("lease", text="リース")
        self.mapping_tree.heading("desc", text="説明")

        self.mapping_tree.column("ext", width=80, anchor=tk.CENTER)
        self.mapping_tree.column("proto", width=90, anchor=tk.CENTER)
        self.mapping_tree.column("client", width=140, anchor=tk.CENTER)
        self.mapping_tree.column("int", width=80, anchor=tk.CENTER)
        self.mapping_tree.column("enabled", width=60, anchor=tk.CENTER)
        self.mapping_tree.column("lease", width=80, anchor=tk.CENTER)
        self.mapping_tree.column("desc", width=220, anchor=tk.W)
        self.mapping_tree.grid(row=0, column=0, sticky="nsew")
        self.mapping_tree.bind("<<TreeviewSelect>>", self.on_mapping_selected)

        tree_scroll = ttk.Scrollbar(mapping_frame, orient=tk.VERTICAL, command=self.mapping_tree.yview)
        tree_scroll.grid(row=0, column=1, sticky="ns")
        self.mapping_tree.configure(yscrollcommand=tree_scroll.set)

        ttk.Label(
            mapping_frame,
            text="一覧を選ぶとフォームに反映されます。削除前の確認にも使えます。",
        ).grid(row=1, column=0, columnspan=2, pady=(8, 0), sticky="w")

        ttk.Label(main_frame, text="ログ").grid(row=5, column=0, sticky="w")
        self.log_area = scrolledtext.ScrolledText(
            main_frame,
            height=10,
            state="disabled",
            font=("Consolas", 9),
        )
        self.log_area.grid(row=6, column=0, sticky="nsew")
        main_frame.rowconfigure(6, weight=1)

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
        def apply() -> None:
            self.status_var.set(text)
            self.status_label.config(foreground=color)

        self.root.after(0, apply)

    def set_external_ip(self, value: str) -> None:
        self.root.after(0, lambda: self.external_ip_var.set(value))

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
        self.set_status("UPnP デバイスを検索中...", "#1d4ed8")
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
        self.set_status(f"接続中: {friendly_name}", "#1d4ed8")
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
