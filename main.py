import re
import socket
import threading
import time
from datetime import datetime, timezone

import requests
from kivy.app import App
from kivy.core.window import Window
from kivy.graphics import Color, RoundedRectangle
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.scrollview import ScrollView
from kivy.clock import Clock

try:
    from jnius import autoclass, cast
    from android import activity as android_activity
    ANDROID_AVAILABLE = True
except ImportError:
    ANDROID_AVAILABLE = False

# App color palette — change these 4 lines to re-theme the whole app
BG_COLOR = (0.09, 0.11, 0.15, 1)
CARD_COLOR = (0.14, 0.16, 0.20, 1)
ACCENT_COLOR = (0.20, 0.55, 0.95, 1)
DANGER_COLOR = (0.80, 0.25, 0.25, 1)
TEXT_COLOR = (0.92, 0.92, 0.94, 1)

Window.clearcolor = BG_COLOR

try:
    import openpyxl
except ImportError:
    openpyxl = None

# ══════════════════════════════════════════════════════════
#  EDIT THESE BEFORE BUILDING
# ══════════════════════════════════════════════════════════
FIREBASE_API_KEY = "AIzaSyATiCDMs5w-RAKZKIB9tIvx27Hb3uruU48"
FIREBASE_PROJECT = "up-data-push"

SERVER_IP, SERVER_PORT = "minesdata.rajasthan.gov.in", 7001
# ══════════════════════════════════════════════════════════

INTERVAL_SECONDS = 10     # how often a full round of packets goes out, same knob as the PC script
SOCKET_TIMEOUT = 15
RECONNECT_DELAY = 5

# The raw packet shape from your PC script (NRM Packet Sender). {} placeholders
# get filled in per device by build_raw_template() below.
RAW_TEMPLATE = (
    "$1,PT4G,WEA1.0,NR,01,L,{imei},{vehicle},1,{date},{time},"
    "{lat},{lat_dir},{lon},{lon_dir},0.0,293.60,0,"
    "0217.19,0.00,0.00,airtel,1,1,12.800,3.700,0,O,23,404,10,0964,"
    "0000000,0|0|00,0011,00,001687,0000,*"
)


# ───────────────────────────────────────────
#  FIREBASE HELPERS (REST API, no SDK needed) — unchanged from your app
# ───────────────────────────────────────────

def firebase_login(email, password):
    url = f"https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key={FIREBASE_API_KEY}"
    r = requests.post(url, json={"email": email, "password": password, "returnSecureToken": True}, timeout=15)
    if r.status_code == 200:
        return r.json()
    return None


def get_tokens(uid, id_token):
    url = f"https://firestore.googleapis.com/v1/projects/{FIREBASE_PROJECT}/databases/(default)/documents/users/{uid}"
    r = requests.get(url, headers={"Authorization": f"Bearer {id_token}"}, timeout=15)
    if r.status_code != 200:
        return 0
    fields = r.json().get("fields", {})
    return int(fields.get("tokens", {}).get("integerValue", 0))


def decrement_token(uid, id_token, current_value):
    new_val = max(0, current_value - 1)
    url = (f"https://firestore.googleapis.com/v1/projects/{FIREBASE_PROJECT}/databases/(default)"
           f"/documents/users/{uid}?updateMask.fieldPaths=tokens")
    requests.patch(url, headers={"Authorization": f"Bearer {id_token}"},
                    json={"fields": {"tokens": {"integerValue": new_val}}}, timeout=15)
    return new_val


# ───────────────────────────────────────────
#  ANDROID FILE PICKER — unchanged from your app
# ───────────────────────────────────────────

PICK_FILE_REQUEST_CODE = 4269


def android_pick_file(callback):
    """callback(local_filepath_or_None, error_or_None) fires once the user picks a file or cancels."""
    if not ANDROID_AVAILABLE:
        callback(None, "Android file access isn't available on this build.")
        return

    PythonActivity = autoclass('org.kivy.android.PythonActivity')
    Intent = autoclass('android.content.Intent')
    Activity = autoclass('android.app.Activity')
    current_activity = cast('android.app.Activity', PythonActivity.mActivity)

    def on_activity_result(request_code, result_code, intent):
        if request_code != PICK_FILE_REQUEST_CODE:
            return
        android_activity.unbind(on_activity_result=on_activity_result)

        if result_code != Activity.RESULT_OK or intent is None:
            callback(None, None)  # user cancelled — not an error
            return

        try:
            uri = intent.getData()
            local_path = _copy_uri_to_local_file(current_activity, uri)
            callback(local_path, None)
        except Exception as e:
            callback(None, f"Could not read the picked file: {e}")

    android_activity.bind(on_activity_result=on_activity_result)

    intent = Intent(Intent.ACTION_OPEN_DOCUMENT)
    intent.addCategory(Intent.CATEGORY_OPENABLE)
    intent.setType("*/*")
    current_activity.startActivityForResult(intent, PICK_FILE_REQUEST_CODE)


def _copy_uri_to_local_file(current_activity, uri):
    OpenableColumns = autoclass('android.provider.OpenableColumns')
    resolver = current_activity.getContentResolver()

    display_name = "picked_file.xlsx"
    try:
        cursor = resolver.query(uri, None, None, None, None)
        if cursor is not None:
            cursor.moveToFirst()
            name_index = cursor.getColumnIndex(OpenableColumns.DISPLAY_NAME)
            if name_index != -1:
                display_name = cursor.getString(name_index) or display_name
            cursor.close()
    except Exception:
        pass

    cache_dir = current_activity.getCacheDir().getAbsolutePath()
    local_path = f"{cache_dir}/{display_name}"

    input_stream = resolver.openInputStream(uri)
    with open(local_path, "wb") as out_file:
        buffer = bytearray(4096)
        while True:
            n = input_stream.read(buffer)
            if n == -1:
                break
            out_file.write(bytes(buffer[:n]))
    input_stream.close()

    return local_path


# ───────────────────────────────────────────
#  EXCEL READER — unchanged, its columns already line up with the new
#  protocol's fields (imei, vehicle, latitude, lat_dir, longitude, lon_dir)
# ───────────────────────────────────────────

def read_devices_from_excel(filepath):
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(h).strip().lower() if h else "" for h in rows[0]]

    def col(*names):
        for n in names:
            if n in headers:
                return headers.index(n)
        return None

    i_imei = col("imei")
    i_veh = col("vehicle_no", "vehicle_number", "vehicle")
    i_lat = col("latitude")
    i_ld = col("lat_dir")
    i_lon = col("longitude")
    i_lod = col("lon_dir")

    if None in (i_imei, i_veh, i_lat, i_ld, i_lon, i_lod):
        raise ValueError(f"Excel missing required columns. Found: {headers}")

    devices = []
    for row in rows[1:]:
        try:
            imei, veh, lat, ld, lon, lod = (
                str(row[i_imei]).strip(), str(row[i_veh]).strip(),
                str(row[i_lat]).strip(), str(row[i_ld]).strip().upper(),
                str(row[i_lon]).strip(), str(row[i_lod]).strip().upper(),
            )
        except (IndexError, TypeError):
            continue
        if any(v in ("", "None", "NAN") for v in [imei, veh, lat, ld, lon, lod]):
            continue
        devices.append({"imei": imei, "vehicle_no": veh, "latitude": lat,
                         "lat_dir": ld, "longitude": lon, "lon_dir": lod})
    return devices


# ───────────────────────────────────────────
#  PROTOCOL — ported from your PC script (NRM Packet Sender), not the old
#  LGN/HEL/PVT/EPB set. This protocol sends ONE packet type repeatedly,
#  bumping the date/time and a frame counter each round, with a CRC16/ARC
#  checksum. See the two fixes called out below.
# ───────────────────────────────────────────

def crc16_arc(data: bytes) -> int:
    crc = 0x0000
    for byte in data:
        crc ^= byte
        for _ in range(8):
            crc = (crc >> 1) ^ 0xA001 if crc & 1 else crc >> 1
    return crc & 0xFFFF


def compute_checksum(body_with_trailing_comma: str) -> str:
    val = crc16_arc(("$" + body_with_trailing_comma).encode("ascii"))
    return f"{val:04X}"


def build_raw_template(imei, vehicle, lat, lat_dir, lon, lon_dir):
    now = datetime.now(timezone.utc)
    return RAW_TEMPLATE.format(
        imei=imei, vehicle=vehicle,
        date=now.strftime("%d%m%Y"), time=now.strftime("%H%M%S"),
        lat=lat, lat_dir=lat_dir, lon=lon, lon_dir=lon_dir,
    )


def parse_packet(raw: str, imei: str) -> dict:
    raw = raw.strip()
    if raw.startswith("$"):
        raw = raw[1:]
    raw = raw.split("*")[0]

    fields = raw.split(",")
    frame_idx = None
    for i in range(len(fields) - 1, max(len(fields) - 10, 0), -1):
        if re.fullmatch(r"\d{4,6}", fields[i]):
            frame_idx = i
            break

    if frame_idx is None:
        raise ValueError(f"[{imei}] Could not locate frame number field.")

    return {"fields": fields, "frame_idx": frame_idx, "date_idx": 9, "time_idx": 10, "imei": imei}


def build_updated_packet(parsed: dict) -> str:
    """Stamps the current UTC time and bumps the frame counter by 1, then
    re-signs the packet with a fresh checksum. Mutates parsed['fields'] in
    place so the NEXT call picks up from where this one left off — the PC
    script instead re-parsed a slice of the outgoing string after every
    send, which is harder to follow and only worked by luck."""
    fields = parsed["fields"]

    now = datetime.now(timezone.utc)
    fields[parsed["date_idx"]] = now.strftime("%d%m%Y")
    fields[parsed["time_idx"]] = now.strftime("%H%M%S")

    old_frame = fields[parsed["frame_idx"]]
    frame_width = len(old_frame)
    # FIX: the PC script wrapped the frame counter at a hardcoded 999999.
    # This template's frame field is "0000" — 4 digits, max value 9999 — so
    # the old wrap point would never trigger and the field would silently
    # grow past its width. Derive the wrap point from the field's own width
    # instead of assuming a fixed size.
    max_frame = 10 ** frame_width - 1
    new_frame_int = int(old_frame) + 1
    if new_frame_int > max_frame:
        new_frame_int = 0
    fields[parsed["frame_idx"]] = str(new_frame_int).zfill(frame_width)

    body = ",".join(fields[:-1]) + ","
    fields[-1] = compute_checksum(body)

    return "$" + ",".join(fields) + "*\r\n"


# ───────────────────────────────────────────
#  TCP SENDING — single socket, all devices multiplexed onto it, same idea
#  as your PC script, but reconnect/backoff now checks a stop_event instead
#  of only reacting to Ctrl+C, so the Stop button can interrupt a retry.
# ───────────────────────────────────────────

class StoppedError(Exception):
    """Raised internally to unwind out of a connect retry once Stop is pressed."""


def make_socket(host, port):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(SOCKET_TIMEOUT)
    s.setsockopt(socket.SOL_SOCKET, socket.SO_KEEPALIVE, 1)
    s.connect((host, port))
    return s


def ensure_connected(sock, host, port, stop_event, log_fn):
    if sock is not None:
        return sock
    while not stop_event.is_set():
        try:
            log_fn(f"Connecting to {host}:{port} …")
            s = make_socket(host, port)
            log_fn("✓ Connected.")
            return s
        except Exception as exc:
            log_fn(f"Connection failed: {exc} — retrying in {RECONNECT_DELAY}s")
            stop_event.wait(RECONNECT_DELAY)
    raise StoppedError()


def send_loop(devices, stop_event, log_fn, token_check_fn, token_charge_fn):
    """Runs until stop_event is set or tokens run out. One 'round' = one
    packet sent for every device in `devices`; a token is charged per round,
    matching how your original app charged per successful device send."""
    sock = None
    try:
        while not stop_event.is_set():
            if not token_check_fn():
                log_fn("🚫 Out of tokens. Stopping.")
                return
            try:
                sock = ensure_connected(sock, SERVER_IP, SERVER_PORT, stop_event, log_fn)
            except StoppedError:
                return

            for d in devices:
                if stop_event.is_set():
                    return
                packet = build_updated_packet(d["parsed"])
                log_fn(f"[{d['imei']}] → {packet.strip()}")
                try:
                    sock.sendall(packet.encode("ascii"))
                except Exception as exc:
                    log_fn(f"[{d['imei']}] socket dropped ({exc}), reconnecting…")
                    sock = None
                    try:
                        sock = ensure_connected(sock, SERVER_IP, SERVER_PORT, stop_event, log_fn)
                    except StoppedError:
                        return
                    sock.sendall(packet.encode("ascii"))

            token_charge_fn()
            stop_event.wait(INTERVAL_SECONDS)
    finally:
        if sock is not None:
            try:
                sock.close()
            except Exception:
                pass


# ───────────────────────────────────────────
#  SMALL UI HELPERS
# ───────────────────────────────────────────

class StyledButton(Button):
    def __init__(self, color=ACCENT_COLOR, **kw):
        super().__init__(**kw)
        self.background_normal = ""
        self.background_down = ""
        self.background_color = (0, 0, 0, 0)
        self.color = (1, 1, 1, 1)
        self.font_size = 16
        self.bold = True
        self._btn_color = color
        with self.canvas.before:
            Color(*color)
            self._rect = RoundedRectangle(pos=self.pos, size=self.size, radius=[10])
        self.bind(pos=self._update_rect, size=self._update_rect)

    def _update_rect(self, *_):
        self._rect.pos = self.pos
        self._rect.size = self.size


def styled_label(**kw):
    kw.setdefault("color", TEXT_COLOR)
    return Label(**kw)


def styled_input(**kw):
    kw.setdefault("multiline", False)
    kw.setdefault("background_color", CARD_COLOR)
    kw.setdefault("foreground_color", TEXT_COLOR)
    kw.setdefault("cursor_color", TEXT_COLOR)
    kw.setdefault("padding", [15, 12, 15, 12])
    kw.setdefault("size_hint_y", None)
    kw.setdefault("height", 44)
    return TextInput(**kw)


# ───────────────────────────────────────────
#  UI SCREENS
# ───────────────────────────────────────────

class LoginScreen(Screen):
    def __init__(self, **kw):
        super().__init__(**kw)
        layout = BoxLayout(orientation="vertical", padding=30, spacing=18)
        layout.add_widget(styled_label(text="AIS-140 Sender", font_size=26, bold=True, size_hint_y=0.25))
        self.email = TextInput(hint_text="Email", multiline=False, size_hint_y=0.09,
                                background_color=CARD_COLOR, foreground_color=TEXT_COLOR,
                                cursor_color=TEXT_COLOR, padding=[15, 15, 15, 15])
        self.password = TextInput(hint_text="Password", password=True, multiline=False, size_hint_y=0.09,
                                   background_color=CARD_COLOR, foreground_color=TEXT_COLOR,
                                   cursor_color=TEXT_COLOR, padding=[15, 15, 15, 15])
        self.status = styled_label(text="", size_hint_y=0.1, color=(0.9, 0.4, 0.4, 1))
        btn = StyledButton(text="Login", size_hint_y=0.12)
        btn.bind(on_press=self.do_login)
        layout.add_widget(self.email)
        layout.add_widget(self.password)
        layout.add_widget(btn)
        layout.add_widget(self.status)
        self.add_widget(layout)

    def do_login(self, *_):
        self.status.text = "Logging in..."
        threading.Thread(target=self._login_thread, daemon=True).start()

    def _login_thread(self):
        result = firebase_login(self.email.text.strip(), self.password.text.strip())
        Clock.schedule_once(lambda dt: self._after_login(result))

    def _after_login(self, result):
        if not result:
            self.status.text = "Login failed — check email/password."
            return
        app = App.get_running_app()
        app.uid = result["localId"]
        app.id_token = result["idToken"]
        app.tokens = get_tokens(app.uid, app.id_token)
        self.status.text = ""
        self.manager.current = "main"


class MainScreen(Screen):
    def __init__(self, **kw):
        super().__init__(**kw)
        self.devices = []            # populated in bulk mode from Excel
        self.selected_path = None
        self.mode = "manual"         # "manual" or "bulk"
        self.stop_event = threading.Event()
        self.sending = False

        root = BoxLayout(orientation="vertical", padding=15, spacing=10)

        self.token_label = styled_label(text="Tokens: -", font_size=16, bold=True, size_hint_y=0.05)
        root.add_widget(self.token_label)

        # Mode toggle
        mode_row = BoxLayout(orientation="horizontal", size_hint_y=0.07, spacing=8)
        self.manual_mode_btn = StyledButton(text="Manual Entry")
        self.bulk_mode_btn = StyledButton(text="Bulk (Excel)", color=CARD_COLOR)
        self.manual_mode_btn.bind(on_press=lambda *_: self.set_mode("manual"))
        self.bulk_mode_btn.bind(on_press=lambda *_: self.set_mode("bulk"))
        mode_row.add_widget(self.manual_mode_btn)
        mode_row.add_widget(self.bulk_mode_btn)
        root.add_widget(mode_row)

        # Container that gets swapped between manual fields and the Excel picker
        self.input_area = BoxLayout(orientation="vertical", size_hint_y=0.38, spacing=8)
        root.add_widget(self.input_area)

        # Start / Stop
        action_row = BoxLayout(orientation="horizontal", size_hint_y=0.09, spacing=8)
        self.start_btn = StyledButton(text="Start Sending")
        self.stop_btn = StyledButton(text="Stop", color=DANGER_COLOR)
        self.stop_btn.disabled = True
        self.start_btn.bind(on_press=self.start_sending)
        self.stop_btn.bind(on_press=self.stop_sending)
        action_row.add_widget(self.start_btn)
        action_row.add_widget(self.stop_btn)
        root.add_widget(action_row)

        scroll = ScrollView(size_hint_y=0.41)
        self.log_label = styled_label(text="", size_hint_y=None, halign="left", valign="top")
        self.log_label.bind(texture_size=lambda *_: setattr(self.log_label, "height", self.log_label.texture_size[1]))
        self.log_label.text_size = (self.log_label.width, None)
        scroll.add_widget(self.log_label)
        root.add_widget(scroll)

        self.add_widget(root)
        self._build_manual_fields()
        self.set_mode("manual")

    # ---- mode switching ----

    def _build_manual_fields(self):
        self.manual_box = BoxLayout(orientation="vertical", spacing=6)
        self.imei_input = styled_input(hint_text="IMEI")
        self.vehicle_input = styled_input(hint_text="Vehicle Number")

        lat_row = BoxLayout(orientation="horizontal", size_hint_y=None, height=44, spacing=6)
        self.lat_input = styled_input(hint_text="Latitude e.g. 026.485982")
        self.latdir_input = styled_input(hint_text="N/S", size_hint_x=0.25)
        lat_row.add_widget(self.lat_input)
        lat_row.add_widget(self.latdir_input)

        lon_row = BoxLayout(orientation="horizontal", size_hint_y=None, height=44, spacing=6)
        self.lon_input = styled_input(hint_text="Longitude e.g. 073.772890")
        self.londir_input = styled_input(hint_text="E/W", size_hint_x=0.25)
        lon_row.add_widget(self.lon_input)
        lon_row.add_widget(self.londir_input)

        for w in (self.imei_input, self.vehicle_input, lat_row, lon_row):
            self.manual_box.add_widget(w)

        self.bulk_box = BoxLayout(orientation="vertical", spacing=8)
        self.file_label = styled_label(text="No file selected", size_hint_y=None, height=30,
                                        color=(0.7, 0.7, 0.75, 1))
        pick_btn = StyledButton(text="Select Excel File", size_hint_y=None, height=44)
        pick_btn.bind(on_press=self.pick_excel)
        self.bulk_box.add_widget(self.file_label)
        self.bulk_box.add_widget(pick_btn)

    def set_mode(self, mode):
        if self.sending:
            self.log("⚠️ Stop the current run before switching modes.")
            return
        self.mode = mode
        self.input_area.clear_widgets()
        if mode == "manual":
            self.input_area.add_widget(self.manual_box)
            self.manual_mode_btn._btn_color = ACCENT_COLOR
            self.bulk_mode_btn._btn_color = CARD_COLOR
        else:
            self.input_area.add_widget(self.bulk_box)
            self.manual_mode_btn._btn_color = CARD_COLOR
            self.bulk_mode_btn._btn_color = ACCENT_COLOR
        # repaint the toggle buttons to reflect which one is active
        for btn in (self.manual_mode_btn, self.bulk_mode_btn):
            btn.canvas.before.clear()
            with btn.canvas.before:
                Color(*btn._btn_color)
                btn._rect = RoundedRectangle(pos=btn.pos, size=btn.size, radius=[10])
            btn.bind(pos=btn._update_rect, size=btn._update_rect)

    def on_pre_enter(self):
        self.token_label.text = f"Tokens remaining: {App.get_running_app().tokens}"

    def log(self, msg):
        def _upd(dt):
            self.log_label.text += msg + "\n"
        Clock.schedule_once(_upd)

    # ---- excel (bulk mode) ----

    def pick_excel(self, *_):
        try:
            android_pick_file(self._on_file_picked)
        except Exception as e:
            self.log(f"❌ Could not open file picker: {e}")

    def _on_file_picked(self, local_path, error):
        def _load(dt):
            if error:
                self.log(f"❌ {error}")
                return
            if not local_path:
                self.log("ℹ️ No file selected.")
                return
            if not local_path.lower().endswith(".xlsx"):
                self.log(f"❌ Please select a .xlsx file (got: {local_path.split('/')[-1]})")
                return
            self.selected_path = local_path
            self.file_label.text = local_path.split("/")[-1]
            try:
                self.devices = read_devices_from_excel(local_path)
                self.log(f"✅ Loaded {len(self.devices)} device(s) from {self.file_label.text}")
            except Exception as e:
                self.log(f"❌ Failed to read Excel: {e}")
        Clock.schedule_once(_load)

    # ---- start / stop ----

    def start_sending(self, *_):
        if self.sending:
            return

        if self.mode == "bulk":
            if not self.devices:
                self.log("❌ Load an Excel file first.")
                return
            raw_devices = self.devices
        else:
            imei = self.imei_input.text.strip()
            vehicle = self.vehicle_input.text.strip()
            lat = self.lat_input.text.strip()
            lon = self.lon_input.text.strip()
            if not all([imei, vehicle, lat, lon]):
                self.log("❌ Fill in IMEI, Vehicle, Latitude and Longitude first.")
                return
            raw_devices = [{
                "imei": imei, "vehicle_no": vehicle,
                "latitude": lat, "lat_dir": (self.latdir_input.text.strip().upper() or "N"),
                "longitude": lon, "lon_dir": (self.londir_input.text.strip().upper() or "E"),
            }]

        prepared = []
        for dv in raw_devices:
            raw = build_raw_template(dv["imei"], dv["vehicle_no"], dv["latitude"],
                                      dv["lat_dir"], dv["longitude"], dv["lon_dir"])
            try:
                prepared.append({"imei": dv["imei"], "parsed": parse_packet(raw, dv["imei"])})
            except Exception as e:
                self.log(f"❌ Skipping {dv['imei']}: {e}")

        if not prepared:
            self.log("❌ No valid devices to send.")
            return

        self.stop_event = threading.Event()
        self.sending = True
        self.start_btn.disabled = True
        self.stop_btn.disabled = False
        self.log(f"\n▶️ Sending for {len(prepared)} device(s) every {INTERVAL_SECONDS}s. Press Stop to end.")
        threading.Thread(target=self._send_thread, args=(prepared,), daemon=True).start()

    def stop_sending(self, *_):
        if not self.sending:
            return
        self.stop_event.set()
        self.log("⏹ Stop requested — finishing current step...")

    def _send_thread(self, prepared):
        app = App.get_running_app()

        def token_check():
            app.tokens = get_tokens(app.uid, app.id_token)
            return app.tokens > 0

        def token_charge():
            app.tokens = decrement_token(app.uid, app.id_token, app.tokens)
            new_total = app.tokens
            Clock.schedule_once(lambda dt: setattr(self.token_label, "text", f"Tokens remaining: {new_total}"))
            self.log(f"   1 token used this round. Remaining: {new_total}")

        send_loop(prepared, self.stop_event, self.log, token_check, token_charge)

        self.log("⏹ Stopped by user." if self.stop_event.is_set() else "✅ Done (out of tokens).")
        self.sending = False
        Clock.schedule_once(lambda dt: self._reset_buttons())

    def _reset_buttons(self):
        self.start_btn.disabled = False
        self.stop_btn.disabled = True


class AIS140App(App):
    uid = None
    id_token = None
    tokens = 0

    def build(self):
        sm = ScreenManager()
        sm.add_widget(LoginScreen(name="login"))
        sm.add_widget(MainScreen(name="main"))
        return sm


if __name__ == "__main__":
    AIS140App().run()
