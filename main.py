import binascii
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
TEXT_COLOR = (0.92, 0.92, 0.94, 1)

Window.clearcolor = BG_COLOR

try:
    import openpyxl
except ImportError:
    openpyxl = None

# ══════════════════════════════════════════════════════════
#  EDIT THESE THREE BEFORE BUILDING
# ══════════════════════════════════════════════════════════
FIREBASE_API_KEY  = "AIzaSyATiCDMs5w-RAKZKIB9tIvx27Hb3uruU48"
FIREBASE_PROJECT  = "up-data-push"

SERVER_IP, SERVER_PORT       = "127.0.0.1", 5001
EMERGENCY_IP, EMERGENCY_PORT = "127.0.0.1", 5000
# ══════════════════════════════════════════════════════════

DELAY_BETWEEN_PACKETS = 1
DELAY_BETWEEN_DEVICES = 2
LINE_ENDING = "\r\n"
BAT_PCT, LOW_BAT, MEM_PCT, RATE_ON, RATE_OFF = 85, 20, 10, 10, 60
REPLY_NUMBER = "0000"

BASE_NRM = (
    "$NRM,WTEX,1.0NTC,NR,01,L,869645080974068,DL01DX3423,1,26072026,080433,"
    "028.442605,N,077.083178,E,000.0,137.99,06,0256.90,2.00,0.44,airtel,1,1,"
    "13.3,4.1,0,C,25,404,10,006a,2aea,9805,006a,18,9804,006a,18,7641,006a,18,"
    "edff,0096,17,0010,00,0.0,0.0,007562,87.026,-,1.0,0,0,"
    "2G_1_97WG_W_1_2_1_27_0_0_0,B44F0E03*"
)

# ───────────────────────────────────────────
#  FIREBASE HELPERS (REST API, no SDK needed)
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
#  ANDROID FILE PICKER (direct Intent + ContentResolver, no plyer)
#  Modern Android hands back a content:// reference, not a plain file
#  path — we copy its bytes into our own app storage so openpyxl can
#  open it like a normal file.
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
    """Copies the bytes behind a content:// URI into our app's private cache
    folder and returns a normal filesystem path openpyxl can open."""
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
        pass  # non-fatal — we'll just use the fallback name above

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
#  EXCEL READER (openpyxl, lighter than pandas for APK builds)
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
    i_veh  = col("vehicle_no", "vehicle_number", "vehicle")
    i_lat  = col("latitude")
    i_ld   = col("lat_dir")
    i_lon  = col("longitude")
    i_lod  = col("lon_dir")

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
#  CHECKSUMS + PARSER + PACKET GENERATORS  (unchanged from original script)
# ───────────────────────────────────────────

def crc32_epb(packet):
    data = packet[:packet.index('*') + 1]
    return f"{binascii.crc32(data.encode('ascii')) & 0xFFFFFFFF:08X}"


def crc16_ibm(data):
    crc = 0x0000
    for b in data.encode('ascii'):
        crc ^= b
        for _ in range(8):
            crc = (crc >> 1) ^ 0xA001 if crc & 1 else crc >> 1
    return f"{crc & 0xFFFF:04X}"


def parse_nrm(nrm):
    nrm = nrm.strip()
    parts = nrm.split(',')
    parts[-1] = parts[-1].split('*')[0].strip()
    return {
        "raw": nrm, "vendor_id": parts[1], "fw_version": parts[2], "packet_type": parts[3],
        "alert_id": parts[4], "pkt_status": parts[5], "imei": parts[6], "vehicle_no": parts[7],
        "gps_fix": parts[8], "date": parts[9], "time": parts[10], "latitude": parts[11],
        "lat_dir": parts[12], "longitude": parts[13], "lon_dir": parts[14], "speed": parts[15],
        "heading": parts[16], "satellites": parts[17], "altitude": parts[18], "pdop": parts[19],
        "hdop": parts[20], "operator": parts[21], "ignition": parts[22], "main_power": parts[23],
        "main_voltage": parts[24], "bat_voltage": parts[25], "emergency": parts[26],
        "tamper": parts[27], "gsm_strength": parts[28], "mcc": parts[29], "mnc": parts[30],
        "lac": parts[31], "cell_id": parts[32], "nmr": ','.join(parts[33:45]),
        "din_status": parts[45], "dout_status": parts[46], "analog1": parts[47],
        "analog2": parts[48], "frame_no": parts[49], "odometer": parts[50],
    }


def gen_lgn(f):
    return f"$LGN,{f['vehicle_no']},{f['imei']},{f['fw_version']},AIS140,{f['latitude']},{f['longitude']}*"


def gen_hel(f):
    din = f['din_status'].zfill(6)
    return (f"$HEL,{f['vendor_id']},{f['fw_version']},{f['imei']},{BAT_PCT},{LOW_BAT},"
            f"{MEM_PCT},{RATE_ON},{RATE_OFF},{din},{f['dout_status']}*")


def gen_pvt(f):
    body = (
        f"$PVT,{f['vendor_id']},{f['fw_version']},{f['packet_type']},{f['alert_id']},{f['pkt_status']},"
        f"{f['imei']},{f['vehicle_no']},{f['gps_fix']},{f['date']},{f['time']},{f['latitude']},"
        f"{f['lat_dir']},{f['longitude']},{f['lon_dir']},{f['speed']},{f['heading']},{f['satellites']},"
        f"{f['altitude']},{f['pdop']},{f['hdop']},{f['operator']},{f['ignition']},{f['main_power']},"
        f"{f['main_voltage']},{f['bat_voltage']},{f['emergency']},{f['tamper']},{f['gsm_strength']},"
        f"{f['mcc']},{f['mnc']},{f['lac']},{f['cell_id']},{f['nmr']},{f['din_status']},"
        f"{f['dout_status']},{f['frame_no']},"
    )
    return body + crc16_ibm(body) + "*"


def gen_epb(f, pkt_type):
    gps_fix = 'A' if f['gps_fix'] == '1' else 'V'
    body = (f"$EPB,{pkt_type},{f['imei']},NM,{f['date']}{f['time']},{gps_fix},{f['latitude']},"
            f"{f['lat_dir']},{f['longitude']},{f['lon_dir']},{f['altitude']},{f['speed']},"
            f"{f['odometer']},G,{f['vehicle_no']},{REPLY_NUMBER}*")
    return body + crc32_epb(body)


def send_packets(packets, ip, port, log):
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(10)
        sock.connect((ip, port))
    except OSError as e:
        log(f"    ❌ Connect failed {ip}:{port} — {e}")
        return False
    try:
        for pkt in packets:
            sock.sendall((pkt["data"] + LINE_ENDING).encode('ascii'))
            log(f"    📤 [{pkt['label']}] {pkt['data']}")
            time.sleep(DELAY_BETWEEN_PACKETS)
    finally:
        sock.close()
    return True


# ───────────────────────────────────────────
#  SMALL UI HELPERS
# ───────────────────────────────────────────

class StyledButton(Button):
    """A Button with rounded corners and our accent color instead of Kivy's default grey box."""
    def __init__(self, **kw):
        super().__init__(**kw)
        self.background_normal = ""
        self.background_down = ""
        self.background_color = (0, 0, 0, 0)  # hide default flat background
        self.color = (1, 1, 1, 1)
        self.font_size = 16
        self.bold = True
        with self.canvas.before:
            Color(*ACCENT_COLOR)
            self._rect = RoundedRectangle(pos=self.pos, size=self.size, radius=[10])
        self.bind(pos=self._update_rect, size=self._update_rect)

    def _update_rect(self, *_):
        self._rect.pos = self.pos
        self._rect.size = self.size


def styled_label(**kw):
    kw.setdefault("color", TEXT_COLOR)
    return Label(**kw)


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
        self.devices = []
        self.selected_path = None
        root = BoxLayout(orientation="vertical", padding=15, spacing=12)

        self.token_label = styled_label(text="Tokens: -", font_size=16, bold=True, size_hint_y=0.06)
        root.add_widget(self.token_label)

        self.file_label = styled_label(text="No file selected", size_hint_y=0.08,
                                        color=(0.7, 0.7, 0.75, 1))
        root.add_widget(self.file_label)

        pick_btn = StyledButton(text="Select Excel File", size_hint_y=0.09)
        pick_btn.bind(on_press=self.pick_excel)
        root.add_widget(pick_btn)

        self.send_btn = StyledButton(text="Send All Devices", size_hint_y=0.09)
        self.send_btn.bind(on_press=self.send_all)
        root.add_widget(self.send_btn)

        scroll = ScrollView(size_hint_y=0.68)
        self.log_label = styled_label(text="", size_hint_y=None, halign="left", valign="top")
        self.log_label.bind(texture_size=lambda *_: setattr(self.log_label, "height", self.log_label.texture_size[1]))
        self.log_label.text_size = (self.log_label.width, None)
        scroll.add_widget(self.log_label)
        root.add_widget(scroll)

        self.add_widget(root)

    def on_pre_enter(self):
        self.token_label.text = f"Tokens remaining: {App.get_running_app().tokens}"

    def log(self, msg):
        def _upd(dt):
            self.log_label.text += msg + "\n"
        Clock.schedule_once(_upd)

    def pick_excel(self, *_):
        """Opens Android's own system file picker instead of drawing our own file browser.
        This needs no storage permission at all — Android hands us the file directly."""
        try:
            android_pick_file(self._on_file_picked)
        except Exception as e:
            self.log(f"❌ Could not open file picker: {e}")

    def _on_file_picked(self, local_path, error):
        # android_pick_file's callback can fire off the UI thread — hop back on it
        # before touching any widgets or doing file I/O that updates the UI.
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

    def send_all(self, *_):
        if not self.devices:
            self.log("❌ Load an Excel file first.")
            return
        threading.Thread(target=self._send_thread, daemon=True).start()

    def _send_thread(self):
        app = App.get_running_app()
        base_f = parse_nrm(BASE_NRM)

        for idx, dev in enumerate(self.devices, 1):
            # Refresh + check token balance BEFORE each send
            app.tokens = get_tokens(app.uid, app.id_token)
            if app.tokens <= 0:
                self.log("🚫 Out of tokens. Contact admin to top up your account.")
                break

            now = datetime.now(timezone.utc)
            f = dict(base_f)
            f.update({
                "imei": dev["imei"], "vehicle_no": dev["vehicle_no"],
                "latitude": dev["latitude"], "lat_dir": dev["lat_dir"],
                "longitude": dev["longitude"], "lon_dir": dev["lon_dir"],
                "date": now.strftime("%d%m%Y"), "time": now.strftime("%H%M%S"),
            })

            self.log(f"\n📱 Device {idx}/{len(self.devices)} — {dev['imei']}")
            ok_main = send_packets(
                [{"label": "LGN", "data": gen_lgn(f)},
                 {"label": "HEL", "data": gen_hel(f)},
                 {"label": "PVT", "data": gen_pvt(f)}],
                SERVER_IP, SERVER_PORT, self.log)
            ok_emr = send_packets(
                [{"label": "EMR", "data": gen_epb(f, "EMR")},
                 {"label": "SEM", "data": gen_epb(f, "SEM")}],
                EMERGENCY_IP, EMERGENCY_PORT, self.log)

            if ok_main and ok_emr:
                app.tokens = decrement_token(app.uid, app.id_token, app.tokens)
                self.log(f"   ✅ 1 token used. Remaining: {app.tokens}")
                Clock.schedule_once(lambda dt: setattr(self.token_label, "text", f"Tokens remaining: {app.tokens}"))
            else:
                self.log("   ❌ Send failed — token not charged.")

            time.sleep(DELAY_BETWEEN_DEVICES)

        self.log("\n✅ Done.")


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
