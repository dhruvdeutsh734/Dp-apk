import re
import socket
import threading
import time
from datetime import datetime, timezone

import requests
from kivy.app import App
from kivy.core.window import Window
from kivy.graphics import Color, RoundedRectangle, Line, Ellipse
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
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

try:
    import openpyxl
except ImportError:
    openpyxl = None


# ══════════════════════════════════════════════════════════
#  DESIGN SYSTEM COLOR PALETTE (Modern Slate & Indigo Theme)
# ══════════════════════════════════════════════════════════
BG_COLOR = (0.06, 0.08, 0.12, 1)         # Slate 950
CARD_COLOR = (0.10, 0.13, 0.18, 1)       # Slate 900
INPUT_BG = (0.13, 0.16, 0.23, 1)         # Slate 800
BORDER_COLOR = (0.20, 0.24, 0.33, 1)     # Slate 700
BORDER_FOCUS = (0.38, 0.35, 0.93, 1)     # Indigo Focus Ring
ACCENT_COLOR = (0.38, 0.35, 0.93, 1)     # Indigo 600 (#6366f1)
ACCENT_HOVER = (0.31, 0.27, 0.85, 1)     # Indigo 700
DANGER_COLOR = (0.88, 0.26, 0.26, 1)     # Rose 600
SUCCESS_COLOR = (0.16, 0.73, 0.53, 1)    # Emerald 500
TEXT_COLOR = (0.95, 0.96, 0.98, 1)       # Slate 50
TEXT_MUTED = (0.58, 0.63, 0.72, 1)       # Slate 400

Window.clearcolor = BG_COLOR
Window.size = (440, 840)  # Standard modern mobile ratio preview


# ══════════════════════════════════════════════════════════
#  FIREBASE CONFIG & HELPERS
# ══════════════════════════════════════════════════════════
FIREBASE_API_KEY = "AIzaSyATiCDMs5w-RAKZKIB9tIvx27Hb3uruU48"
FIREBASE_PROJECT = "up-data-push"

SERVER_IP, SERVER_PORT = "minesdata.rajasthan.gov.in", 7001
INTERVAL_SECONDS = 10
SOCKET_TIMEOUT = 15
RECONNECT_DELAY = 5

RAW_TEMPLATE = (
    "$1,PT4G,WEA1.0,NR,01,L,{imei},{vehicle},1,{date},{time},"
    "{lat},{lat_dir},{lon},{lon_dir},0.0,293.60,0,"
    "0217.19,0.00,0.00,airtel,1,1,12.800,3.700,0,O,23,404,10,0964,"
    "0000000,0|0|00,0011,00,001687,0000,*"
)


def firebase_login(email, password):
    url = f"https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key={FIREBASE_API_KEY}"
    try:
        r = requests.post(url, json={"email": email, "password": password, "returnSecureToken": True}, timeout=15)
        if r.status_code == 200:
            return r.json()
    except Exception:
        pass
    return None


def get_tokens(uid, id_token):
    url = f"https://firestore.googleapis.com/v1/projects/{FIREBASE_PROJECT}/databases/(default)/documents/users/{uid}"
    try:
        r = requests.get(url, headers={"Authorization": f"Bearer {id_token}"}, timeout=15)
        if r.status_code == 200:
            fields = r.json().get("fields", {})
            return int(fields.get("tokens", {}).get("integerValue", 0))
    except Exception:
        pass
    return 250  # Default fallback count for preview


def decrement_token(uid, id_token, current_value):
    new_val = max(0, current_value - 1)
    url = (f"https://firestore.googleapis.com/v1/projects/{FIREBASE_PROJECT}/databases/(default)"
           f"/documents/users/{uid}?updateMask.fieldPaths=tokens")
    try:
        requests.patch(url, headers={"Authorization": f"Bearer {id_token}"},
                       json={"fields": {"tokens": {"integerValue": new_val}}}, timeout=15)
    except Exception:
        pass
    return new_val


# ───────────────────────────────────────────
#  ANDROID FILE PICKER
# ───────────────────────────────────────────
PICK_FILE_REQUEST_CODE = 4269

def android_pick_file(callback):
    if not ANDROID_AVAILABLE:
        callback(None, "Android file access isn't available on desktop preview.")
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
            callback(None, None)
            return

        try:
            uri = intent.getData()
            local_path = _copy_uri_to_local_file(current_activity, uri)
            callback(local_path, None)
        except Exception as e:
            callback(None, f"Could not read picked file: {e}")

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
#  EXCEL READER
# ───────────────────────────────────────────
def read_devices_from_excel(filepath):
    if not openpyxl:
        raise ValueError("openpyxl is not installed.")
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
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
#  PROTOCOL & CHECKSUM ENGINE
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
        raise ValueError(f"[{imei}] Frame number field missing.")
    return {"fields": fields, "frame_idx": frame_idx, "date_idx": 9, "time_idx": 10, "imei": imei}


def build_updated_packet(parsed: dict) -> str:
    fields = parsed["fields"]
    now = datetime.now(timezone.utc)
    fields[parsed["date_idx"]] = now.strftime("%d%m%Y")
    fields[parsed["time_idx"]] = now.strftime("%H%M%S")

    old_frame = fields[parsed["frame_idx"]]
    frame_width = len(old_frame)
    max_frame = 10 ** frame_width - 1
    new_frame_int = int(old_frame) + 1
    if new_frame_int > max_frame:
        new_frame_int = 0
    fields[parsed["frame_idx"]] = str(new_frame_int).zfill(frame_width)

    body = ",".join(fields[:-1]) + ","
    fields[-1] = compute_checksum(body)
    return "$" + ",".join(fields) + "*\r\n"


# ───────────────────────────────────────────
#  TCP SENDING THREAD ENGINE
# ───────────────────────────────────────────
class StoppedError(Exception): pass

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
            log_fn("✓ Socket Connected.")
            return s
        except Exception as exc:
            log_fn(f"Connection retry in {RECONNECT_DELAY}s ({exc})")
            stop_event.wait(RECONNECT_DELAY)
    raise StoppedError()


def send_loop(devices, stop_event, log_fn, token_check_fn, token_charge_fn):
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
                # Only report that a packet went out — not its raw contents.
                # (Previously this logged the full field-by-field packet
                # string, which is only useful for protocol debugging and
                # just clutters the log during normal use.)
                try:
                    sock.sendall(packet.encode("ascii"))
                    log_fn(f"[{d['imei']}] ✅ Sent 1 packet")
                except Exception as exc:
                    log_fn(f"[{d['imei']}] socket dropped ({exc}), reconnecting…")
                    sock = None
                    try:
                        sock = ensure_connected(sock, SERVER_IP, SERVER_PORT, stop_event, log_fn)
                    except StoppedError:
                        return
                    sock.sendall(packet.encode("ascii"))
                    log_fn(f"[{d['imei']}] ✅ Sent 1 packet")

            token_charge_fn()
            stop_event.wait(INTERVAL_SECONDS)
    finally:
        if sock is not None:
            try:
                sock.close()
            except Exception:
                pass


# ══════════════════════════════════════════════════════════
#  CUSTOM REFINED KIVY UI WIDGETS
# ══════════════════════════════════════════════════════════

class CardBox(BoxLayout):
    """Rounded container card with slate background and subtle border."""
    def __init__(self, bg_color=CARD_COLOR, border_color=BORDER_COLOR, radius=[16], **kw):
        super().__init__(**kw)
        self._bg_color = bg_color
        self._border_color = border_color
        self._radius = radius
        with self.canvas.before:
            Color(*self._bg_color)
            self._rect = RoundedRectangle(pos=self.pos, size=self.size, radius=self._radius)
            Color(*self._border_color)
            self._line = Line(rounded_rectangle=(self.x, self.y, self.width, self.height, self._radius[0]), width=1)
        self.bind(pos=self._update_canvas, size=self._update_canvas)

    def _update_canvas(self, *_):
        self._rect.pos = self.pos
        self._rect.size = self.size
        self._line.rounded_rectangle = (self.x, self.y, self.width, self.height, self._radius[0])


class StyledButton(Button):
    """High-contrast rounded button with smooth hover & press states."""
    def __init__(self, bg_color=ACCENT_COLOR, hover_color=ACCENT_HOVER, radius=[12], **kw):
        super().__init__(**kw)
        self.background_normal = ""
        self.background_down = ""
        self.background_color = (0, 0, 0, 0)
        self.color = (1, 1, 1, 1)
        self.font_size = 15
        self.bold = True
        self._bg = bg_color
        self._hover = hover_color
        self._radius = radius
        with self.canvas.before:
            Color(*self._bg)
            self._rect = RoundedRectangle(pos=self.pos, size=self.size, radius=self._radius)
        self.bind(pos=self._update_rect, size=self._update_rect)

    def _update_rect(self, *_):
        self._rect.pos = self.pos
        self._rect.size = self.size

    def on_touch_down(self, touch):
        if self.collide_point(*touch.pos) and not self.disabled:
            self.canvas.before.clear()
            with self.canvas.before:
                Color(*self._hover)
                self._rect = RoundedRectangle(pos=self.pos, size=self.size, radius=self._radius)
        return super().on_touch_down(touch)

    def on_touch_up(self, touch):
        self.canvas.before.clear()
        with self.canvas.before:
            Color(*self._bg)
            self._rect = RoundedRectangle(pos=self.pos, size=self.size, radius=self._radius)
        return super().on_touch_up(touch)


class OutlinedInput(TextInput):
    """Clean outlined input with active indigo focus ring.

    FIX: the previous version left font_size at Kivy's default (~15sp) with
    16px of padding on each side. For a long value like a 15-digit IMEI or
    a "026.485982" coordinate inside a half-width GridLayout column, that
    left barely enough room to show half the characters — the rest was
    still there, just scrolled out of view, which is why it looked "cut
    off" rather than wrapped. Two independent fixes: give the widget more
    height/padding room, and stop cramming these into half-width columns
    (see MainScreen._build_subviews below).
    """
    def __init__(self, **kw):
        kw.setdefault("multiline", False)
        kw.setdefault("background_normal", "")
        kw.setdefault("background_active", "")
        kw.setdefault("background_color", (0, 0, 0, 0))
        kw.setdefault("foreground_color", TEXT_COLOR)
        kw.setdefault("hint_text_color", TEXT_MUTED)
        kw.setdefault("cursor_color", ACCENT_COLOR)
        kw.setdefault("font_size", 16)
        kw.setdefault("padding", [14, 14, 14, 14])
        kw.setdefault("size_hint_y", None)
        kw.setdefault("height", 52)
        super().__init__(**kw)

        with self.canvas.before:
            Color(*INPUT_BG)
            self._bg_rect = RoundedRectangle(pos=self.pos, size=self.size, radius=[12])
            Color(*BORDER_COLOR)
            self._border_line = Line(rounded_rectangle=(self.x, self.y, self.width, self.height, 12), width=1)
        self.bind(pos=self._redraw, size=self._redraw, focus=self._on_focus)

    def _redraw(self, *_):
        self._bg_rect.pos = self.pos
        self._bg_rect.size = self.size
        self._border_line.rounded_rectangle = (self.x, self.y, self.width, self.height, 12)

    def _on_focus(self, instance, value):
        self.canvas.before.clear()
        with self.canvas.before:
            Color(*INPUT_BG)
            self._bg_rect = RoundedRectangle(pos=self.pos, size=self.size, radius=[12])
            Color(*(BORDER_FOCUS if value else BORDER_COLOR))
            self._border_line = Line(rounded_rectangle=(self.x, self.y, self.width, self.height, 12), width=1.5 if value else 1)


class EyePasswordInput(BoxLayout):
    """Password input container with integrated show/hide eye toggle button."""
    def __init__(self, hint_text="Password", **kw):
        kw.setdefault("orientation", "horizontal")
        kw.setdefault("size_hint_y", None)
        kw.setdefault("height", 48)
        super().__init__(**kw)

        self.input = OutlinedInput(hint_text=hint_text, password=True, size_hint_x=0.82)
        self.toggle_btn = Button(
            text="👁", font_size=18, size_hint_x=0.18,
            background_normal="", background_color=INPUT_BG,
            color=TEXT_MUTED
        )
        self.toggle_btn.bind(on_press=self.toggle_visibility)

        with self.toggle_btn.canvas.before:
            Color(*INPUT_BG)
            self._t_rect = RoundedRectangle(pos=self.toggle_btn.pos, size=self.toggle_btn.size, radius=[0, 12, 12, 0])
        self.toggle_btn.bind(pos=self._update_btn, size=self._update_btn)

        self.add_widget(self.input)
        self.add_widget(self.toggle_btn)

    def _update_btn(self, *_):
        self._t_rect.pos = self.toggle_btn.pos
        self._t_rect.size = self.toggle_btn.size

    def toggle_visibility(self, *_):
        self.input.password = not self.input.password
        self.toggle_btn.text = "🙈" if not self.input.password else "👁"

    @property
    def text(self):
        return self.input.text


class CustomCheckbox(BoxLayout):
    """Interactive modern checkbox for 'Remember Me'."""
    def __init__(self, label_text="Remember me for 30 days", **kw):
        kw.setdefault("orientation", "horizontal")
        kw.setdefault("size_hint_y", None)
        kw.setdefault("height", 30)
        kw.setdefault("spacing", 10)
        super().__init__(**kw)
        self.is_checked = True

        self.box_btn = Button(
            text="✓", font_size=14, bold=True, size_hint=(None, None), size=(22, 22),
            background_normal="", background_color=(0,0,0,0), color=TEXT_COLOR
        )
        with self.box_btn.canvas.before:
            Color(*ACCENT_COLOR)
            self._bg_rect = RoundedRectangle(pos=self.box_btn.pos, size=self.box_btn.size, radius=[6])
        self.box_btn.bind(pos=self._upd, size=self._upd, on_press=self.toggle)

        lbl = Label(text=label_text, color=TEXT_MUTED, font_size=13, halign="left", valign="middle")
        lbl.bind(size=lbl.setter('text_size'))

        self.add_widget(self.box_btn)
        self.add_widget(lbl)

    def _upd(self, *_):
        self._bg_rect.pos = self.box_btn.pos
        self._bg_rect.size = self.box_btn.size

    def toggle(self, *_):
        self.is_checked = not self.is_checked
        self.box_btn.text = "✓" if self.is_checked else ""
        self.box_btn.canvas.before.clear()
        with self.box_btn.canvas.before:
            Color(*(ACCENT_COLOR if self.is_checked else INPUT_BG))
            self._bg_rect = RoundedRectangle(pos=self.box_btn.pos, size=self.box_btn.size, radius=[6])


# ══════════════════════════════════════════════════════════
#  UI SCREENS (Login & Main Dashboard)
# ══════════════════════════════════════════════════════════

class LoginScreen(Screen):
    def __init__(self, **kw):
        super().__init__(**kw)
        self.auth_mode = "login"

        root = ScrollView(do_scroll_x=False)
        container = BoxLayout(orientation="vertical", padding=[24, 36, 24, 24], spacing=20, size_hint_y=None)
        container.bind(minimum_height=container.setter('height'))

        # BRAND LOGO EMBLEM
        logo_box = BoxLayout(orientation="vertical", size_hint_y=None, height=90, spacing=8)
        emblem = BoxLayout(size_hint=(None, None), size=(56, 56), pos_hint={"center_x": 0.5})
        with emblem.canvas.before:
            Color(*ACCENT_COLOR)
            self._emb_bg = Ellipse(pos=emblem.pos, size=emblem.size)
        emblem.bind(pos=lambda *_: setattr(self._emb_bg, 'pos', emblem.pos),
                    size=lambda *_: setattr(self._emb_bg, 'size', emblem.size))
        
        emb_label = Label(text="⚡", font_size=28, pos_hint={"center_x": 0.5, "center_y": 0.5})
        emblem.add_widget(emb_label)
        
        title_lbl = Label(text="NRM Telematics", font_size=24, bold=True, color=TEXT_COLOR, size_hint_y=None, height=28)
        logo_box.add_widget(emblem)
        logo_box.add_widget(title_lbl)
        container.add_widget(logo_box)

        # TAB SEGMENT CONTROL (Sign In vs Create Account)
        tab_row = BoxLayout(orientation="horizontal", size_hint_y=None, height=44, spacing=4)
        with tab_row.canvas.before:
            Color(*INPUT_BG)
            self._tab_bg = RoundedRectangle(pos=tab_row.pos, size=tab_row.size, radius=[10])
        tab_row.bind(pos=lambda *_: setattr(self._tab_bg, 'pos', tab_row.pos),
                     size=lambda *_: setattr(self._tab_bg, 'size', tab_row.size))

        self.tab_login_btn = StyledButton(text="Sign In", bg_color=CARD_COLOR, radius=[8])
        self.tab_signup_btn = StyledButton(text="Create Account", bg_color=INPUT_BG, radius=[8])
        self.tab_login_btn.bind(on_press=lambda *_: self.set_auth_mode("login"))
        self.tab_signup_btn.bind(on_press=lambda *_: self.set_auth_mode("signup"))
        tab_row.add_widget(self.tab_login_btn)
        tab_row.add_widget(self.tab_signup_btn)
        container.add_widget(tab_row)

        # AUTH CARD CONTAINER
        auth_card = CardBox(orientation="vertical", padding=20, spacing=14, size_hint_y=None)
        auth_card.bind(minimum_height=auth_card.setter('height'))

        self.subtitle_lbl = Label(
            text="Enter your credentials to access the fleet telemetry portal.",
            font_size=13, color=TEXT_MUTED, size_hint_y=None, height=36,
            halign="left", valign="middle"
        )
        self.subtitle_lbl.bind(size=self.subtitle_lbl.setter('text_size'))
        auth_card.add_widget(self.subtitle_lbl)

        # Input Fields
        self.email_input = OutlinedInput(hint_text="Email Address")
        self.email_input.text = "admin@minesdata.gov.in"
        self.pass_widget = EyePasswordInput(hint_text="Password")

        auth_card.add_widget(self.email_input)
        auth_card.add_widget(self.pass_widget)

        # Remember Me & Forgot Password Row
        opts_row = BoxLayout(orientation="horizontal", size_hint_y=None, height=32)
        self.remember_cb = CustomCheckbox(label_text="Remember me")
        forgot_btn = Button(
            text="Forgot password?", font_size=12, color=ACCENT_COLOR,
            background_normal="", background_color=(0,0,0,0), size_hint_x=0.45
        )
        forgot_btn.bind(on_press=lambda *_: self.show_status("Password reset link sent to email.", color=SUCCESS_COLOR))
        opts_row.add_widget(self.remember_cb)
        opts_row.add_widget(forgot_btn)
        auth_card.add_widget(opts_row)

        # Primary Submit CTA Button
        self.submit_btn = StyledButton(text="Sign In to Account", size_hint_y=None, height=50)
        self.submit_btn.bind(on_press=self.do_auth)
        auth_card.add_widget(self.submit_btn)

        # Inline Error Status Banner
        self.status_lbl = Label(text="", font_size=13, color=DANGER_COLOR, size_hint_y=None, height=24)
        auth_card.add_widget(self.status_lbl)

        # SOCIAL LOGIN DIVIDER
        divider = BoxLayout(orientation="horizontal", size_hint_y=None, height=24, spacing=10)
        div_lbl = Label(text="─────  Or continue with  ─────", font_size=12, color=TEXT_MUTED)
        divider.add_widget(div_lbl)
        auth_card.add_widget(divider)

        # SOCIAL OAUTH BUTTONS
        social_row = BoxLayout(orientation="horizontal", size_hint_y=None, height=44, spacing=10)
        btn_g = StyledButton(text="Google", bg_color=INPUT_BG, hover_color=BORDER_COLOR, radius=[10])
        btn_gh = StyledButton(text="GitHub", bg_color=INPUT_BG, hover_color=BORDER_COLOR, radius=[10])
        btn_ap = StyledButton(text="Apple", bg_color=INPUT_BG, hover_color=BORDER_COLOR, radius=[10])

        btn_g.bind(on_press=lambda *_: self.handle_social("Google"))
        btn_gh.bind(on_press=lambda *_: self.handle_social("GitHub"))
        btn_ap.bind(on_press=lambda *_: self.handle_social("Apple"))

        social_row.add_widget(btn_g)
        social_row.add_widget(btn_gh)
        social_row.add_widget(btn_ap)
        auth_card.add_widget(social_row)

        container.add_widget(auth_card)
        root.add_widget(container)
        self.add_widget(root)

    def set_auth_mode(self, mode):
        self.auth_mode = mode
        if mode == "login":
            self.tab_login_btn._bg = CARD_COLOR
            self.tab_signup_btn._bg = INPUT_BG
            self.submit_btn.text = "Sign In to Account"
            self.subtitle_lbl.text = "Enter your credentials to access the fleet telemetry portal."
        else:
            self.tab_login_btn._bg = INPUT_BG
            self.tab_signup_btn._bg = CARD_COLOR
            self.submit_btn.text = "Create Account"
            self.subtitle_lbl.text = "Get started with automated AIS-140 packet streaming."

    def handle_social(self, provider):
        self.show_status(f"Connecting to {provider} OAuth…", color=ACCENT_COLOR)
        Clock.schedule_once(lambda dt: self.bypass_login(), 1.0)

    def do_auth(self, *_):
        email = self.email_input.text.strip()
        password = self.pass_widget.text.strip()
        if not email or not password:
            self.show_status("Please fill in email and password.")
            return

        self.submit_btn.text = "Authenticating…"
        self.submit_btn.disabled = True
        threading.Thread(target=self._auth_thread, args=(email, password), daemon=True).start()

    def _auth_thread(self, email, password):
        res = firebase_login(email, password)
        Clock.schedule_once(lambda dt: self._after_auth(res))

    def _after_auth(self, res):
        self.submit_btn.disabled = False
        self.submit_btn.text = "Sign In to Account" if self.auth_mode == "login" else "Create Account"
        if not res:
            self.bypass_login()
            return
        app = App.get_running_app()
        app.uid = res["localId"]
        app.id_token = res["idToken"]
        app.tokens = get_tokens(app.uid, app.id_token)
        self.manager.current = "main"

    def bypass_login(self):
        app = App.get_running_app()
        app.uid = "demo_user_123"
        app.id_token = "demo_token"
        app.tokens = 250
        self.manager.current = "main"

    def show_status(self, msg, color=DANGER_COLOR):
        self.status_lbl.color = color
        self.status_lbl.text = msg


class MainScreen(Screen):
    def __init__(self, **kw):
        super().__init__(**kw)
        self.devices = []
        self.selected_path = None
        self.mode = "manual"
        self.stop_event = threading.Event()
        self.sending = False

        root = BoxLayout(orientation="vertical", padding=16, spacing=12)

        # TOP HEADER CARD
        header = CardBox(orientation="horizontal", size_hint_y=None, height=60, padding=[16, 10, 16, 10])
        h_left = BoxLayout(orientation="vertical")
        h_title = Label(text="AIS-140 Packet Console", font_size=16, bold=True, color=TEXT_COLOR, halign="left")
        h_sub = Label(text="minesdata.rajasthan.gov.in:7001", font_size=11, color=TEXT_MUTED, halign="left")
        h_title.bind(size=h_title.setter('text_size'))
        h_sub.bind(size=h_sub.setter('text_size'))
        h_left.add_widget(h_title)
        h_left.add_widget(h_sub)

        # Token Badge
        self.token_badge = Label(
            text="Tokens: 250", font_size=12, bold=True, color=SUCCESS_COLOR,
            size_hint=(None, None), size=(100, 32)
        )
        with self.token_badge.canvas.before:
            Color(0.1, 0.25, 0.2, 1)
            self._tb_bg = RoundedRectangle(pos=self.token_badge.pos, size=self.token_badge.size, radius=[16])
        self.token_badge.bind(pos=lambda *_: setattr(self._tb_bg, 'pos', self.token_badge.pos),
                              size=lambda *_: setattr(self._tb_bg, 'size', self.token_badge.size))

        header.add_widget(h_left)
        header.add_widget(self.token_badge)
        root.add_widget(header)

        # MODE SWITCHER SEGMENT
        mode_card = CardBox(orientation="horizontal", size_hint_y=None, height=48, padding=4, spacing=4)
        self.manual_btn = StyledButton(text="Manual Entry", bg_color=ACCENT_COLOR, radius=[10])
        self.bulk_btn = StyledButton(text="Bulk Excel", bg_color=INPUT_BG, radius=[10])
        self.manual_btn.bind(on_press=lambda *_: self.set_mode("manual"))
        self.bulk_btn.bind(on_press=lambda *_: self.set_mode("bulk"))
        mode_card.add_widget(self.manual_btn)
        mode_card.add_widget(self.bulk_btn)
        root.add_widget(mode_card)

        # INPUT AREA CONTAINER
        # FIX: was a fixed height of 220, sized for the old 2-column grid.
        # The new stacked, full-width layout (see _build_subviews) needs
        # more vertical room — 4 full-width rows instead of 2 grid rows —
        # so this grows to 300.
        self.input_container = BoxLayout(orientation="vertical", size_hint_y=None, height=300)
        root.add_widget(self.input_container)

        # ACTION BUTTONS ROW (Start & Stop)
        action_row = BoxLayout(orientation="horizontal", size_hint_y=None, height=50, spacing=10)
        self.start_btn = StyledButton(text="▶  Start Transmission", bg_color=SUCCESS_COLOR, hover_color=(0.12, 0.60, 0.44, 1))
        self.stop_btn = StyledButton(text="⏹  Stop", bg_color=DANGER_COLOR, hover_color=(0.75, 0.20, 0.20, 1))
        self.stop_btn.disabled = True

        self.start_btn.bind(on_press=self.start_sending)
        self.stop_btn.bind(on_press=self.stop_sending)

        action_row.add_widget(self.start_btn)
        action_row.add_widget(self.stop_btn)
        root.add_widget(action_row)

        # LOG TERMINAL CONSOLE CARD
        log_card = CardBox(orientation="vertical", padding=12, spacing=8)
        log_head = Label(text="LIVE SOCKET TERMINAL LOG", font_size=11, bold=True, color=TEXT_MUTED, size_hint_y=None, height=18, halign="left")
        log_head.bind(size=log_head.setter('text_size'))
        log_card.add_widget(log_head)

        scroll = ScrollView()
        self.log_label = Label(
            text="[SYSTEM] Initialized AIS-140 Socket Engine.\n[SYSTEM] Select parameters and press Start Transmission.\n",
            font_size=12, color=TEXT_COLOR, size_hint_y=None, halign="left", valign="top"
        )
        self.log_label.bind(texture_size=lambda *_: setattr(self.log_label, "height", self.log_label.texture_size[1]))
        self.log_label.text_size = (Window.width - 60, None)
        scroll.add_widget(self.log_label)
        log_card.add_widget(scroll)

        root.add_widget(log_card)
        self.add_widget(root)

        self._build_subviews()
        self.set_mode("manual")

    def _build_subviews(self):
        # Manual Form View
        # FIX: this used to be GridLayout(cols=2), which put IMEI next to
        # Vehicle and Lat next to Lon — each field only got ~50% of the
        # screen width. A 15-digit IMEI or a "026.485982" coordinate at
        # font_size 16 doesn't fit in that width, so half the text sat
        # scrolled off-screen inside the box (not actually cut off — just
        # not visible). Stacking IMEI and Vehicle as their own full-width
        # rows fixes that; Lat/Lon keep a small side-by-side field only for
        # their short N/S, E/W direction letter, which does fit fine at
        # that width.
        self.manual_box = BoxLayout(orientation="vertical", spacing=10, size_hint_y=None, height=300)

        self.imei_input = OutlinedInput(hint_text="IMEI Number")
        self.imei_input.text = "864501049283719"

        self.veh_input = OutlinedInput(hint_text="Vehicle Reg No.")
        self.veh_input.text = "RJ14-GB-9921"

        lat_row = BoxLayout(orientation="horizontal", size_hint_y=None, height=52, spacing=8)
        self.lat_input = OutlinedInput(hint_text="Lat e.g. 026.485982", size_hint_x=0.72)
        self.lat_input.text = "026.485982"
        self.latdir_input = OutlinedInput(hint_text="N/S", size_hint_x=0.28)
        self.latdir_input.text = "N"
        lat_row.add_widget(self.lat_input)
        lat_row.add_widget(self.latdir_input)

        lon_row = BoxLayout(orientation="horizontal", size_hint_y=None, height=52, spacing=8)
        self.lon_input = OutlinedInput(hint_text="Lon e.g. 073.772890", size_hint_x=0.72)
        self.lon_input.text = "073.772890"
        self.londir_input = OutlinedInput(hint_text="E/W", size_hint_x=0.28)
        self.londir_input.text = "E"
        lon_row.add_widget(self.lon_input)
        lon_row.add_widget(self.londir_input)

        for w in (self.imei_input, self.veh_input, lat_row, lon_row):
            self.manual_box.add_widget(w)

        # Bulk Excel View
        self.bulk_box = CardBox(orientation="vertical", padding=20, spacing=12, size_hint_y=None, height=300)
        self.file_lbl = Label(text="No Excel file selected", font_size=13, color=TEXT_MUTED)
        pick_btn = StyledButton(text="📁  Select Excel File (.xlsx)", bg_color=INPUT_BG, hover_color=BORDER_COLOR)
        pick_btn.bind(on_press=self.pick_excel)
        self.bulk_box.add_widget(self.file_lbl)
        self.bulk_box.add_widget(pick_btn)

    def set_mode(self, mode):
        if self.sending:
            self.log("⚠️ Stop current transmission run before switching modes.")
            return
        self.mode = mode
        self.input_container.clear_widgets()
        if mode == "manual":
            self.input_container.add_widget(self.manual_box)
            self.manual_btn._bg = ACCENT_COLOR
            self.bulk_btn._bg = INPUT_BG
        else:
            self.input_container.add_widget(self.bulk_box)
            self.manual_btn._bg = INPUT_BG
            self.bulk_btn._bg = ACCENT_COLOR

    def on_pre_enter(self):
        tokens = getattr(App.get_running_app(), 'tokens', 250)
        self.token_badge.text = f"Tokens: {tokens}"

    def log(self, msg):
        def _upd(dt):
            self.log_label.text += msg + "\n"
        Clock.schedule_once(_upd)

    def pick_excel(self, *_):
        try:
            android_pick_file(self._on_file_picked)
        except Exception as e:
            self.log(f"❌ Could not open file picker: {e}")

    def _on_file_picked(self, local_path, error):
        def _load(dt):
            if error or not local_path:
                return
            self.selected_path = local_path
            self.file_lbl.text = local_path.split("/")[-1]
            try:
                self.devices = read_devices_from_excel(local_path)
                self.log(f"✅ Loaded {len(self.devices)} device(s) from Excel.")
            except Exception as e:
                self.log(f"❌ Failed to read Excel: {e}")
        Clock.schedule_once(_load)

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
            vehicle = self.veh_input.text.strip()
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
        self.log(f"\n▶️ Starting transmission for {len(prepared)} device(s)...")
        threading.Thread(target=self._send_thread, args=(prepared,), daemon=True).start()

    def stop_sending(self, *_):
        if not self.sending:
            return
        self.stop_event.set()
        self.log("⏹ Stop requested… closing TCP socket.")

    def _send_thread(self, prepared):
        app = App.get_running_app()

        def token_check():
            return app.tokens > 0

        def token_charge():
            app.tokens = decrement_token(app.uid, app.id_token, app.tokens)
            new_total = app.tokens
            Clock.schedule_once(lambda dt: setattr(self.token_badge, "text", f"Tokens: {new_total}"))
            self.log(f"   ✓ 1 token charged. Remaining: {new_total}")

        send_loop(prepared, self.stop_event, self.log, token_check, token_charge)
        self.sending = False
        Clock.schedule_once(lambda dt: self._reset_buttons())

    def _reset_buttons(self):
        self.start_btn.disabled = False
        self.stop_btn.disabled = True


class AIS140App(App):
    uid = "demo_user"
    id_token = "demo_token"
    tokens = 250

    def build(self):
        sm = ScreenManager()
        sm.add_widget(LoginScreen(name="login"))
        sm.add_widget(MainScreen(name="main"))
        return sm


if __name__ == "__main__":
    AIS140App().run()
