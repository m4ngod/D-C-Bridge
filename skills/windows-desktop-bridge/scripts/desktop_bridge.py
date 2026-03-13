import base64
import ctypes
import io
import json
import os
import subprocess
import time
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

SW_RESTORE = 9
VK_SHIFT = 0x10
VK_CONTROL = 0x11
VK_MENU = 0x12
VK_RETURN = 0x0D
VK_TAB = 0x09
VK_ESCAPE = 0x1B
VK_SPACE = 0x20
VK_LEFT = 0x25
VK_UP = 0x26
VK_RIGHT = 0x27
VK_DOWN = 0x28

KEY_MAP = {
    'enter': VK_RETURN,
    'tab': VK_TAB,
    'esc': VK_ESCAPE,
    'escape': VK_ESCAPE,
    'space': VK_SPACE,
    'left': VK_LEFT,
    'up': VK_UP,
    'right': VK_RIGHT,
    'down': VK_DOWN,
    'ctrl': VK_CONTROL,
    'shift': VK_SHIFT,
    'alt': VK_MENU,
}

class RECT(ctypes.Structure):
    _fields_ = [('left', ctypes.c_long), ('top', ctypes.c_long), ('right', ctypes.c_long), ('bottom', ctypes.c_long)]

class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ('biSize', ctypes.c_uint32), ('biWidth', ctypes.c_long), ('biHeight', ctypes.c_long),
        ('biPlanes', ctypes.c_ushort), ('biBitCount', ctypes.c_ushort), ('biCompression', ctypes.c_uint32),
        ('biSizeImage', ctypes.c_uint32), ('biXPelsPerMeter', ctypes.c_long), ('biYPelsPerMeter', ctypes.c_long),
        ('biClrUsed', ctypes.c_uint32), ('biClrImportant', ctypes.c_uint32)
    ]

class BITMAPINFO(ctypes.Structure):
    _fields_ = [('bmiHeader', BITMAPINFOHEADER), ('bmiColors', ctypes.c_uint32 * 3)]

EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

gdi32 = ctypes.windll.gdi32


def get_window_text(hwnd):
    length = user32.GetWindowTextLengthW(hwnd)
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)
    return buf.value


def is_window_visible(hwnd):
    return bool(user32.IsWindowVisible(hwnd))


def enum_windows():
    result = []
    @EnumWindowsProc
    def callback(hwnd, lparam):
        if is_window_visible(hwnd):
            title = get_window_text(hwnd)
            if title:
                pid = ctypes.c_ulong()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                rect = RECT()
                user32.GetWindowRect(hwnd, ctypes.byref(rect))
                result.append({
                    'hwnd': int(hwnd),
                    'title': title,
                    'pid': int(pid.value),
                    'rect': {'left': rect.left, 'top': rect.top, 'right': rect.right, 'bottom': rect.bottom}
                })
        return True
    user32.EnumWindows(callback, 0)
    return result


def find_window(title_substring):
    needle = (title_substring or '').lower()
    for w in enum_windows():
        if needle in w['title'].lower():
            return w
    return None


def activate_window(hwnd):
    user32.ShowWindow(hwnd, SW_RESTORE)
    return bool(user32.SetForegroundWindow(hwnd))


def launch(path, args=None):
    cmd = [path] + (args or [])
    p = subprocess.Popen(cmd)
    return {'pid': p.pid, 'path': path, 'args': args or []}


def keybd_event(vk, flags=0):
    user32.keybd_event(vk, 0, flags, 0)


def send_hotkey(keys):
    keycodes = [KEY_MAP[k.lower()] for k in keys]
    for code in keycodes:
        keybd_event(code, 0)
        time.sleep(0.03)
    for code in reversed(keycodes):
        keybd_event(code, 2)
        time.sleep(0.03)
    return {'sent': keys}


def send_text(text):
    for ch in text:
        user32.keybd_event(ord(ch.upper()) if len(ch) == 1 else 0, 0, 0, 0)
        user32.keybd_event(ord(ch.upper()) if len(ch) == 1 else 0, 0, 2, 0)
        time.sleep(0.01)
    return {'typed': text}


def click(x, y):
    user32.SetCursorPos(int(x), int(y))
    user32.mouse_event(0x0002, 0, 0, 0, 0)
    user32.mouse_event(0x0004, 0, 0, 0, 0)
    return {'clicked': {'x': x, 'y': y}}


def screenshot_b64():
    width = user32.GetSystemMetrics(0)
    height = user32.GetSystemMetrics(1)
    hdc_screen = user32.GetDC(0)
    hdc_mem = gdi32.CreateCompatibleDC(hdc_screen)
    hbmp = gdi32.CreateCompatibleBitmap(hdc_screen, width, height)
    gdi32.SelectObject(hdc_mem, hbmp)
    SRCCOPY = 0x00CC0020
    gdi32.BitBlt(hdc_mem, 0, 0, width, height, hdc_screen, 0, 0, SRCCOPY)

    bmi = BITMAPINFO()
    bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bmi.bmiHeader.biWidth = width
    bmi.bmiHeader.biHeight = -height
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 32
    bmi.bmiHeader.biCompression = 0

    buf_len = width * height * 4
    buffer = ctypes.create_string_buffer(buf_len)
    bits = gdi32.GetDIBits(hdc_mem, hbmp, 0, height, buffer, ctypes.byref(bmi), 0)
    if bits != height:
        raise RuntimeError('GetDIBits failed')

    import zlib, struct
    raw = buffer.raw
    stride = width * 4
    scanlines = b''.join(b'\x00' + raw[y*stride:(y+1)*stride] for y in range(height))
    def chunk(tag, data):
        return struct.pack('!I', len(data)) + tag + data + struct.pack('!I', zlib.crc32(tag + data) & 0xffffffff)
    png = b'\x89PNG\r\n\x1a\n'
    png += chunk(b'IHDR', struct.pack('!2I5B', width, height, 8, 6, 0, 0, 0))
    png += chunk(b'IDAT', zlib.compress(scanlines, 9))
    png += chunk(b'IEND', b'')

    gdi32.DeleteObject(hbmp)
    gdi32.DeleteDC(hdc_mem)
    user32.ReleaseDC(0, hdc_screen)
    return base64.b64encode(png).decode('ascii')


class Handler(BaseHTTPRequestHandler):
    def _json(self, code, obj):
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.end_headers()
        self.wfile.write(json.dumps(obj, ensure_ascii=False).encode('utf-8'))

    def do_GET(self):
        path = urlparse(self.path).path
        try:
            if path == '/health':
                return self._json(200, {'ok': True})
            if path == '/windows':
                return self._json(200, {'windows': enum_windows()})
            return self._json(404, {'error': 'not_found'})
        except Exception as e:
            return self._json(500, {'error': str(e)})

    def do_POST(self):
        path = urlparse(self.path).path
        length = int(self.headers.get('Content-Length', '0'))
        body = self.rfile.read(length) if length else b'{}'
        try:
            data = json.loads(body.decode('utf-8')) if body else {}
            if path == '/launch':
                return self._json(200, launch(data['path'], data.get('args')))
            if path == '/activate':
                w = find_window(data.get('title', ''))
                if not w:
                    return self._json(404, {'error': 'window_not_found'})
                ok = activate_window(w['hwnd'])
                return self._json(200, {'ok': ok, 'window': w})
            if path == '/hotkey':
                return self._json(200, send_hotkey(data['keys']))
            if path == '/type':
                return self._json(200, send_text(data['text']))
            if path == '/click':
                return self._json(200, click(data['x'], data['y']))
            if path == '/screenshot':
                return self._json(200, {'png_base64': screenshot_b64()})
            return self._json(404, {'error': 'not_found'})
        except Exception as e:
            return self._json(500, {'error': str(e)})


def main():
    host = '127.0.0.1'
    port = 8765
    httpd = HTTPServer((host, port), Handler)
    print(f'desktop bridge listening on http://{host}:{port}', flush=True)
    httpd.serve_forever()


if __name__ == '__main__':
    main()
