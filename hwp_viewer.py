#!/usr/bin/env python3
"""
HWP 뷰어
pyhwp(hwp5)로 HTML 변환 → tkinterweb으로 렌더링
한글 프로그램 없이 HWP 파일을 원본에 가깝게 열람합니다.

설치: pip install pyhwp tkinterweb olefile
빌드: pyinstaller --onefile --windowed hwp_viewer.py
"""

import os
import sys
import shutil
import tempfile
import subprocess
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from urllib.request import pathname2url

# ── 의존성 체크 ────────────────────────────────────────────────────────────────

try:
    from tkinterweb import HtmlFrame
    HAS_TKWEB = True
except ImportError:
    HAS_TKWEB = False

try:
    import hwp5  # noqa
    HAS_HWP5 = True
except ImportError:
    HAS_HWP5 = False

try:
    import olefile
    HAS_OLE = True
except ImportError:
    HAS_OLE = False


# ── HWP → HTML 변환 ────────────────────────────────────────────────────────────

def _convert_hwp_api(hwp_path: str, out_dir: str):
    """
    pyhwp Python API를 직접 호출해 HTML 변환합니다.
    PyInstaller 번들 환경에서 subprocess 대신 사용합니다.
    """
    old_argv = sys.argv[:]
    try:
        sys.argv = ['hwp5html', '--output', out_dir, hwp_path]
        from hwp5.hwp5html import main
        main()
    except SystemExit as e:
        if e.code not in (0, None):
            raise RuntimeError(f"hwp5html 변환 실패 (exit {e.code})")
    finally:
        sys.argv = old_argv


def _find_hwp5html() -> list:
    """일반 Python 환경에서 hwp5html 실행 명령어를 찾아 반환합니다."""
    import sysconfig

    found = shutil.which('hwp5html') or shutil.which('hwp5html.exe')
    if found:
        return [found]

    scripts = Path(sysconfig.get_path('scripts'))
    for name in ('hwp5html', 'hwp5html.exe'):
        cand = scripts / name
        if cand.exists():
            return [str(cand)]

    base = Path(sys.executable).parent
    for cand in [
        base / 'hwp5html',
        base / 'hwp5html.exe',
        base / 'Scripts' / 'hwp5html',
        base / 'Scripts' / 'hwp5html.exe',
    ]:
        if cand.exists():
            return [str(cand)]

    return [sys.executable, '-m', 'hwp5.hwp5html']


def convert_hwp_to_html(hwp_path: str) -> tuple:
    """
    HWP 파일을 HTML로 변환합니다.
    Returns: (html_file_path: str, out_dir: str)
    - out_dir 는 사용 후 caller 가 shutil.rmtree() 로 정리해야 합니다.
    """
    if not HAS_HWP5:
        raise RuntimeError(
            "pyhwp 패키지가 설치되어 있지 않습니다.\n\n"
            "pip install pyhwp 를 실행하거나,\n"
            "인터넷 가능한 PC에서 패키지를 포함해 빌드하세요."
        )

    out_dir = tempfile.mkdtemp(prefix='hwpview_')

    # PyInstaller 번들 환경: sys.executable이 .exe 자신이므로
    # subprocess 대신 Python API를 직접 호출해야 함
    if getattr(sys, 'frozen', False):
        try:
            _convert_hwp_api(hwp_path, out_dir)
        except Exception as exc:
            shutil.rmtree(out_dir, ignore_errors=True)
            raise RuntimeError(f"변환 실패: {exc}")
    else:
        # 일반 Python 환경: subprocess로 hwp5html 호출
        cmd = _find_hwp5html() + ['--output', out_dir, hwp_path]
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=60,
                encoding='utf-8',
                errors='replace',
            )
        except subprocess.TimeoutExpired:
            shutil.rmtree(out_dir, ignore_errors=True)
            raise RuntimeError("변환 시간이 초과되었습니다 (60초).")
        except FileNotFoundError:
            shutil.rmtree(out_dir, ignore_errors=True)
            raise RuntimeError("hwp5html 명령을 찾을 수 없습니다.\npip install pyhwp 를 실행해 주세요.")

        if result.returncode != 0:
            shutil.rmtree(out_dir, ignore_errors=True)
            detail = (result.stderr or result.stdout or '').strip()
            raise RuntimeError(f"변환 실패 (exit {result.returncode})\n\n{detail}")

    # 생성된 HTML/XHTML 파일 검색
    for pattern in ('*.html', '*.xhtml', '*.htm'):
        files = sorted(Path(out_dir).glob(pattern))
        if files:
            return str(files[0]), out_dir

    shutil.rmtree(out_dir, ignore_errors=True)
    raise RuntimeError("변환 후 HTML 파일을 찾을 수 없습니다.")


def path_to_file_url(path: str) -> str:
    """로컬 파일 경로 → file:// URL 변환 (Windows 경로 포함)."""
    abs_path = os.path.abspath(path)
    # Windows: C:\foo\bar → /C:/foo/bar
    url_part = pathname2url(abs_path)
    if not url_part.startswith('/'):
        url_part = '/' + url_part
    return 'file://' + url_part


# ── 텍스트 추출 (olefile 기반 fallback) ────────────────────────────────────────

def _extract_text_fallback(hwp_path: str) -> str:
    """pyhwp 없이 텍스트만 추출하는 fallback 함수입니다."""
    import struct, zlib

    HWPTAG_PARA_TEXT = 67

    def iter_records(data):
        pos = 0
        while pos + 4 <= len(data):
            hdr = struct.unpack_from('<I', data, pos)[0]
            tag  = hdr & 0x3FF
            size = (hdr >> 14) & 0x3FFFF
            pos += 4
            if size == 0x3FFFF:
                if pos + 4 > len(data):
                    break
                size = struct.unpack_from('<I', data, pos)[0]
                pos += 4
            yield tag, data[pos: pos + size]
            pos += size

    def para_to_str(rec):
        chars = []
        for i in range(0, len(rec) - 1, 2):
            code = struct.unpack_from('<H', rec, i)[0]
            if   code == 13: chars.append('\n')
            elif code ==  9: chars.append('\t')
            elif code >= 32:
                try: chars.append(chr(code))
                except (ValueError, OverflowError): pass
        return ''.join(chars)

    ole = olefile.OleFileIO(hwp_path)
    fh  = ole.openstream('FileHeader').read()
    compressed = bool(struct.unpack_from('<I', fh, 36)[0] & 1)

    parts, i = [], 0
    while ole.exists(f'BodyText/Section{i}'):
        raw = ole.openstream(f'BodyText/Section{i}').read()
        if compressed:
            try:    raw = zlib.decompress(raw, -zlib.MAX_WBITS)
            except: pass
        for tag, rec in iter_records(raw):
            if tag == HWPTAG_PARA_TEXT:
                parts.append(para_to_str(rec))
        i += 1
    ole.close()
    return ''.join(parts)


# ── GUI ────────────────────────────────────────────────────────────────────────

class HwpViewer(tk.Tk):
    """HWP 뷰어 메인 윈도우."""

    def __init__(self):
        super().__init__()
        self.title("HWP 뷰어")
        self.geometry("1060x780")
        self.minsize(600, 400)

        self._tmp_dirs: list = []        # 임시 디렉토리 (종료 시 정리)
        self._current_path = None        # 현재 열린 HWP 경로
        self._current_html  = None        # 현재 HTML 파일 경로

        self._build_ui()
        self.protocol('WM_DELETE_WINDOW', self._on_close)
        self._show_welcome()

    # ── UI 구성 ────────────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_menu()
        self._build_toolbar()
        self._build_viewer()
        self._build_statusbar()

    def _build_menu(self):
        bar = tk.Menu(self)

        file_menu = tk.Menu(bar, tearoff=0)
        file_menu.add_command(label="열기…",              accelerator="Ctrl+O", command=self.open_file)
        file_menu.add_separator()
        file_menu.add_command(label="브라우저에서 열기",  command=self.open_in_browser)
        file_menu.add_separator()
        file_menu.add_command(label="종료",               command=self._on_close)
        bar.add_cascade(label="파일", menu=file_menu)

        view_menu = tk.Menu(bar, tearoff=0)
        view_menu.add_command(label="뒤로",    command=self._go_back)
        view_menu.add_command(label="앞으로",  command=self._go_forward)
        view_menu.add_separator()
        view_menu.add_command(label="새로고침", command=self._reload)
        bar.add_cascade(label="보기", menu=view_menu)

        self.config(menu=bar)
        self.bind('<Control-o>', lambda _: self.open_file())

    def _build_toolbar(self):
        bar = tk.Frame(self, bg='#e8e8e8', pady=5)
        bar.pack(fill=tk.X, side=tk.TOP)

        btn_kw = dict(
            relief=tk.FLAT, cursor='hand2', padx=12, pady=4,
            bg='#4a90d9', fg='white', activebackground='#357abd',
        )
        tk.Button(bar, text="파일 열기",       command=self.open_file,       **btn_kw).pack(side=tk.LEFT, padx=(8, 4))
        tk.Button(bar, text="브라우저에서 열기", command=self.open_in_browser, **btn_kw).pack(side=tk.LEFT, padx=4)

        nav_kw = dict(
            relief=tk.FLAT, cursor='hand2', padx=8, pady=4,
            bg='#888', fg='white', activebackground='#666',
        )
        tk.Button(bar, text="◀", command=self._go_back,    **nav_kw).pack(side=tk.LEFT, padx=(12, 2))
        tk.Button(bar, text="▶", command=self._go_forward, **nav_kw).pack(side=tk.LEFT, padx=2)

        self._lbl_file = tk.Label(bar, text="파일을 선택하세요", bg='#e8e8e8', fg='#555')
        self._lbl_file.pack(side=tk.LEFT, padx=10)

    def _build_viewer(self):
        container = tk.Frame(self)
        container.pack(fill=tk.BOTH, expand=True)

        if HAS_TKWEB:
            self._html_view = HtmlFrame(
                container,
                horizontal_scrollbar='auto',
                vertical_scrollbar='auto',
                messages_enabled=False,
            )
            self._html_view.pack(fill=tk.BOTH, expand=True)
            self._view_mode = 'html'
        else:
            # tkinterweb 미설치 → Text 위젯 fallback
            sy = tk.Scrollbar(container, orient=tk.VERTICAL)
            sx = tk.Scrollbar(container, orient=tk.HORIZONTAL)
            sy.pack(side=tk.RIGHT,  fill=tk.Y)
            sx.pack(side=tk.BOTTOM, fill=tk.X)
            self._text_view = tk.Text(
                container,
                font=('Malgun Gothic', 11), wrap=tk.NONE,
                state=tk.DISABLED, relief=tk.FLAT, bg='white',
                padx=16, pady=12,
                yscrollcommand=sy.set,
                xscrollcommand=sx.set,
            )
            self._text_view.pack(fill=tk.BOTH, expand=True)
            sy.config(command=self._text_view.yview)
            sx.config(command=self._text_view.xview)
            self._view_mode = 'text'

    def _build_statusbar(self):
        self._status_var = tk.StringVar(value="준비")
        tk.Label(
            self,
            textvariable=self._status_var,
            anchor=tk.W,
            bg='#e0e0e0', fg='#444',
            relief=tk.SUNKEN, padx=8,
        ).pack(fill=tk.X, side=tk.BOTTOM)

    def _show_welcome(self):
        """시작 화면 표시 및 의존성 상태 알림."""
        if self._view_mode == 'html':
            self._html_view.load_html(WELCOME_HTML)

        missing = []
        if not HAS_HWP5:  missing.append("pyhwp")
        if not HAS_TKWEB: missing.append("tkinterweb")
        if not HAS_OLE:   missing.append("olefile")

        if missing:
            self._set_status(f"미설치 패키지: {', '.join(missing)} — pip install {' '.join(missing)}")
        else:
            mode = "HTML 렌더링 모드" if self._view_mode == 'html' else "텍스트 모드"
            self._set_status(f"준비 ({mode})")

    # ── 파일 열기 ───────────────────────────────────────────────────────────

    def open_file(self):
        path = filedialog.askopenfilename(
            title="HWP 파일 열기",
            filetypes=[("HWP 파일", "*.hwp *.HWP"), ("모든 파일", "*.*")],
        )
        if path:
            self._load(path)

    def _load(self, hwp_path: str):
        self._current_path = hwp_path
        self._lbl_file.config(text=Path(hwp_path).name)
        self._set_status("변환 중…")
        self.update_idletasks()

        # ── pyhwp → HTML ──────────────────────────────────────────────────
        if HAS_HWP5:
            try:
                html_file, out_dir = convert_hwp_to_html(hwp_path)
                self._tmp_dirs.append(out_dir)
                self._current_html = html_file
                url = path_to_file_url(html_file)

                if self._view_mode == 'html':
                    self._html_view.load_url(url)
                    self._set_status(f"{Path(hwp_path).name}  |  HTML 렌더링")
                else:
                    # tkinterweb 없음 → 자동으로 브라우저 열기
                    webbrowser.open(url)
                    self._set_status(f"브라우저에서 열림: {Path(html_file).name}")
                return
            except Exception as exc:
                if not HAS_OLE:
                    messagebox.showerror("변환 오류", str(exc))
                    self._set_status("오류")
                    return
                # olefile fallback 시도
                self._set_status("HTML 변환 실패, 텍스트 모드로 전환…")

        # ── olefile 텍스트 추출 fallback ────────────────────────────────
        if HAS_OLE:
            try:
                content = _extract_text_fallback(hwp_path)
                self._show_plain_text(content)
                self._set_status(f"{Path(hwp_path).name}  |  텍스트 모드 (pyhwp 미설치)")
            except Exception as exc:
                messagebox.showerror("오류", str(exc))
                self._set_status("오류")
        else:
            messagebox.showerror(
                "패키지 없음",
                "pyhwp 또는 olefile 패키지가 필요합니다.\n\n"
                "pip install pyhwp olefile tkinterweb"
            )
            self._set_status("패키지 없음")

    def _show_plain_text(self, content: str):
        """텍스트 내용을 Text 위젯 또는 HTML 뷰에 표시합니다."""
        if self._view_mode == 'html':
            escaped = content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            html = (
                '<html><head><meta charset="utf-8">'
                '<style>body{font-family:"Malgun Gothic",sans-serif;font-size:13px;'
                'white-space:pre-wrap;padding:20px;line-height:1.7}</style></head>'
                f'<body>{escaped}</body></html>'
            )
            self._html_view.load_html(html)
        else:
            self._text_view.config(state=tk.NORMAL)
            self._text_view.delete('1.0', tk.END)
            self._text_view.insert('1.0', content.strip())
            self._text_view.config(state=tk.DISABLED)

    # ── 브라우저에서 열기 ────────────────────────────────────────────────────

    def open_in_browser(self):
        if self._current_html and Path(self._current_html).exists():
            webbrowser.open(path_to_file_url(self._current_html))
        elif self._current_path:
            self._set_status("변환 중…")
            self.update_idletasks()
            try:
                html_file, out_dir = convert_hwp_to_html(self._current_path)
                self._tmp_dirs.append(out_dir)
                self._current_html = html_file
                webbrowser.open(path_to_file_url(html_file))
                self._set_status("브라우저에서 열기 완료")
            except Exception as exc:
                messagebox.showerror("오류", str(exc))
        else:
            messagebox.showinfo("알림", "먼저 HWP 파일을 열어주세요.")

    # ── 탐색 ────────────────────────────────────────────────────────────────

    def _go_back(self):
        if self._view_mode == 'html':
            try: self._html_view.go_back()
            except: pass

    def _go_forward(self):
        if self._view_mode == 'html':
            try: self._html_view.go_forward()
            except: pass

    def _reload(self):
        if self._current_html:
            self._html_view.load_url(path_to_file_url(self._current_html))

    # ── 유틸 ────────────────────────────────────────────────────────────────

    def _set_status(self, msg: str):
        self._status_var.set(msg)

    def _on_close(self):
        for d in self._tmp_dirs:
            shutil.rmtree(d, ignore_errors=True)
        self.destroy()


# ── 시작 화면 HTML ─────────────────────────────────────────────────────────────

WELCOME_HTML = """<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body {
    font-family: "Malgun Gothic", "Apple SD Gothic Neo", sans-serif;
    display: flex; flex-direction: column;
    align-items: center; justify-content: center;
    height: 100vh; margin: 0;
    background: #f5f7fa; color: #333;
  }
  h1  { font-size: 28px; margin-bottom: 8px; color: #4a90d9; }
  p   { font-size: 14px; color: #666; margin: 4px 0; }
  .tip { margin-top: 24px; background: #fff; border-radius: 8px;
         padding: 16px 24px; box-shadow: 0 2px 8px rgba(0,0,0,.08);
         font-size: 13px; line-height: 1.8; }
  kbd { background:#eee; border-radius:3px; padding:1px 6px;
        font-size:12px; border:1px solid #ccc; }
</style>
</head>
<body>
  <h1>HWP 뷰어</h1>
  <p>한글 프로그램 없이 HWP 파일을 열람합니다</p>
  <div class="tip">
    <b>사용법</b><br>
    • 상단 <b>파일 열기</b> 버튼 클릭 또는 <kbd>Ctrl+O</kbd><br>
    • <b>브라우저에서 열기</b>를 누르면 더 나은 렌더링으로 확인 가능
  </div>
</body>
</html>"""


# ── 진입점 ─────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    app = HwpViewer()
    app.mainloop()
