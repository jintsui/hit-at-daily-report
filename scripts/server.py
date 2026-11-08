# -*- coding: utf-8 -*-
"""
HIT AT 效率诊断日报 · 本地服务
- 提供报告页面与静态资源
- POST /api/build  接收上传的 ZIP(遥测/能耗/日志 CSV) 与 XLSX(位移), 覆盖式重算并更新 HTML
用法: python server.py [--port 8760] [--open]
"""
import argparse, glob, json, os, re, sys, tempfile, threading, traceback, webbrowser
import http.server, socketserver, urllib.parse

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import build_report as B  # noqa

REPORT_GLOB = "HIT AT 效率诊断日报 *.html"
DATA_DIR = os.path.join(ROOT, "data")
TEMPLATE = os.path.join(HERE, "template.html")
ECHARTS = os.path.join(ROOT, "assets", "echarts.min.js")


def latest_report():
    fs = glob.glob(os.path.join(ROOT, REPORT_GLOB))
    return max(fs, key=os.path.getmtime) if fs else None


def do_build(zip_path, xlsx_path):
    """重新计算并覆盖生成 HTML, 返回输出路径"""
    RD = B.build(None, xlsx_path, zip_path)
    out = os.path.join(ROOT, "HIT AT 效率诊断日报 %s.html" % RD["meta"]["report_date"])
    B.render(RD, TEMPLATE, ECHARTS, out)
    return out, RD


def parse_multipart(body, ctype):
    """极简 multipart/form-data 解析, 返回 [(name, filename, bytes), ...]"""
    m = re.search(r'boundary=(?:"([^"]+)"|([^;]+))', ctype)
    if not m:
        return []
    boundary = (m.group(1) or m.group(2)).strip()
    delim = b"--" + boundary.encode()
    out = []
    for p in body.split(delim):
        if not p or p in (b"--", b"--\r\n", b"\r\n"):
            continue
        if p.startswith(b"\r\n"):
            p = p[2:]
        if p.endswith(b"\r\n"):
            p = p[:-2]
        if p.startswith(b"--"):
            continue
        i = p.find(b"\r\n\r\n")
        if i < 0:
            continue
        head = p[:i].decode("utf-8", "ignore")
        data = p[i + 4:]
        fm = re.search(r'filename="([^"]*)"', head)
        nm = re.search(r'name="([^"]*)"', head)
        out.append((nm.group(1) if nm else None, fm.group(1) if fm else None, data))
    return out


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=ROOT, **kw)

    def log_message(self, fmt, *args):
        sys.stderr.write("[server] %s\n" % (fmt % args))

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            rp = latest_report()
            if rp:
                self.send_response(302)
                self.send_header("Location", "/" + urllib.parse.quote(os.path.basename(rp)))
                self.end_headers()
                return
        return super().do_GET()

    def do_POST(self):
        if self.path.split("?")[0] != "/api/build":
            self.send_error(404, "not found")
            return
        try:
            ctype = self.headers.get("Content-Type", "")
            if "multipart/form-data" not in ctype:
                self._json(400, {"ok": False, "error": "expect multipart/form-data"})
                return
            n = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(n)
            parts = parse_multipart(body, ctype)

            tmp = tempfile.mkdtemp(prefix="hit_upload_")
            zip_path = xlsx_path = None
            saved = []
            for name, fn, data in parts:
                if not fn:
                    continue
                fn = os.path.basename(fn)
                dst = os.path.join(tmp, fn)
                with open(dst, "wb") as fp:
                    fp.write(data)
                saved.append(fn)
                low = fn.lower()
                if low.endswith(".zip"):
                    zip_path = dst
                elif low.endswith(".xlsx") or low.endswith(".xls"):
                    xlsx_path = dst

            if not saved:
                self._json(400, {"ok": False, "error": "no file uploaded"})
                return

            # 若 xlsx 在 zip 内, 解压后自动寻找
            extract_dir = os.path.join(tmp, "_csv")
            if zip_path:
                os.makedirs(extract_dir, exist_ok=True)
                B.extract_zip(zip_path, extract_dir)
                if not xlsx_path:
                    cand = glob.glob(os.path.join(extract_dir, "**", "*.xlsx"), recursive=True)
                    if cand:
                        xlsx_path = cand[0]
            if not xlsx_path:
                # 使用历史 data 目录中的 xlsx 兜底
                cand = glob.glob(os.path.join(DATA_DIR, "*.xlsx"))
                if cand:
                    xlsx_path = cand[0]
            if not xlsx_path:
                self._json(400, {"ok": False, "error": "未找到位移数据 XLSX，请一并上传"})
                return

            data_dir = extract_dir if zip_path else None
            if not data_dir:
                cand = glob.glob(os.path.join(DATA_DIR, "*.csv"))
                if cand:
                    data_dir = DATA_DIR

            RD = B.build(data_dir, xlsx_path, None)
            out = os.path.join(ROOT, "HIT AT 效率诊断日报 %s.html" % RD["meta"]["report_date"])
            B.render(RD, TEMPLATE, ECHARTS, out)
            self._json(200, {"ok": True, "report": os.path.basename(out),
                             "boxes": RD["movement"]["total"], "fleet": RD["meta"]["vehicle_count"],
                             "files": saved})
        except Exception as e:
            traceback.print_exc()
            self._json(500, {"ok": False, "error": str(e)})

    def _json(self, code, obj):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


class Server(socketserver.ThreadingTCPServer):
    allow_reuse_address = True
    daemon_threads = True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8760)
    ap.add_argument("--open", action="store_true")
    a = ap.parse_args()

    if not latest_report():
        # 尝试用 data/ 目录首次生成
        zips = glob.glob(os.path.join(DATA_DIR, "*.zip"))
        xlsxs = glob.glob(os.path.join(DATA_DIR, "*.xlsx"))
        if xlsxs:
            try:
                out, RD = do_build(zips[0] if zips else None, xlsxs[0])
                print("[server] 已生成报告: %s" % out)
            except Exception:
                traceback.print_exc()

    httpd = Server(("127.0.0.1", a.port), Handler)
    url = "http://127.0.0.1:%d/" % a.port
    print("[server] 报告服务已启动: %s" % url)
    print("[server] 按 Ctrl+C 停止")
    if a.open:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n[server] 已停止")


if __name__ == "__main__":
    main()
