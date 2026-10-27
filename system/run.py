import argparse
import socket
import threading
import webbrowser

import uvicorn


def main():
    parser = argparse.ArgumentParser(description="本机型式评价报告工作台")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()
    with socket.socket() as probe:
        try:
            probe.bind(("127.0.0.1", args.port))
        except OSError as exc:
            parser.error(f"端口 {args.port} 已被占用，请更换 --port；没有停止其他进程。{exc}")
    if not args.no_browser:
        timer = threading.Timer(1.5, webbrowser.open, args=(f"http://127.0.0.1:{args.port}",))
        timer.daemon = True
        timer.start()
    uvicorn.run("app:create_app", factory=True, host="127.0.0.1", port=args.port)


if __name__ == "__main__":
    main()
