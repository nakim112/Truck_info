#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
INC4 HUB 차량 정보 관리 시스템 - 웹 버전 실행 진입점.

exe 없이 실행:
    python run_web.py
    python run_web.py --host 0.0.0.0 --port 8080

환경변수:
    TRUCK_INFO_DATA   레코드 저장 JSON 경로 (기본: web/data.json)
"""
import argparse

import uvicorn


def main() -> None:
    parser = argparse.ArgumentParser(description="INC4 HUB 차량 정보 시스템 (Web)")
    parser.add_argument("--host", default="127.0.0.1", help="바인드 호스트 (기본 127.0.0.1)")
    parser.add_argument("--port", type=int, default=8000, help="포트 (기본 8000)")
    parser.add_argument("--reload", action="store_true", help="코드 변경 시 자동 리로드")
    args = parser.parse_args()

    uvicorn.run("web.server:app", host=args.host, port=args.port, reload=args.reload)


if __name__ == "__main__":
    main()
