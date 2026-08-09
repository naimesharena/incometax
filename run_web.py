#!/usr/bin/env python3
"""Start the Open ITR Suite web application (ITR-1 / ITR-2 / ITR-4).

Usage::

    pip install -r requirements.txt
    python run_web.py [--port 5000] [--debug]

Then open http://127.0.0.1:5000/ in your browser.
"""

import argparse

from itr1.webapp import create_app


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=5000)
    parser.add_argument("--debug", action="store_true")
    args = parser.parse_args()

    app = create_app()
    print(f"\n  Open ITR Suite — ITR-1 / ITR-2 / ITR-4 (AY 2026-27)")
    print(f"  → http://{args.host}:{args.port}/\n")
    app.run(host=args.host, port=args.port, debug=args.debug)


if __name__ == "__main__":
    main()
