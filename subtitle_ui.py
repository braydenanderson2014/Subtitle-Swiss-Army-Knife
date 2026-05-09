import argparse

from subtitle_tool import run_gui


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Subtitle Tool GUI launcher")
    parser.add_argument("--clear", action="store_true", help="Clear saved UI state/memory")
    parser.add_argument("--no-ai", action="store_true", help="Disable AI subtitle generation and save preference")
    parser.add_argument("--use-ai", action="store_true", help="Enable AI subtitle generation and save preference")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    use_ai = None
    if args.no_ai:
        use_ai = False
    elif args.use_ai:
        use_ai = True

    return run_gui(clear_memory=bool(args.clear), use_ai=use_ai)


if __name__ == "__main__":
    raise SystemExit(main())
