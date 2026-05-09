import argparse
import sys

from subtitle_tool import build_parser, run_cli_action


CLI_MODES = {
    "scan",
    "remove",
    "include",
    "extract",
    "tag-audio-language",
    "sync-subtitles",
}


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    mode = getattr(args, "mode", "")
    if mode not in CLI_MODES:
        print(
            "Usage: subtitle_cli.py <mode> [options]\n"
            "Modes: scan, remove, include, extract, tag-audio-language, sync-subtitles",
            file=sys.stderr,
        )
        return 2

    return run_cli_action(args)


if __name__ == "__main__":
    raise SystemExit(main())
