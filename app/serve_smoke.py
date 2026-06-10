import argparse
import os
from getpass import getpass

from app.runner import load_config, run_serve


DEFAULT_CONFIG_PATH = os.path.join("config", "config.yaml")


def _positive_int(value: str) -> int:
    parsed = int(value)
    if parsed < 0:
        raise argparse.ArgumentTypeError("must be 0 or greater")
    return parsed


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Smoke-test the WAI Array of Services flow. This searches the input Excel rows, "
            "opens the first owned student, fills the configured WAI 0.5 hour fields, scrolls to the bottom, "
            "and pauses before closing. It does not click Save unless --save is provided."
        )
    )
    parser.add_argument("--excel", required=True, help="Input Excel file using the normal Workabili-Bot headers.")
    parser.add_argument(
        "--config",
        default=DEFAULT_CONFIG_PATH,
        help=f"Path to config YAML. Default: {DEFAULT_CONFIG_PATH}",
    )
    parser.add_argument(
        "--username",
        default="",
        help="WAI username. Defaults to config credentials.username, then prompts.",
    )
    parser.add_argument(
        "--password-env",
        default="WAI_PASSWORD",
        help="Environment variable to read the WAI password from. Default: WAI_PASSWORD",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run browser headless. Omit this when you want to observe the page.",
    )
    parser.add_argument(
        "--slow-mo-ms",
        type=_positive_int,
        default=None,
        help="Override Playwright slow motion in milliseconds. Defaults to config run.slow_mo_ms.",
    )
    parser.add_argument(
        "--hours",
        default="0.5",
        help="Hours value to fill in the WAI Array of Services fields. Default: 0.5",
    )
    parser.add_argument(
        "--save",
        action="store_true",
        help="Click Save after filling the first owned student. Without this flag, the command is a dry-run.",
    )
    return parser.parse_args()


def resolve_username(args: argparse.Namespace, cfg: dict) -> str:
    username = (args.username or "").strip()
    if username:
        return username

    username = str(cfg.get("credentials", {}).get("username", "") or "").strip()
    if username:
        return username

    return input("WAI Username: ").strip()


def resolve_password(args: argparse.Namespace) -> str:
    env_name = str(args.password_env or "").strip()
    if env_name:
        password = os.environ.get(env_name, "")
        if password:
            return password

    return getpass("WAI Password: ")


def main() -> int:
    args = parse_args()
    cfg = load_config(args.config)

    username = resolve_username(args, cfg)
    password = resolve_password(args)

    if not username:
        raise SystemExit("Missing username.")
    if not password:
        raise SystemExit("Missing password.")

    mode = "save run" if args.save else "dry-run check"
    print(f"Opening WAI and running Serve {mode}...")
    if args.save:
        print("This command WILL click Save after filling the first owned student.")
    else:
        print("This command will fill the Array of Services hour fields, but it will not click Save.")
    print("The browser will pause after the first owned student is filled/saved.")

    run_serve(
        excel_path=args.excel,
        config_path=args.config,
        username=username,
        password=password,
        log_fn=print,
        override_headless=bool(args.headless),
        override_slow_mo_ms=args.slow_mo_ms,
        pause_after_first_fill=True,
        save_after_fill=bool(args.save),
        hours=str(args.hours),
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
