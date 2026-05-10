import argparse
import os
from getpass import getpass

from playwright.sync_api import sync_playwright

from app.runner import (
    determine_search_outcome,
    get_found_row_owning_org,
    goto_student_records,
    is_already_owned_by_us,
    load_config,
    login,
    open_existing_student_edit,
    resolve_school,
    search_by_ssid,
    transfer_prior_year_student,
)


DEFAULT_CONFIG_PATH = os.path.join("config", "config.yaml")


def _positive_int(value: str) -> int:
    parsed = int(value)
    if parsed < 0:
        raise argparse.ArgumentTypeError("must be 0 or greater")
    return parsed


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Smoke-test the WAI student search flow without launching the GUI, "
            "creating records, requesting transfers, or writing to the run ledger."
        )
    )
    parser.add_argument("--ssid", required=True, help="Student SSID to search for.")
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
        "--no-pause",
        action="store_true",
        help="Close the browser immediately after the search instead of waiting for Enter.",
    )
    parser.add_argument(
        "--exercise-prior-year-transfer",
        action="store_true",
        help=(
            "After a FOUND search result, open Edit, click the prior-year Transfer link if present, "
            "fill the modal dropdowns, and pause before Save. This never clicks Save."
        ),
    )
    parser.add_argument(
        "--school",
        default="",
        help=(
            "School of Attendance value to use in the prior-year transfer modal. "
            "Required with --exercise-prior-year-transfer unless --aeries-school is provided."
        ),
    )
    parser.add_argument(
        "--aeries-school",
        default="",
        help=(
            "Aeries school name/key to resolve through config school_mapping for the prior-year "
            "transfer modal. Alternative to --school."
        ),
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


def resolve_prior_year_wai_school(args: argparse.Namespace, cfg: dict) -> str:
    school = (args.school or "").strip()
    if school:
        return school

    aeries_school = (args.aeries_school or "").strip()
    if aeries_school:
        return resolve_school(aeries_school, cfg)

    raise SystemExit(
        "--exercise-prior-year-transfer requires --school or --aeries-school "
        "so the modal School of Attendance can be filled."
    )


def main() -> int:
    args = parse_args()
    cfg = load_config(args.config)

    username = resolve_username(args, cfg)
    password = resolve_password(args)

    if not username:
        raise SystemExit("Missing username.")
    if not password:
        raise SystemExit("Missing password.")

    slow_mo_ms = args.slow_mo_ms
    if slow_mo_ms is None:
        slow_mo_ms = int(cfg.get("run", {}).get("slow_mo_ms", 0) or 0)

    print(f"Opening WAI and searching SSID {args.ssid}...")
    print("This command will not request envelope transfer, create a record, or write the ledger.")
    if args.exercise_prior_year_transfer:
        print("Prior-year transfer dry-run is enabled: the modal may be filled, but Save will not be clicked.")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=bool(args.headless),
            slow_mo=slow_mo_ms if slow_mo_ms > 0 else None,
        )
        context = browser.new_context(viewport={"width": 1600, "height": 1400})
        page = context.new_page()

        try:
            login(page, username, password, cfg)
            goto_student_records(page, cfg=cfg, run_id="smoke")
            search_by_ssid(page, args.ssid, cfg)

            outcome = determine_search_outcome(page, timeout_ms=5000)
            print(f"Search outcome: {outcome}")

            if outcome == "FOUND":
                owning_org = get_found_row_owning_org(page, args.ssid)
                owned_by_us = is_already_owned_by_us(owning_org, cfg)
                print(f"Owning org: {owning_org}")
                print(f"Owned by configured org: {owned_by_us}")

                if args.exercise_prior_year_transfer:
                    wai_school = resolve_prior_year_wai_school(args, cfg)
                    print(f"Opening Edit and testing prior-year Transfer modal with school: {wai_school}")
                    open_existing_student_edit(page)
                    prior_year_transfer = transfer_prior_year_student(
                        page,
                        cfg,
                        wai_school,
                        save=False,
                        log_fn=print,
                        run_id="smoke",
                    )
                    if prior_year_transfer:
                        print("Prior-year Transfer modal filled successfully. Save was skipped.")
                        print(f"Transfer to: {prior_year_transfer['transfer_to_project']}")
                        print(f"School: {prior_year_transfer['school']}")
                    else:
                        print("No prior-year Transfer link was visible after opening Edit.")
            elif args.exercise_prior_year_transfer:
                print("Prior-year transfer dry-run skipped because the student was not found.")

            if not args.no_pause:
                input("Browser is paused for inspection. Press Enter to close...")

        finally:
            context.close()
            browser.close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
