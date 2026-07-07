"""
scheduler.py
------------
Optional weekly scheduler (bonus deliverable). Runs the full pipeline
every Monday morning and logs execution. Intended to run as a
long-lived process (e.g. `python src/scheduler.py`) inside a
container, systemd service, or CI scheduled job.
"""

from __future__ import annotations

import glob
import logging
import os
import sys

from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.triggers.cron import CronTrigger

sys.path.insert(0, os.path.dirname(__file__))
from main import run_pipeline  # noqa: E402

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("scheduler")

INPUT_GLOB = os.environ.get("PHA_INPUT_GLOB", "data/*.xlsx")
OUTPUT_DIR = os.environ.get("PHA_OUTPUT_DIR", "outputs")


def weekly_job():
    logger.info("Starting scheduled weekly run...")
    files = sorted(glob.glob(INPUT_GLOB))
    if not files:
        logger.warning("No input files found matching %s; skipping this run.", INPUT_GLOB)
        return
    try:
        portfolio = run_pipeline(files, OUTPUT_DIR, build_deck=True)
        logger.info("Weekly run complete: %d project(s) processed.", len(portfolio))
    except Exception:
        logger.exception("Weekly run failed.")


if __name__ == "__main__":
    scheduler = BlockingScheduler()
    # Every Monday at 07:00 local time
    scheduler.add_job(weekly_job, CronTrigger(day_of_week="mon", hour=7, minute=0), id="weekly_health_report")
    logger.info("Scheduler started. Weekly job will run every Monday at 07:00.")
    logger.info("Running an initial pass now so today's outputs are available immediately...")
    weekly_job()
    try:
        scheduler.start()
    except (KeyboardInterrupt, SystemExit):
        logger.info("Scheduler stopped.")
