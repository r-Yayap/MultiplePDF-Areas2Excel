import logging, sys
def setup_logging(level=logging.INFO) -> None:
    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname).1s %(name)s: %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )
