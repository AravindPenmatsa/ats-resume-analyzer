import logging
import sys

def setup_logging():
    # 1. Configure the root logger to catch everything at DEBUG level.
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        handlers=[
            logging.StreamHandler()
        ]
    )

    # 2. Get the specific loggers for your application and third-party libraries.
    app_logger = logging.getLogger("app")
    app_logger.setLevel(logging.INFO)

    # 3. Set the logging level for noisy third-party libraries.
    logging.getLogger("weasyprint").setLevel(logging.DEBUG)
    logging.getLogger("fontTools").setLevel(logging.DEBUG)

    # 4. Silence overly verbose libraries if needed.
    logging.getLogger("uvicorn.access").setLevel(logging.INFO)
    
    return app_logger

logger = setup_logging()
