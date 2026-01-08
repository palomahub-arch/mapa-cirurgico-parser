import logging
import os
from datetime import datetime


def configurar_logger():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    pasta_logs = os.path.join(base_dir, "logs")
    os.makedirs(pasta_logs, exist_ok=True)

    data = datetime.now().strftime("%Y-%m-%d")
    caminho_log = os.path.join(
        pasta_logs, f"mapa_cirurgico_{data}.log"
    )

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(caminho_log, encoding="utf-8"),
            logging.StreamHandler()
        ]
    )

    logging.info("Logger iniciado")
