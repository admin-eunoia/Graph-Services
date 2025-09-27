import logging
import os
import sys
from dotenv import load_dotenv
from flask import Flask, request
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

from routes.routes import graph_bp
from Postgress.connection import init_db, SessionLocal

class GraphAPIApp:
    def __init__(self):
        load_dotenv()
        self.app = self.create_app()
        self.configure_logging()

    def configure_logging(self):
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
            stream=sys.stdout,
        )

    def create_app(self):
        app = Flask(__name__)
        app.secret_key = os.environ.get("FLASK_SECRET_KEY")

        # Inicializar DB
        init_db()

        # Rate limiting
        default_limits = os.getenv("RATE_LIMITS", "120 per minute; 5000 per hour")
        limits_list = [limit.strip() for limit in default_limits.split(";") if limit.strip()]
        self.limiter = Limiter(
            get_remote_address,
            app=app,
            default_limits=limits_list,
            headers_enabled=True,
        )

        trusted_api_key = os.getenv("TRUSTED_API_KEY")

        @self.limiter.request_filter
        def _exempt_trusted_key():
            if not trusted_api_key:
                return False
            return request.headers.get("X-Api-Key") == trusted_api_key

        # Middleware de DB session
        @app.before_request
        def create_session():
            request.environ["db_session"] = SessionLocal()

        @app.teardown_request
        def shutdown_session(exception=None):
            db = request.environ.get("db_session")
            if db:
                if exception:
                    db.rollback()
                else:
                    db.commit()
                db.close()

        # Rutas
        app.register_blueprint(graph_bp, url_prefix="/graph")
        return app

    def run(self):
        logging.info("Flask app started")
        self.app.run(host="0.0.0.0", port=int(os.getenv("PORT", 8000)), debug=True)

if __name__ == "__main__":
    graph_api = GraphAPIApp()
    graph_api.run()
