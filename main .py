import logging
import os
import sys
from dotenv import load_dotenv
from flask import Flask
from routes.routes import graph_bp


class WhatsAppAPI():
    def __init__(self):
        load_dotenv()
        self.app = self.create_app()
        self.configure_logging()
        
    def load_configurations(self, app):
        pass

    def configure_logging(self):
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
            stream=sys.stdout,
        )

    def create_app(self):
        app = Flask(__name__)
        app.secret_key = os.environ.get('FLASK_SECRET_KEY')

        # Configuraciones
        self.load_configurations(app)

        # Rutas
        app.register_blueprint(graph_bp)

        return app

    def run(self):
        logging.info("Flask app started")
        self.app.run(host="0.0.0.0", port=int(os.getenv("PORT", 5000)), debug=True)


if __name__ == "__main__":
    whatsapp_api = WhatsAppAPI()
    whatsapp_api.run()
