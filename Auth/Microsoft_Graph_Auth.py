from msal import ConfidentialClientApplication
import os
import requests
from urllib.parse import quote
from dotenv import load_dotenv

class MicrosoftGraphAuthenticator:
    def __init__(self):
        # Cargar variables de entorno
        load_dotenv()
        self.client_id = os.getenv("MICROSOFT_CLIENT_ID")
        self.client_secret = os.getenv("MICROSOFT_CLIENT_SECRET")
        self.tenant_id = os.getenv("MICROSOFT_TENANT_ID")
        
        # El link es la puerta oficial de entrada a Microsoft Azure AD (No cambia), {tenant_id} Es el identificador único de tu organización en Azure (cambia dependiendo de quien tiene acceso)
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        
        #.Default le da permisos que fueron concedidos a esta app en Azure.
        self.scope = ["https://graph.microsoft.com/.default"]
        
        #Id de identificación de la API de Microsoft Graph
        self.user_id = "RodrigoAguilera@Eunoia8.onmicrosoft.com"

        #URL base de Microsoft Graph API
        self.graph_url = "https://graph.microsoft.com/v1.0"

        # Token de acceso temporal generado por Microsoft que te da permiso para hacer llamadas a la API
        self.token = None

        #Le manda el token de acceso a microsoft para que de acceso a la API
        self.headers = None

        # Inicializar cliente MSAL
        self.app = ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=self.authority
        )

    def get_access_token(self):
        #Obtiene un token de acceso válido desde Microsoft Identity Platform.
        result = self.app.acquire_token_for_client(scopes=self.scope)
        if "access_token" in result:
            #Asigna el valor de el token de acceso a la variable token
            self.token = result["access_token"]
            return self.token
        else:
            raise Exception(f" Error al obtener token: {result.get('error_description')}")

    def get_headers(self):
        """
        Devuelve los headers de autorización listos para usar con Microsoft Graph.
        Refresca el token si aún no existe en memoria.
        """
        if not self.token:
            self.get_access_token()
        self.headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
        }
        return self.headers
    
    def get_file_id(self, ruta_archivo: str) -> str:
        access_token = self.get_access_token()

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

        ruta_archivo_codificada = quote(ruta_archivo)
        url = f"{self.graph_url}/users/{self.user_id}/drive/root:/{ruta_archivo_codificada}"

        response = requests.get(url, headers=headers)

        if response.status_code == 200:
            return response.json()["id"]
        else:
            raise Exception(f"❌ Error obteniendo ID del archivo: {response.status_code} - {response.text}")

    def get_folder_id(self, ruta_carpeta: str) -> str:
        #Obtiene el ID de una carpeta en OneDrive por su ruta relativa
        access_token = self.get_access_token()
        headers= {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        url = f"{self.graph_url}/users/{self.user_id}/drive/root:/{ruta_carpeta}"
        response = requests.get(url, headers=headers)

        if response.status_code == 200:
            return response.json()["id"]
        else:
            raise Exception(f"❌ Error obteniendo ID de la carpeta: {response.status_code} - {response.text}")

