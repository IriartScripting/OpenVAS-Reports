import sys
from argparse import ArgumentParser, Namespace, RawTextHelpFormatter
import xml.dom.minidom
import xml.etree.ElementTree as ET
import base64

from gvm.connections import UnixSocketConnection  # Importa la clase UnixSocketConnection desde gvm.connections
from gvm.protocols.gmp import Gmp  # Importa la clase Gmp desde gvm.protocols.gmp
from gvmtools.helper import Table  # Importa la clase Table desde gvmtools.helper
from gvm.errors import GvmError  # Importa la clase GvmError desde gvm.errors
from gvm.transforms import EtreeCheckCommandTransform

class GvmdClient:
    def __init__(self, socket, username, password):
        """
        Inicializa la clase GvmdClient.

        Args:
        - socket (str): Ruta al socket de Unix para la conexión con GVMD.
        - username (str): Nombre de usuario para autenticarse en GVMD.
        - password (str): Contraseña para autenticarse en GVMD.
        """
        self.socket = socket
        self.username = username
        self.password = password
        self.gmp = None  # Inicializa como None; se asignará una instancia de Gmp después de la conexión
    
    def connect(self):
        """
        Establece la conexión con GVMD y autentica al usuario.

        Raises:
        - GvmError: Si ocurre un error durante la conexión o autenticación.
        """
        self.connection = UnixSocketConnection(path=self.socket)  # Crea una conexión UnixSocketConnection
        self.transform = EtreeCheckCommandTransform()
        try:
            with Gmp(self.connection, transform=self.transform) as gmp:  # Crea una instancia de Gmp y la usa dentro de un contexto 'with'
                if gmp.is_connected():  # Verifica si la conexión con Gmp está establecida
                    gmp.authenticate(self.username, self.password)  # Intenta autenticarse con el nombre de usuario y la contraseña
                    print(f"Connected as {self.username}\n")  # Imprime un mensaje de conexión exitosa
        except GvmError as e:
            # Maneja errores específicos que pueden ocurrir durante la conexión o autenticación
            if "Failed to connect" in str(e):
                print(f"Failed to connect: {e}")
            elif "AuthenticationError" in str(e):
                print(f"Authentication failed: {e}")
            elif "ConnectionError" in str(e):
                print(f"Connection error: {e}")
            else:
                print(f"An unknown error occurred during connection: {e}")
#            raise  # Eleva la excepción para manejarla en niveles superiores

    def get_reports_list(self, filter_id):
        """
        Obtiene la lista de informes disponibles desde GVMD.

        Returns:
        - response (list): Lista de informes obtenidos.

        Raises:
        - Exception: Si no está conectado a GVMD. Conectar primero.
        - GvmError: Si ocurre un error al obtener los informes.
        """

    
    
        try:
            with Gmp(self.connection) as gmp:  # Crea una instancia de Gmp y la usa dentro de un contexto 'with'
                if gmp.is_connected():  # Verifica si la conexión con Gmp está establecida
                    gmp.authenticate(self.username, self.password)  # Intenta autenticarse con el nombre de usuario y la contraseña
                    
                    response = gmp.get_reports(filter_id=filter_id, note_details=0, override_details=0)
                    response_xml = response # Obtén la salida XML en forma de cadena
                    
                # Formatea la cadena XML para que sea legible
                    dom = xml.dom.minidom.parseString(response_xml)
                    pretty_xml_as_string = dom.toprettyxml()
                    report_id = None
                # Parsear el XML para obtener el report_id
                    root = ET.fromstring(response_xml)
                    report_ids = []
                    for report in root.findall('.//report'):
                        report_id = report.get('id')
                        if report_id:
                            report_ids.append(report_id)

                    return report_id
                            
        except GvmError as e:
            # Maneja errores específicos que pueden ocurrir al intentar obtener los informes
            print(f"Error fetching reports: {e}")
            if "RemoteException" in str(e):
                print(f"Remote exception occurred: {e}")
            elif "MissingElementError" in str(e):
                print(f"Missing element in the response: {e}")
            elif "XmlError" in str(e):
                print(f"XML error occurred: {e}")
            else:
                print(f"An unknown error occurred: {e}")
            raise  # Eleva la excepción para manejarla en niveles superiores
        
    def download_report_as_csv(self, report_id_proccessed):

        try:
            with Gmp(self.connection) as gmp:
                if gmp.is_connected():
                    gmp.authenticate(self.username, self.password)

                    # Usar el report_format_id para CSV
                    csv_report_format_id = 'c1645568-627a-11e3-a660-406186ea4fc5'  # Asegúrate de que este ID es el correcto para tu instalación
                    
                    response = gmp.get_report(
                            report_id=report_id_proccessed, 
                            report_format_id=csv_report_format_id, 
                            details=True, 
                            ignore_pagination=True)

                    report_id = None

                    
                    # Encontrar y decodificar el contenido CSV
                    start_tag = "</report_format>"
                    end_tag = "</report>"
                    start_index = response.find(start_tag)
                    end_index = response.find(end_tag, start_index + len(start_tag))
                    
                    if start_index != -1 and end_index != -1:
                        base64_content = response[start_index + len(start_tag):end_index].strip()
                        binary_base64_encoded_csv = base64_content.encode("ascii")
                        binary_csv = base64.b64decode(binary_base64_encoded_csv)
                        
                        return binary_csv
                        
                    else:
                        print("\nNo fue posible procesar el contenido\n\n" + response)

 #                   else:
#                        return start_index


        except GvmError as e:
            print(f"Error fetching report {report_id}: {e}")
            if "InvalidArgumentError" in str(e):
                print(f"Invalid argument: {e}")
            elif "RemoteException" in str(e):
                print(f"Remote exception occurred: {e}")
            elif "MissingElementError" in str(e):
                print(f"Missing element in the response: {e}")
            elif "XmlError" in str(e):
                print(f"XML error occurred: {e}")
            raise



