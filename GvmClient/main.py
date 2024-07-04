from gvmclient import *
from reportmanager import *

def main():
    """
    Función principal para la ejecución del programa.
    """
    path = '/'
    username = ''
    password = ''
    filters_file = 'discovery.txt'  # Archivo con los IDs de los filtros
    smtp_server = ''
    smtp_port = 25
    smtp_user = 'openvas@jojos.jeje'
    smtp_password = ""

    report_manager = ReportManager(filters_file, smtp_server, smtp_port, smtp_user, smtp_password)
    filters = report_manager.read_filter_ids()

    client = GvmdClient(path, username, password)  # Crea una instancia de GvmdClient
    client.connect()

    try:
        
        for nombre_filtro, filter_id in filters:     
               
        #     # Descarga el reporte como CSV y guarda el archivo
        #     # con el nombre report_report_id.csv en el directorio actual.
        #     # Si el reporte no se pudo descargar, imprime un mensaje de error.
            last_report = client.get_reports_list(filter_id)
            report = client.download_report_as_csv(last_report)
            
            if report:
                    with open(f"./reports/{nombre_filtro}.csv", 'wb') as csv_file:
                        csv_file.write(report)
                    print(f"CSV report saved as {nombre_filtro}.csv")


    #    # Definir el directorio y nombre del excel del report
        csv_folder = "./reports"
        path_report = "./combined_reports/"
        excel_filename = "Report.xlsx"
        full_path_file = path_report + excel_filename
        today = datetime.today().strftime('%Y-%m-%d')
        full_path_file_with_date = full_path_file + "_" + today + ".xlsx"
    #    # Procesar todos los csv en un solo reporte.
        report_manager.csvs_to_excel(csv_folder, full_path_file)
        
    #    # Agregar un gráfico resumen al principio del archivo Excel
        report_manager.add_summary_chart(full_path_file)


        to_address = ['correo1@jojos.jeje', 'correo2@jojos.jeje']
        subject = 'Reporte generado'
        body = 'Adjunto se encuentra el reporte: ' + excel_filename
        for correo in to_address:
            report_manager.send_email(correo, subject, body, full_path_file_with_date)
        
    except Exception as e:
        print(f"Error downloading or saving CSV report: {e}")



if __name__ == "__main__":
    main()