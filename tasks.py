"""Librerías"""
from robocorp.tasks import task
from robocorp import browser
from RPA.PDF import PDF
from fpdf import FPDF
import csv
import os
import shutil
import zipfile
from RPA.HTTP import HTTP

"""Clases y métodos"""

class CustomPDF(FPDF):
    def header(self):
        # Deja el encabezado vacío o incluye un encabezado genérico si lo necesitas
        pass



def delete_directory(directory_path):
    """Elimina un directorio y todo su contenido."""
    if os.path.exists(directory_path):
        shutil.rmtree(directory_path)
        print(f"Directorio '{directory_path}' eliminado exitosamente.")
    else:
        print(f"El directorio '{directory_path}' no existe.")

def clean_up_directories():
    """Elimina los directorios de salida después de crear el ZIP."""
    delete_directory('output/png')
    delete_directory('output/pdf')

def create_directory(path):
    if not os.path.exists(path):
        os.makedirs(path)

def capture_div_screenshot(order_number):
    page = browser.page()
    img_element = page.locator("#robot-preview-image")
    screenshot_path = f"output/png/{order_number}.png"
    img_element.screenshot(path=screenshot_path)
    return screenshot_path

def store_receipt_as_pdf(order_number):
    page = browser.page()
    if page.locator("p.badge.badge-success").count() > 0:
        order_number = page.text_content("p.badge.badge-success")
        print(f"Texto capturado: {order_number}")
    else:
        print("Elemento no encontrado.")
        
    # Captura solo el HTML del recibo sin incluir el número de orden
    receipt = page.locator("#receipt").inner_html()
    pdf = CustomPDF()
    pdf.add_page()
    pdf.set_font('Times', '', 12)
    
    # Escribe el contenido del recibo en el PDF
    pdf.write_html(receipt)

    # Captura de pantalla del robot
    screenshot_path = capture_div_screenshot(order_number)

    # Añadir la imagen en la misma página que el recibo
    pdf.image(screenshot_path, x=10, y=pdf.get_y() + 10, w=100)

    # Guarda el PDF con el número de orden como nombre de archivo
    pdf_path = f"output/pdf/{order_number}.pdf"
    pdf.output(pdf_path, 'F')
    return pdf_path

def close_annoying_modal():
    page = browser.page()
    ok_button_locator = "button:text('OK')"
    if page.locator(ok_button_locator).count() > 0:
        page.click(ok_button_locator)
        print("Botón 'OK' pulsado.")
    else:
        print("No se encontró el botón 'OK'.")

def create_zip_from_directory(directory_path):
    zip_filename = "mi_archivo.zip"
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for foldername, subfolders, filenames in os.walk(directory_path):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                zipf.write(file_path, os.path.relpath(file_path, directory_path))
    print(f"Archivo ZIP '{zip_filename}' creado exitosamente con todos los archivos de '{directory_path}'.")

def get_orders():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def fill_form_with_csv_data(path):
    try:
        with open(path, 'r') as archivo_csv:
            lector_csv = csv.DictReader(archivo_csv)
            for robot_order in lector_csv:
                print(robot_order)
                fill_and_submit_robot_form(robot_order)
    except Exception as e:
        print(f"Error al procesar el archivo CSV: {e}")

def fill_and_submit_robot_form(robot_order):
    page = browser.page()

    try:
        close_annoying_modal()

        page.select_option("select#head.custom-select", index=int(robot_order["Head"]))
        numero_boton = robot_order["Body"]
        boton_selector = f"input[type='radio']#id-body-{numero_boton}"
        boton = page.locator(boton_selector)
        boton.click()

        page.fill("#address", robot_order["Address"])
        page.fill("input[placeholder='Enter the part number for the legs']", robot_order["Legs"])
        
        page.click("button:text('Preview')")
        page.click("button:text('Show model info')")
        page.click("button:text('ORDER')")

        while page.locator(".alert-danger").count() > 0:
            print("Alerta detectada. Reintentando...")
            page.click("button:text('ORDER')")

        print(f"Orden completada: {robot_order}")
        store_receipt_as_pdf(robot_order["Order number"])

    except Exception as e:
        print(f"Error al procesar la orden {robot_order}: {e}")

    try:
        if page.locator("button:text('Order Another Robot')").count() > 0 and int(robot_order["Order number"]) < 20:
            page.click("button:text('Order Another Robot')")
            print("Preparando para la siguiente orden.")
    except Exception as e:
        print(f"Error al hacer clic en 'Order Another Robot': {e}")

@task
def order_robots_from_RobotSpareBin():
    browser.configure(slowmo=100)
    create_directory('output/pdf')
    create_directory('output/png')
    open_robot_order_website()
    get_orders()
    fill_form_with_csv_data("orders.csv")
    create_zip_from_directory("output/pdf")
    clean_up_directories()  # Limpia los directorios después de crear el ZIP


