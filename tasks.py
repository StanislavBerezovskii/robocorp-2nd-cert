import shutil

from robocorp import browser
from robocorp.tasks import task
from RPA.Archive import Archive
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Tables import Tables

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(slowmo=100)
    open_robot_order_website()
    download_csv_file()
    fill_order_form_with_csv_data()
    archive_receipts()
    clean_up()

def open_robot_order_website():
    """Navigates to the robot order website URL and clicks on the consent button"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    page = browser.page()
    page.click("button:text('OK')")

def fill_and_submit_order_form(order):
    """Fills the robot order form with data and submits it"""
    page = browser.page()
    page.select_option("#head", str(order["Head"]))
    page.click(f"#id-body-{order['Body']}")
    page.fill("input[placeholder='Enter the part number for the legs']", str(order["Legs"]))
    page.fill("#address", str(order["Address"]))
    while True:
        page.click("#order")
        if page.query_selector("#order-another"):
            pdf_path = store_receipt_as_pdf(order["Order number"])
            screenshot_path = screenshot_robot(order["Order number"])
            embed_screenshot_to_pdf(pdf_path, screenshot_path)
            page.click("#order-another")
            page.click("text=OK")
            break


def download_csv_file():
    """Downloads the robot orders CSV file"""
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv", overwrite=True)

def fill_order_form_with_csv_data():
    """Fills the order form with the robot orders CSV file data"""
    csv = Tables()
    orders = csv.read_table_from_csv("orders.csv", header=True, delimiters=",",  columns=["Order number", "Head", "Body", "Legs", "Address"])
    
    for row in orders:
        fill_and_submit_order_form(row)

def store_receipt_as_pdf(order_number):
    """Stores the order receipt as a PDF file and returns the file path"""
    page = browser.page()
    order_receipt_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    pdf_path = f"output/receipts/order_{order_number}_receipt.pdf"
    pdf.html_to_pdf(order_receipt_html, pdf_path)
    return pdf_path

def screenshot_robot(order_number):
    """Creates a screenshot of the ordered robot and returns the file path"""
    page = browser.page()
    screenshot_path = f"output/screenshots/order_{order_number}_screenshot.png"
    page.locator("#robot-preview-image").screenshot(path=screenshot_path)
    return screenshot_path

def embed_screenshot_to_pdf(pdf_path, screenshot_path):
    """Embeds the screenshot of the robot to the PDF receipt"""
    pdf = PDF()
    pdf.add_files_to_pdf(files=[f"{screenshot_path}:align=center",], target_document=pdf_path, append=True)

def archive_receipts():
    """Packs the order receipt PDFs into a .zip archive"""
    archive = Archive()
    archive.archive_folder_with_zip("output/receipts", "output/receipts.zip")

def clean_up():
    """Cleans up the output folder"""
    shutil.rmtree("output/receipts")
    shutil.rmtree("output/screenshots")
