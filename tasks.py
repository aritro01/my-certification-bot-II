from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables 
from RPA.PDF import PDF
import os
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    open_robot_order_website()
    download_excel_file()
    orders = get_orders()
    for order in orders:
        fill_the_form(order)
    archive_receipts()

def open_robot_order_website():
    # TODO: Implement your function here
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def close_annoying_modal():
    page = browser.page()
    while page.is_visible(".modal-body"):
        page.click("text=OK")

def download_excel_file():
    '''Download the .csv file'''
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv", target_file="output/orders.csv", overwrite=True)

def get_orders():
    '''Open the .csv file'''
    table = Tables()
    return table.read_table_from_csv("output/orders.csv", header=True)

def fill_the_form(order):
    """Fills the order form with given details"""  
    page = browser.page()
    close_annoying_modal()
    radio_id = str("#id-body-" + str(order["Body"]))
    page.select_option("#head",index=int(order["Head"]))
    page.click(radio_id)
    page.get_by_placeholder("Enter the part number for the legs").fill(order["Legs"])
    page.fill("#address", order["Address"])
    page.click("#preview")
    page.click("#order")

    while not page.query_selector("#order-another"):
        page.click("#order")
    store_receipt_as_pdf(order["Order number"])
    page.click("#order-another")

def store_receipt_as_pdf(order_number):
    """Stores receipts to PDF and takes a screenshot of ordered robots"""
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf_file = f"output/receipt/{order_number}.pdf"
    screenshot = f"output/receipt/{order_number}.png"
    pdf.html_to_pdf(receipt_html, pdf_file)
    screenshot_robot(order_number)
    embed_screenshot_to_receipt(screenshot, pdf_file)
    os.remove(screenshot)
    



def screenshot_robot(order_number):
    """Screenshot Robot"""
    page = browser.page()
    screenshot = f"output/receipt/{order_number}.png"
    robo = page.query_selector("#robot-preview-image")
    robo.screenshot(path=screenshot)

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Embed screenshot to pdf"""
    pdf = PDF()
    pdf.add_files_to_pdf(files=[screenshot],target_document=pdf_file, append=True)

def archive_receipts():
    archive = Archive()
    archive.archive_folder_with_zip(folder="output/receipt/", archive_name="output/receipts.zip", include="*.pdf")