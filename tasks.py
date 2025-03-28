import time
import os
from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_robot_SpareBin():
    '''
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    '''
    browser.configure(
        slowmo=100,
    )
    open_robot_order_website()
    orders = get_orders()
    for order in orders:
        close_annoying_modal()
        fill_the_form(order)
        store_receipt_as_pdf(order)
        screenshot_robot(order)
        order_next_robot()
    archive_receipts()

def close_annoying_modal():
    page = browser.page()
    page.click('button:text("OK")')


def open_robot_order_website():
    '''
    Opens the robot order website
    '''
    browser.goto('https://robotsparebinindustries.com/#/robot-order')
    get_orders()

def get_orders():
    http = HTTP()
    http.download('https://robotsparebinindustries.com/orders.csv', overwrite=True)
    library = Tables()
    orders = library.read_table_from_csv("orders.csv", columns=['Order number', 'Head', 'Body', 'Legs', 'Address'])
    return orders

def fill_the_form(order):
    page = browser.page()
    page.select_option("#head", str(order["Head"]))
    page.click(f"#id-body-{order['Body']}")
    page.fill("input[placeholder='Enter the part number for the legs']", str(order["Legs"]))
    page.fill('#address', order['Address'])
    page.click('button:text("Preview")')
    submit_order_with_retry()

def submit_order_with_retry(max_attempts=5):
    """Submit the order with retry logic for failures"""
    page = browser.page()
    attempts = 0
    
    while attempts < max_attempts:
        page.click("#order")
        
        # Check if order was successful by looking for receipt element
        receipt = page.query_selector("#receipt")
        if receipt:
            return receipt
        
        # Check for error message
        error = page.query_selector(".alert-danger")
        if error:
            print(f"Order submission failed, attempt {attempts + 1}/{max_attempts}")
            attempts += 1
            time.sleep(1)  # Brief pause before retrying
        else:
            # No error but no receipt either - might need different handling
            attempts += 1
            time.sleep(1)
    raise Exception(f"Failed to submit order after {max_attempts} attempts")

def store_receipt_as_pdf(order_number):
    os.makedirs("output/receipts", exist_ok=True) # create directory if it doesn't exist
    page = browser.page()
    receipt = page.query_selector("#receipt") # get the receipt element
    pdf = PDF()
    receipt_html = receipt.inner_html() # get the HTML content of the receipt
    pdf.html_to_pdf(receipt_html, f"output/receipts/receipt_{order_number}.pdf")
    
def screenshot_robot(order_number):
    os.makedirs("output/images", exist_ok=True) # create directory if it doesn't exist
    page = browser.page()
    screenshot_path = f"output/screenshots/robot_{order_number}.png"
    robot_preview = page.query_selector("#robot-preview-image")
    if robot_preview:
        robot_preview.screenshot(path=screenshot_path)
    else:
        page.screenshot(path=screenshot_path)
    embed_screenshot_to_receipt(screenshot_path, f"output/receipts/receipt_{order_number}.pdf")


def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    pdf.add_files_to_pdf(files=[screenshot], target_document=pdf_file, append=True)

def order_next_robot():
    page = browser.page()
    page.click("#order-another")

def archive_receipts():
    os.makedirs("output/receipts", exist_ok=True) # create directory if it doesn't exist
    archive = Archive()
    archive.archive_folder_with_zip("output/receipts", "output/receipts.zip")