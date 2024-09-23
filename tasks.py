from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
import os

@task
def initiate_robot_order_process():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    prepare_output_directories()
    access_robot_order_page()
    download_order_data()
    handle_robot_orders()
    create_zip_archives()

def prepare_output_directories():
    """Creates necessary output directories for receipts and screenshots."""
    os.makedirs("C:/Robocorp Projects/Leval2 doc/receipts", exist_ok=True)
    os.makedirs("C:/Robocorp Projects/Leval2 doc/screenshots", exist_ok=True)

def access_robot_order_page():
    """Navigates to the robot order website and accepts the pop-up."""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    page = browser.page()
    page.click('text=OK')

def download_order_data():
    """Downloads the orders CSV file from the given URL."""
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv", overwrite=True)

def handle_robot_orders():
    """Processes each robot order from the downloaded CSV."""
    csv_handler = Tables()
    robot_orders = csv_handler.read_table_from_csv("orders.csv")
    for order in robot_orders:
        submit_robot_order(order)

def submit_robot_order(order_data):
    """Fills in robot order details and submits the order."""
    page = browser.page()
    head_options = {
        "1": "Roll-a-thor head",
        "2": "Peanut crusher head",
        "3": "D.A.V.E head",
        "4": "Andy Roid head",
        "5": "Spanner mate head",
        "6": "Drillbit 2000 head"
    }
    
    head_choice = order_data["Head"]
    page.select_option("#head", head_options.get(head_choice))
    page.click(f'//*[@id="root"]/div/div[1]/div/div[1]/form/div[2]/div/div[{order_data["Body"]}]/label')
    page.fill("input[placeholder='Enter the part number for the legs']", order_data["Legs"])
    page.fill("#address", order_data["Address"])

    while True:
        page.click("#order")
        if page.query_selector("#order-another"):
            pdf_path = save_order_receipt_as_pdf(order_data["Order number"])
            screenshot_path = capture_robot_image(order_data["Order number"])
            insert_image_into_receipt(screenshot_path, pdf_path)
            request_new_order()
            confirm_order()
            break

def request_new_order():
    """Requests to order another robot."""
    page = browser.page()
    page.click("#order-another")

def confirm_order():
    """Confirms the order by clicking OK."""
    page = browser.page()
    page.click('text=OK')

def save_order_receipt_as_pdf(order_id):
    """Saves the robot order receipt as a PDF file."""
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf_path = f"C:/Robocorp Projects/Leval2 doc/receipts/{order_id}.pdf"
    
    try:
        pdf.html_to_pdf(receipt_html, pdf_path)
        print(f"Receipt PDF created at {pdf_path}")
    except Exception as e:
        print(f"Error creating PDF for order {order_id}: {e}")
    
    return pdf_path

def capture_robot_image(order_id):
    """Captures a screenshot of the ordered robot."""
    page = browser.page()
    screenshot_path = f"C:/Robocorp Projects/Leval2 doc/screenshots/{order_id}.png"
    
    try:
        page.locator("#robot-preview-image").screenshot(path=screenshot_path)
        print(f"Screenshot saved at {screenshot_path}")
    except Exception as e:
        print(f"Error taking screenshot for order {order_id}: {e}")
    
    return screenshot_path

def insert_image_into_receipt(image_path, pdf_path):
    """Embeds the robot image into the PDF receipt."""
    pdf = PDF()
    
    try:
        pdf.add_watermark_image_to_pdf(image_path=image_path, 
                                        source_path=pdf_path, 
                                        output_path=pdf_path)
        print(f"Image embedded into PDF at {pdf_path}")
    except Exception as e:
        print(f"Error embedding image into PDF for {pdf_path}: {e}")

def create_zip_archives():
    """Creates ZIP archives for receipts and screenshots."""
    archive_tool = Archive()
    
    try:
        receipts_zip_path = "C:/Robocorp Projects/Leval2 doc/receipts.zip"
        screenshots_zip_path = "C:/Robocorp Projects/Leval2 doc/screenshots.zip"
        
        # Archive receipts
        archive_tool.archive_folder_with_zip("C:/Robocorp Projects/Leval2 doc/receipts", receipts_zip_path)
        print(f"Receipts archived to {receipts_zip_path}")
        
        # Archive screenshots
        archive_tool.archive_folder_with_zip("C:/Robocorp Projects/Leval2 doc/screenshots", screenshots_zip_path)
        print(f"Screenshots archived to {screenshots_zip_path}")
    
    except Exception as e:
        print(f"Error during archiving: {e}")
