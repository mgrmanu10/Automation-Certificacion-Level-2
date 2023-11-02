from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA import FileSystem


http = HTTP()
csv = Tables()
pdf = PDF()
rpa_browser = Selenium()
lib = Archive()
file_sys = FileSystem.FileSystem()

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """    
    establish_directories()
    open_robot_order_website();
    download_csv_file();
    get_orders();
    archive_receipts();
    

def open_robot_order_website():
    """Navigates to the given URL"""
    rpa_browser.open_available_browser("https://robotsparebinindustries.com/#/robot-order");
    rpa_browser.maximize_browser_window()
    
    
    
def accept_cookies_terms():
    try:
        rpa_browser.wait_until_element_is_visible('//*[@id="root"]/div/div[2]/div/div/div/div/div/button[1]', 20)
        rpa_browser.click_element('//*[@id="root"]/div/div[2]/div/div/div/div/div/button[1]')
    except Exception as e:
        try:
            rpa_browser.click_element("//button[@id='order-another']")
        except Exception as e:
            pass
        accept_cookies_terms()
    
def download_csv_file():
    """Downloads csv file from the given URL"""
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True);
    
def fill_forms(order):
    accept_cookies_terms()
    
    rpa_browser.wait_until_element_is_enabled(f'//select[@id="head"]/option[@value="{order["Head"]}"]', 100)
    rpa_browser.click_element(f'//select[@id="head"]/option[@value="{order["Head"]}"]')
    rpa_browser.click_element(f'//input[@id="id-body-{order["Body"]}"]')
    rpa_browser.input_text("//*[@id='address']", str(order["Address"]))
    rpa_browser.input_text("//*[@placeholder='Enter the part number for the legs']",order["Legs"])
    rpa_browser.click_element("//button[@id='preview']")
    rpa_browser.click_element("//button[@id='order']")
    
    try:
        rpa_browser.wait_until_element_is_visible("//*[@id='receipt']")
    except Exception as e:
        rpa_browser.click_element("//button[@id='order']")
    store_receipt_as_pdf(order["Order number"])
    

def get_orders(): 
    orders_csv = csv.read_table_from_csv("orders.csv", header=True);
    
    for row in orders_csv:
        fill_forms(row);
        
        
def store_receipt_as_pdf(order_number):
    
    try:
        rpa_browser.wait_until_element_is_visible("//*[@id='receipt']")
        html = rpa_browser.find_element("//*[@id='receipt']").get_attribute("innerHTML")
        rpa_browser.screenshot("//div[@id='robot-preview-image']", f"output/screenshots/{order_number}.png")
        pdf.html_to_pdf(html, f"output/receipts/{order_number}.pdf")
        embed_screenshot_to_receipt(f"output/screenshots/{order_number}.png", f"output/receipts/{order_number}.pdf")
        rpa_browser.wait_until_element_is_visible("//button[@id='order-another']")
        rpa_browser.click_element("//button[@id='order-another']")
    except Exception as e:
        rpa_browser.wait_until_element_is_visible("//button[@id='order']")
        rpa_browser.click_element("//button[@id='order']")
    
    
def embed_screenshot_to_receipt(screenshot, pdf_file): 
    pdf.add_files_to_pdf(
        files=[screenshot],
        target_document=pdf_file,
        append=True
    )
    
def archive_receipts():
    lib.archive_folder_with_zip('./output/receipts', 'output/receipts.zip', recursive=True)

def establish_directories():
    dir_exist = FileSystem.Path('output').is_dir()
    if dir_exist:
        FileSystem.FileSystem().remove_directory('output', True)
    FileSystem.Path('output/screenshots').mkdir(parents=True, exist_ok=True)
    FileSystem.Path('output/receipts').mkdir(parents=True, exist_ok=True)
    file_sys.create_file("output/output.robolog")