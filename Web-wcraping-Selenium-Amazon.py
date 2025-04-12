from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import time
import random
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
from bs4 import BeautifulSoup
import re
from plyer import notification


def send_notification(category):
    notification.notify(
        title=f"{category} Scraping Completed",
        message=f"Scraping for {category} is completed!",
        timeout=10
    )


def send_notification_mistake(category):
    notification.notify(
        title=f"Error in {category}",
        message=f"There was an error with the scraping for {category}!",
        timeout=10
    )


def parse_sales_text(sales_text):
    sales_text = sales_text.strip()
    match = re.search(r"(\d+(?:\.\d+)?)([KM]?)\+?", sales_text, re.IGNORECASE)
    if match:
        number = float(match.group(1))
        unit = match.group(2).upper()
        if unit == 'K':
            return int(number * 1_000)
        elif unit == 'M':
            return int(number * 1_000_000)
        else:
            return int(number)
    return None


def parse_amazon_html_to_xlsx(html_file, category):
    with open(html_file, "r", encoding="utf-8") as f:
        html_content = f.read()
    soup = BeautifulSoup(html_content, "html.parser")
    products = soup.find_all("div", class_="s-result-item")
    data = []
    for product in products:
        title_tag = product.find("h2")
        title = title_tag.get_text(strip=True) if title_tag else "No Title"
        price_whole_tag = product.find("span", class_="a-price-whole")
        price_fraction_tag = product.find("span", class_="a-price-fraction")
        price_whole = price_whole_tag.get_text(strip=True) if price_whole_tag else "0"
        price_fraction = price_fraction_tag.get_text(strip=True) if price_fraction_tag else "00"
        price = f"{price_whole}{price_fraction}" if price_whole_tag else "No Price"
        rating_tag = product.find("span", class_="a-icon-alt")
        rating = rating_tag.get_text(strip=True) if rating_tag else "No Rating"
        review_count_tag = product.find("span", class_="a-size-base s-underline-text")
        review_count = review_count_tag.get_text(strip=True) if review_count_tag else "No Reviews"
        sales_tag = product.find("span", class_="a-size-base a-color-secondary")
        if sales_tag:
            sales_text = sales_tag.get_text()
            sales = parse_sales_text(sales_text)
            if sales is None:
                sales = "No Sales Data"
        else:
            sales = "No Sales Data"

        link_tag = product.find("a", class_="a-link-normal s-line-clamp-2 s-link-style a-text-normal")
        link = "https://www.amazon.com" + link_tag["href"] if link_tag and link_tag.has_attr("href") else "No Link"
        data.append([title, price, rating, review_count, sales, link])

    # Remove rows with no price
    data = [row for row in data if row[1] != "No Price"]
    df = pd.DataFrame(data, columns=["Product Name", "Price", "Rating", "Review Count", "Sales", "Product Link"])
    remove_keywords = ["Results", "Current Trends", "More Results", "Related Searches", "Need Help?", "Top Rated", "No Title"]
    df_cleaned = df[~df["Product Name"].astype(str).str.contains('|'.join(remove_keywords), na=False)]
    output_file = f"amazon_{category.replace(' ', '_')}.xlsx"
    df_cleaned.to_excel(output_file, index=False, engine="openpyxl")
    print(f" Cleaning complete, saved as: {output_file}")
    send_notification(category)


def find_next_page_button(driver):
    possible_classes = [
        "s-pagination-next",
        "s-pagination-button",
        "s-pagination-item s-pagination-next s-pagination-button",
        "s-pagination-button-accessibility"
    ]
    for class_name in possible_classes:
        try:
            next_page_btn = driver.find_element(By.CLASS_NAME, class_name)
            if next_page_btn.is_displayed() and next_page_btn.is_enabled():
                return next_page_btn
        except NoSuchElementException:
            continue
    return None


def create_driver(edge_profile_path):
    options = Options()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument(f"--user-data-dir={edge_profile_path}")
    options.add_argument("--profile-directory=Default")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920x1080")
    options.add_argument("--incognito")
    options.add_argument("--log-level=3")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36")
    options.add_argument("referer=https://www.google.com/")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    service = Service(EdgeChromiumDriverManager().install())
    driver = webdriver.Edge(service=service, options=options)
    return driver


# Define the categories and corresponding URLs (can be expanded as needed)
categories = {
    "Watch": "https://www.amazon.com/s?k=Watch&crid=2TCM5XFJ9A7QB&sprefix=watch%2Caps%2C548&ref=nb_sb_noss_1",
    "Toilet paper": "https://www.amazon.com/s?k=toilet+paper&crid=30ALVAWJICYWE&sprefix=Toilet+Paper%2Caps%2C534&ref=nb_sb_ss_mvt-t9-ranker_ci_hl-bn-left_1_12",
}

edge_profile_path = "C:\\Users\\123\\AppData\\Local\\Microsoft\\Edge\\User Data\\acc1345"
start_time = time.time()

# Scraping each category sequentially
for category, url in categories.items():
    print(f" Starting scraping for category: {category}")
    driver = create_driver(edge_profile_path)
    driver.get(url)
    time.sleep(random.uniform(5, 8))  # Wait for page to load

    NUMBER = 1
    html_file = f"amazon_search_{category.replace(' ', '_')}.html"

    # Save the first page HTML
    with open(html_file, "w", encoding="utf-8") as file:
        file.write(driver.page_source)
    print(f" First page saved: {html_file}")

    # Main scraping loop: page by page
    while True:
        print(f" Processing page {NUMBER}...")
        next_page_visible = False
        while not next_page_visible:
            scroll_distance = random.randint(300, 700)
            try:
                driver.execute_script(f"window.scrollBy(0, {scroll_distance});")
            except Exception as e:
                print(f"\n Error with execute_script: {e}")
                send_notification_mistake(category)
                input(" Timeout or other error occurred, press ENTER to continue after checking...")
                continue  # Retry scrolling

            next_page_btn = find_next_page_button(driver)
            if next_page_btn:
                print(" Found 'Next Page' button, stopping scrolling!")
                next_page_visible = True
            else:
                if NUMBER == 1:
                    print(" 'Next Page' button not found on first page, might need manual operation!")
                    send_notification_mistake(category)
                    input(" Please check the page or manually press the button, then press ENTER to continue...")
                    continue  # Recheck for next page
                else:
                    print(" 'Next Page' button not found, likely the last page!")
                    break
            # Simulate random mouse movement
            x = random.randint(100, 1200)
            y = random.randint(100, 800)
            driver.execute_script(f"document.elementFromPoint({x}, {y}).scrollIntoView();")
            time.sleep(random.uniform(1.5, 3))

        print(" Scrolling complete!")
        with open(html_file, "a", encoding="utf-8") as file:
            file.write("\n<!-- Next page starts -->\n")
            file.write(driver.page_source)
        print(f" Added to file: {html_file}")
        time.sleep(3)

        try:
            next_page_btn = find_next_page_button(driver)
            if not next_page_btn:
                print(" No 'Next Page' button found, scraping ends!")
                break
            old_url = driver.current_url
            print(" Preparing to click 'Next Page' button...")
            ActionChains(driver).move_to_element(next_page_btn).click().perform()
            time.sleep(random.uniform(5, 8))
            if driver.current_url == old_url:
                print(" Failed to navigate to next page, possibly the last page!")
                break
            NUMBER += 1
            print(f" Successfully navigated to page {NUMBER}!")
        except Exception as e:
            print(f"\n Error while clicking 'Next Page': {e}")
            input(" Timeout or error occurred, please check the page and press ENTER to continue...")
            send_notification_mistake(category)
            continue  # Let the program continue after manual handling

    print(f" All pages for category {category} scraped, {NUMBER} pages in total")
    driver.quit()
    parse_amazon_html_to_xlsx(html_file, category)
    end_time = time.time()
    elapsed_time = (end_time - start_time) / 60
    print(f" Elapsed time: {elapsed_time:.2f} minutes")

    # Wait 60 seconds before starting the next category
    print(" Waiting 60 seconds before starting the next category...")
    time.sleep(60)

end_time = time.time()
elapsed_time = (end_time - start_time) / 60
print(f" Total elapsed time: {elapsed_time:.2f} minutes")
