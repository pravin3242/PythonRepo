from datetime import datetime

import pandas as pd
from playwright.sync_api import sync_playwright

# Read configuration from Excel file
df = pd.read_excel("C:/Users/lnv0165/Downloads/Input123.xlsx", sheet_name='Configuration')
username = df.loc[0, 'Username']
password = df.loc[0, 'Password']
keywd = df.loc[0, 'Keyword']

# Define column names for DataFrame
column_names = ["Job Summary", "Description", "Skills", "About Client", "URL"]

# Create an empty DataFrame to store the scraped data
result_df = pd.DataFrame(columns=column_names)


def scrape_job_data(page):
    # Extract job data
    href = page.locator("//div//a[@data-test='slider-open-in-new-window UpLink']").get_attribute('href')
    URL = "https://www.upwork.com" + href
    title = page.locator("//h4[@class='d-flex align-items-center mt-0 mb-5']").text_content()
    description = page.locator("//div[@data-test='Description']").text_content()
    skills = page.locator("//section[@data-test='Expertise']").text_content()
    about_client = page.locator(
        "//div[@data-test='about-client-container AboutClientUserShared AboutClientUser']").text_content()

    # Append data to DataFrame
    result_df.loc[len(result_df)] = [title, description, skills, about_client, URL]
    print(result_df)


def run(playwright):
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://www.upwork.com/ab/account-security/login")
    page.get_by_placeholder("Username or Email").click()
    page.get_by_placeholder("Username or Email").fill(username)
    page.get_by_role("button", name="Continue", exact=True).click()
    page.get_by_role("textbox", name="Password").fill(password)
    page.get_by_role("button", name="Log in").click()
    page.get_by_role("dialog").get_by_role("button", name="Close", exact=True).click()
    page.get_by_label("Close").click()
    page.get_by_placeholder("Search for jobs").click()
    page.get_by_placeholder("Search for jobs").fill(keywd)
    page.get_by_placeholder("Search for jobs").press("Enter")
   # page.get_by_text("Sort by: Newest").click()
    page.wait_for_timeout(3000)
    page.click("//div[@data-test='jobs_per_page UpCDropdown']")
    page.wait_for_timeout(3000)
    page.keyboard.press('ArrowDown')
    page.keyboard.press('ArrowDown')
    page.keyboard.press('Enter')
    page.wait_for_timeout(5000)

    # Extract job data
    elements = page.locator("//article//div//small//span[contains(text(),'minutes ago')] | //article//div//small//span[contains(text(),'hours ago')] | //article//div//small//span[contains(text(),'hour ago')]").element_handles()

    for element in elements:
        element.click()
        scrape_job_data(page)
        page.keyboard.press('Escape')  # Close the job detail popup

    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)

current_time = datetime.now()
sheet_name = current_time.strftime('Sheet_%Y-%m-%d_%H-%M-%S')
print(result_df)
# Write DataFrame to Excel file
# Path to your Excel file
file_path = "C:/Users/lnv0165/Downloads/Input123.xlsx"

# Attempt to write to a new sheet in the existing workbook
try:
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='new') as writer:
        result_df.to_excel(writer, sheet_name=sheet_name, index=False)
except FileNotFoundError:
    # If the file does not exist, create it with the DataFrame
    result_df.to_excel(file_path, sheet_name=sheet_name, index=False)
