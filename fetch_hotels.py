import time, os
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv

load_dotenv('c:\\Users\\ATT-CODING2\\Desktop\\ram_github\\.env')

playwright = sync_playwright().start()
browser = playwright.chromium.launch(headless=True)
page = browser.new_page()

page.goto(os.environ['WEBSITE_URL'])
page.wait_for_load_state('networkidle')
time.sleep(2)
page.click('input[name="username"]')
page.fill('input[name="username"]', os.environ['BOT_USERNAME'])
page.click('input[name="password"]')
page.keyboard.type(os.environ['BOT_PASSWORD'])
page.click('button:has-text("Sign in")')
page.wait_for_load_state('networkidle')
time.sleep(5)

page.locator('section').get_by_text('Availability Consolidated').click()
page.wait_for_load_state('networkidle')
time.sleep(3)

page.locator('.dropdown-toggle').click()
time.sleep(2)

options = page.locator('a.dropdown-item, li.dropdown-item, [role="option"]').all()
print('Found options:', len(options))
for opt in options:
    print(opt.inner_text().strip())

browser.close()
playwright.stop()
