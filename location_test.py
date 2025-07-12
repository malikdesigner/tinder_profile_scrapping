import sys
import asyncio
from playwright.async_api import async_playwright, Playwright, Page
from base_morelogin.base_func import requestHeader, postRequest
import traceback
import time
import random
import json
import os
import urllib.parse
import re
import requests
from datetime import datetime
import pandas as pd
import logging

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('tinder_scraper.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Load configuration
with open("config.json", "r") as config_file:
    config = json.load(config_file)

APPID = config.get("APPID")
SECRETKEY = config.get("SECRETKEY")
BASEURL = config.get("BASEURL")
envId = config.get("envId")
uniqueId = config.get("uniqueId")

# Cities for geographic diversity
CITIES = [
    "New York, NY", "Los Angeles, CA", "Chicago, IL", "Houston, TX", "Phoenix, AZ",
    "Philadelphia, PA", "San Antonio, TX", "San Diego, CA", "Dallas, TX", "San Jose, CA",
    "Austin, TX", "Jacksonville, FL", "Fort Worth, TX", "Columbus, OH", "Charlotte, NC",
    "San Francisco, CA", "Indianapolis, IN", "Seattle, WA", "Denver, CO", "Washington, DC",
    "Boston, MA", "El Paso, TX", "Detroit, MI", "Nashville, TN", "Portland, OR",
    "Memphis, TN", "Oklahoma City, OK", "Las Vegas, NV", "Louisville, KY", "Baltimore, MD"
]

class TinderScraper:
    def __init__(self, profiles_per_city: int = 5):
        self.profiles_per_city = profiles_per_city
        # self.scraped_data = []
        # self.base_folder = "tinder_profiles"
        # self.excel_file = os.path.join(self.base_folder, f"tinder_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        # self.progress_file = os.path.join(self.base_folder, "scraping_progress.json")
        
        # # Initialize progress tracking
        # self.load_progress()
        # self.setup_folders()
        
        # logger.info(f"TinderScraper initialized - Starting from profile {self.profile_count + 1}")
        # logger.info(f"Current city: {self.get_current_city()}")
        



    async def test_change_location_with_arrow_key(page):
        city = "Los Angeles, CA"

        async with async_playwright() as p:
           
            try:
                print("Navigating to Tinder passport settings...")
                await page.goto("https://tinder.com/app/settings/plus/passport", wait_until="domcontentloaded", timeout=30000)
                await asyncio.sleep(5)

                print("Focusing the search input...")
                search_input = await page.wait_for_selector('input[aria-label="Search a Location"]', timeout=10000)
                await search_input.click()
                await search_input.fill("")  # Clear any existing text
                await asyncio.sleep(1)

                print(f"Typing city: {city}")
                await search_input.type(city, delay=100)
                await asyncio.sleep(3)  # Wait for dropdown to populate

                print("Sending ArrowDown key to select first suggestion...")
                await search_input.press("ArrowDown")
                await asyncio.sleep(1)

                print("Pressing Enter to confirm...")
                await search_input.press("Enter")
                await asyncio.sleep(5)

                print("✅ City selection completed successfully.")

            except Exception as e:
                print(f"❌ Error occurred: {e}")
            finally:
                await browser.close()
    async def scrape_profiles(self, page, total_profiles: int = 100):
        """Main scraping function"""
        try:
            # Navigate to Tinder
            success = await self.test_change_location_with_arrow_key(page)
            
        except Exception as e:
            logger.error(f"Critical error in main scraping function: {e}")
            traceback.print_exc()

    async def navigate_to_next_profile(self, page):
        """Navigate to next profile (swipe right or use next button)"""
        try:
            # Try to find and click pass/like buttons
            pass_buttons = await page.query_selector_all('.gamepad-button-wrapper button')
            if pass_buttons and len(pass_buttons) > 0:
                # Randomly choose to pass or like (use index 0 or 2 if available)
                button_index = random.choice([0, min(2, len(pass_buttons) - 1)])
                button = pass_buttons[2]
                print(f"next button : {button_index}")
                await button.click()
                await asyncio.sleep(2)
            else:
                # Try keyboard navigation
                await page.keyboard.press('ArrowRight')
                await asyncio.sleep(2)
                
        except Exception as e:
            logger.error(f"Error navigating to next profile: {e}")
            # Fallback to keyboard
            try:
                await page.keyboard.press('ArrowRight')
                await asyncio.sleep(2)
            except:
                pass

async def operationEnv(cdpUrl, playwright, total_profiles):
    """Main operation function"""
    try:
        logger.info(f"Connecting to browser at: {cdpUrl}")
        browser = await playwright.chromium.connect_over_cdp(cdpUrl)
        defaultContext = browser.contexts[0]
        
        # Create new page
        page = await defaultContext.new_page()
        
        # Initialize scraper
        scraper = TinderScraper()
        
        # Start scraping
        await scraper.scrape_profiles(page, total_profiles)
        
        return {
            'success': True, 
            'profiles_scraped': scraper.profile_count,
            'excel_file': scraper.excel_file
        }
        
    except Exception as e:
        logger.error(f"Error in operationEnv: {e}")
        traceback.print_exc()
        return {'error': str(e)}

async def startEnv(envId, uniqueId, appId, secretKey, baseUrl):
    """Start MoreLogin environment"""
    try:
        requestPath = baseUrl + '/api/env/start'
        data = {'envId': str(envId), 'uniqueId': str(uniqueId)}
        headers = requestHeader(appId, secretKey)

        response = postRequest(requestPath, data, headers)
        logger.info("Environment start response:", response.text)

        if response.status_code != 200:
            logger.error(f"Error: {response.status_code}, {response.text}")
            return None

        response_json = response.json()
        if response_json['code'] != 0:
            logger.error(response_json['msg'])
            return None

        port = response_json['data']['debugPort']
        cdpUrl = f'http://127.0.0.1:{port}'
        return cdpUrl
    except Exception as e:
        logger.error(f"Error starting environment: {e}")
        return None

async def stopEnv(envId, uniqueId, appId, secretKey, baseUrl):
    """Stop MoreLogin environment"""
    try:
        requestPath = '/api/env/close'
        data = {'envId': envId, 'uniqueId': uniqueId}
        headers = requestHeader(appId, secretKey)
        response = postRequest(baseUrl + requestPath, data, headers).json()
        if response['code'] == -1:
            return False
        return True
    except Exception as e:
        logger.error(f"Error stopping environment: {e}")
        return False

async def main(envId: str, uniqueId: int, total_profiles: int = 100):
    """Main function"""
    async with async_playwright() as playwright:
        return await run(playwright, envId, uniqueId, total_profiles)

async def run(playwright: Playwright, envId: str, uniqueId: int, total_profiles: int):
    """Run the scraper"""
    try:
        logger.info("Starting MoreLogin environment...")
        cdpUrl = await startEnv(envId, uniqueId, APPID, SECRETKEY, BASEURL)
        
        if not cdpUrl:
            logger.error("Failed to start environment")
            return {'error': 'Failed to start environment'}

        logger.info(f"Environment started at: {cdpUrl}")
        
        # Run the scraper
        result = await operationEnv(cdpUrl, playwright, total_profiles)
        return result

    except Exception as e:
        logger.error(f"Error in run: {e}")
        traceback.print_exc()
        return {'error': str(e)}

    finally:
        # Stop environment (uncomment if needed)
        # await stopEnv(envId, uniqueId, APPID, SECRETKEY, BASEURL)
        logger.info('Tinder scraping automation completed.')

def parse_arguments(args):
    """Parse command line arguments"""
    arguments = {}
    for arg in args[1:]:
        if "=" in arg:
            key, value = arg.split("=", 1)
            arguments[key] = value
    return arguments

if __name__ == "__main__":
    args = parse_arguments(sys.argv)
    
    # Get total profiles to scrape (default: 100)
    total_profiles = int(args.get("total_profiles", 100))
    
    print("="*60)
    print("TINDER PROFILE SCRAPER - ENHANCED VERSION")
    print("="*60)
    print(f"Total profiles to scrape: {total_profiles}")
    print("Features:")
    print("- Resume functionality with progress tracking")
    print("- Profile-by-profile Excel saving")
    print("- Geographic diversity (switching cities every 5 profiles)")
    print("- Images named with person's name: profileId_personName_imageId")
    print("- Organized folder structure by city")
    print("- Local image paths saved in Excel")
    print("- Crash recovery support")
    print("="*60)
    
    try:
        result = asyncio.run(main(envId, uniqueId, total_profiles))
        
        print("\n" + "="*60)
        print("SCRAPING COMPLETED!")
        print("="*60)
        
        if isinstance(result, dict):
            if 'error' in result:
                print(f"❌ Error occurred: {result['error']}")
            elif 'success' in result:
                print(f"✅ Successfully scraped {result['profiles_scraped']} profiles")
                print(f"📊 Excel file saved: {result['excel_file']}")
                print(f"📁 Images organized in city folders under 'tinder_profiles/'")
                print(f"📄 Progress file: tinder_profiles/scraping_progress.json")
                print(f"🖼️  Images named as: profileId_personName_imageId.jpg")
        
        print("="*60)
        
    except Exception as e:
        logger.error(f"Script failed with error: {e}")
        traceback.print_exc()