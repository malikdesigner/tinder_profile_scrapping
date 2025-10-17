# updated script
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

# ADD THESE IMPORTS AT THE TOP OF YOUR FILE (after the existing imports)
import re
import unicodedata
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
def clean_text_for_excel(text: str) -> str:
    """
    Clean text by removing emojis, special characters, and other problematic characters
    that can cause issues when saving to Excel files.
    """
    if not text or not isinstance(text, str):
        return ""
    
    # Remove emojis using regex pattern
    emoji_pattern = re.compile(
        "["
        "\U0001F600-\U0001F64F"  # emoticons
        "\U0001F300-\U0001F5FF"  # symbols & pictographs
        "\U0001F680-\U0001F6FF"  # transport & map symbols
        "\U0001F1E0-\U0001F1FF"  # flags (iOS)
        "\U00002702-\U000027B0"  # dingbats
        "\U000024C2-\U0001F251"  # enclosed characters
        "\U0001F900-\U0001F9FF"  # supplemental symbols and pictographs
        "\U0001FA70-\U0001FAFF"  # symbols and pictographs extended-a
        "\U00002600-\U000026FF"  # miscellaneous symbols
        "\U00002700-\U000027BF"  # dingbats
        "]+",
        flags=re.UNICODE
    )
    
    # Remove emojis
    text = emoji_pattern.sub('', text)
    
    # Remove other problematic characters
    replacements = {
        '\u2019': "'",  # Right single quotation mark
        '\u2018': "'",  # Left single quotation mark
        '\u201c': '"',  # Left double quotation mark
        '\u201d': '"',  # Right double quotation mark
        '\u2013': '-',  # En dash
        '\u2014': '-',  # Em dash
        '\u2026': '...',  # Horizontal ellipsis
        '\u00a0': ' ',  # Non-breaking space
        '\u00ad': '',   # Soft hyphen
        '\u200b': '',   # Zero-width space
        '\u200c': '',   # Zero-width non-joiner
        '\u200d': '',   # Zero-width joiner
        '\ufeff': '',   # Zero-width no-break space (BOM)
    }
    
    for old, new in replacements.items():
        text = text.replace(old, new)
    
    # Remove any remaining non-printable characters except newlines and tabs
    text = ''.join(char for char in text if unicodedata.category(char)[0] != 'C' or char in '\n\t')
    
    # Remove control characters but keep basic whitespace
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', text)
    
    # Clean up extra whitespace
    text = re.sub(r'\s+', ' ', text)
    text = text.strip()
    
    return text
def clean_name_only(text: str) -> str:
    """
    Clean name by removing ONLY emojis, but keeping all other Unicode characters
    including Chinese, Arabic, Cyrillic, accented characters, etc.
    This is specifically for names to preserve international characters.
    """
    if not text or not isinstance(text, str):
        return "Unknown"
    
    # Remove emojis using comprehensive regex pattern
    emoji_pattern = re.compile(
        "["
        "\U0001F600-\U0001F64F"  # emoticons
      
        "]+",
        flags=re.UNICODE
    )
    
    # Remove emojis only
    text = emoji_pattern.sub('', text)
    
    # Remove only control characters that are problematic, but keep all printable Unicode
    # Only remove ASCII control characters (0x00-0x1F and 0x7F-0x9F), but keep Unicode text
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', text)
    
    # Clean up extra whitespace
    text = re.sub(r'\s+', ' ', text)
    text = text.strip()
    
    # Return "Unknown" if text is empty after cleaning
    if not text or text.isspace():
        text = "Unknown"
    
    return text
def clean_filename_text(text: str) -> str:
    """
    Clean text specifically for use in filenames.
    """
    if not text or not isinstance(text, str):
        return "Unknown"
    
    # First apply general text cleaning
    text = clean_text_for_excel(text)
    
    # Remove/replace characters not allowed in filenames
    forbidden_chars = r'[<>:"|?*\\/]'
    text = re.sub(forbidden_chars, '', text)
    
    # Remove leading/trailing dots and spaces
    text = text.strip('. ')
    
    # Replace multiple spaces/underscores with single underscore
    text = re.sub(r'[\s_]+', '_', text)
    
    # Limit length to avoid filesystem issues
    if len(text) > 50:
        text = text[:50].rstrip('_')
    
    if not text or text.isspace():
        text = "Unknown"
    
    return text

def extract_city_names(csv_path: str) -> list:
    """
    Reads the worldcities_filtered.csv CSV and returns a list of unique city names.
    """
    try:
        # Use UTF-8 encoding explicitly
        with open(csv_path, encoding='utf-8') as f:
            first_line = f.readline()
            sep = '\t' if '\t' in first_line else ','

        # Read with the correct encoding
        df = pd.read_csv(csv_path, sep=sep, encoding='utf-8')

        if 'city' not in df.columns:
            raise ValueError("No 'city' column found in the file.")

        cities = df['city'].dropna().drop_duplicates().tolist()
        logger.info(f"Extracted {len(cities)} unique cities from CSV")
        return cities

    except Exception as e:
        logger.error(f"Error reading CSV: {e}")
        return []

def load_cities_from_csv_or_fallback(csv_file_path: str = "worldcities_filtered.csv", max_cities: int = 1000) -> list:
    """
    Load cities from CSV file or use fallback list if CSV is not available.
    """
    # Try to load from CSV first
    cities_from_csv = extract_city_names(csv_file_path)
    
    if cities_from_csv:
        # Take first 1000 cities or all if less than 1000
        selected_cities = cities_from_csv[:max_cities]
        logger.info(f"Using {len(selected_cities)} cities from CSV file")
        return selected_cities
    else:
        # Fallback to hardcoded cities
        fallback_cities = [
            "New York, NY", "Los Angeles, CA", "Chicago, IL", "Houston, TX", "Phoenix, AZ",
            "Philadelphia, PA", "San Antonio, TX", "San Diego, CA", "Dallas, TX", "San Jose, CA",
            "Austin, TX", "Jacksonville, FL", "Fort Worth, TX", "Columbus, OH", "Charlotte, NC",
            "San Francisco, CA", "Indianapolis, IN", "Seattle, WA", "Denver, CO", "Washington, DC",
            "Boston, MA", "El Paso, TX", "Detroit, MI", "Nashville, TN", "Portland, OR",
            "Memphis, TN", "Oklahoma City, OK", "Las Vegas, NV", "Louisville, KY", "Baltimore, MD"
        ]
        logger.warning("CSV file not found or invalid, using fallback cities list")
        return fallback_cities

# Load cities from CSV or use fallback
CITIES = load_cities_from_csv_or_fallback("worldcities_filtered.csv", 1000)

class TinderScraper:
    def __init__(self, profiles_per_city: int = 5):
        self.profiles_per_city = profiles_per_city
        self.scraped_data = []
        self.base_folder = "tinder_profiles"
        self.excel_file = os.path.join(self.base_folder, f"tinder_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        self.progress_file = os.path.join(self.base_folder, "scraping_progress.json")
        self.cities_progress_file = os.path.join(self.base_folder, "cities_progress.json")
        
        # Initialize progress tracking
        self.load_progress()
        self.load_cities_progress()
        self.setup_folders()
        
        logger.info(f"TinderScraper initialized - Starting from profile {self.profile_count + 1}")
        logger.info(f"Current city: {self.get_current_city()}")
        logger.info(f"Total cities available: {len(CITIES)}")
        

    
    async def handle_modal_popup(self, page):
        """
        Handle modal popups that might appear when page loads.
        This function safely attempts to close modals without crashing if none are found.
        """
        try:
            logger.debug("Checking for modal popups...")
            
            # Wait a moment for any modals to appear
            await asyncio.sleep(1)
            
            # List of common modal close button selectors and text patterns
            close_selectors = [
                # Common close button selectors
                '[aria-label="Close"]',
                '[data-testid="close"]',
                '[data-testid="modal-close"]',
                '.close-button',
                '.modal-close',
                'button[aria-label*="close" i]',
                'button[title*="close" i]',
                '[role="button"][aria-label*="close" i]',
                # X button patterns
                'button:has-text("×")',
                'button:has-text("✕")',
                'span:has-text("×")',
                'span:has-text("✕")',
                # Overlay/backdrop areas (usually clickable to close)
                '.modal-backdrop',
                '.overlay',
                '[data-testid="modal-backdrop"]'
            ]
            
            # Text-based close patterns (case insensitive)
            close_texts = [
                "Maybe later",
                "No thanks",
                "Accept",
                "Agree"
            ]
            
            modal_found = False
            
            # Try selector-based approaches first
            for selector in close_selectors:
                try:
                    element = await page.query_selector(selector)
                    if element:
                        # Check if element is visible
                        is_visible = await element.is_visible()
                        if is_visible:
                            logger.info(f"Found modal close element with selector: {selector}")
                            await element.click()
                            modal_found = True
                            await asyncio.sleep(1)  # Wait for modal to close
                            break
                except Exception as e:
                    # Continue to next selector if this one fails
                    logger.debug(f"Selector {selector} failed: {e}")
                    continue
            
            # If no selector worked, try text-based approach
            if not modal_found:
                for text in close_texts:
                    try:
                        # Try different button patterns with the text
                        text_selectors = [
                            f'button:has-text("{text}")',
                            f'[role="button"]:has-text("{text}")',
                            f'div:has-text("{text}")',
                            f'span:has-text("{text}")',
                            f'a:has-text("{text}")'
                        ]
                        
                        for text_selector in text_selectors:
                            try:
                                element = await page.query_selector(text_selector)
                                if element:
                                    is_visible = await element.is_visible()
                                    if is_visible:
                                        logger.info(f"Found modal close element with text: {text}")
                                        await element.click()
                                        modal_found = True
                                        await asyncio.sleep(1)
                                        break
                            except:
                                continue
                        
                        if modal_found:
                            break
                            
                    except Exception as e:
                        logger.debug(f"Text pattern {text} failed: {e}")
                        continue
            
            # Try ESC key as fallback
            if not modal_found:
                try:
                    await page.keyboard.press('Escape')
                    logger.debug("Tried ESC key for modal")
                    await asyncio.sleep(0.5)
                except:
                    pass
            
            if modal_found:
                logger.info("Successfully handled modal popup")
            else:
                logger.debug("No modal popup found or already closed")
                
        except Exception as e:
            # Log the error but don't crash the script
            logger.debug(f"Modal handling encountered an error (non-critical): {e}")
        
    def setup_folders(self):
        """Create necessary folders for organized storage"""
        try:
            if not os.path.exists(self.base_folder):
                os.makedirs(self.base_folder)
            
            # Create city folders for current batch of cities (avoid creating too many folders at once)
            cities_to_create = CITIES[self.current_city_index:min(self.current_city_index + 50, len(CITIES))]
            for city in cities_to_create:
                city_folder = os.path.join(self.base_folder, self.sanitize_city_name(city))
                if not os.path.exists(city_folder):
                    os.makedirs(city_folder)
            
            logger.info("Folder structure created successfully")
        except Exception as e:
            logger.error(f"Error creating folder structure: {e}")
    
    def load_cities_progress(self):
        """Load cities progress from previous session"""
        try:
            if os.path.exists(self.cities_progress_file):
                with open(self.cities_progress_file, 'r') as f:
                    cities_progress = json.load(f)
                
                self.completed_cities = set(cities_progress.get('completed_cities', []))
                self.current_city_batch = cities_progress.get('current_city_batch', 0)
                
                logger.info(f"Loaded cities progress: {len(self.completed_cities)} cities completed")
            else:
                # Initialize new cities progress
                self.completed_cities = set()
                self.current_city_batch = 0
                self.save_cities_progress()
                logger.info("Starting new cities progress tracking")
                
        except Exception as e:
            logger.error(f"Error loading cities progress: {e}")
            # Initialize with default values
            self.completed_cities = set()
            self.current_city_batch = 0
    
    def save_cities_progress(self):
        """Save cities progress to file"""
        try:
            cities_progress = {
                'completed_cities': list(self.completed_cities),
                'current_city_batch': self.current_city_batch,
                'total_cities_available': len(CITIES),
                'completion_percentage': (len(self.completed_cities) / len(CITIES)) * 100 if CITIES else 0,
                'last_updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'current_city': self.get_current_city() if hasattr(self, 'current_city_index') else None
            }
            
            with open(self.cities_progress_file, 'w') as f:
                json.dump(cities_progress, f, indent=2)
            
            logger.debug(f"Cities progress saved: {len(self.completed_cities)}/{len(CITIES)} cities completed")
            
        except Exception as e:
            logger.error(f"Error saving cities progress: {e}")
    
    def mark_city_completed(self, city_name: str):
        """Mark a city as completed"""
        self.completed_cities.add(city_name)
        self.save_cities_progress()
        logger.info(f"City marked as completed: {city_name} ({len(self.completed_cities)}/{len(CITIES)})")
    
    def get_cities_stats(self) -> dict:
        """Get statistics about cities progress"""
        total_cities = len(CITIES)
        completed_count = len(self.completed_cities)
        remaining_count = total_cities - completed_count
        completion_percentage = (completed_count / total_cities) * 100 if total_cities > 0 else 0
        
        return {
            'total_cities': total_cities,
            'completed_cities': completed_count,
            'remaining_cities': remaining_count,
            'completion_percentage': round(completion_percentage, 2),
            'current_city': self.get_current_city() if hasattr(self, 'current_city_index') else None
        }
    
    def load_progress(self):
        """Load progress from previous session"""
        try:
            if os.path.exists(self.progress_file):
                with open(self.progress_file, 'r') as f:
                    progress_data = json.load(f)
                
                self.profile_count = progress_data.get('profile_count', 0)
                self.current_city_index = progress_data.get('current_city_index', 0)
                self.city_profile_count = progress_data.get('city_profile_count', 0)
                self.excel_file = progress_data.get('excel_file', self.excel_file)
                
                logger.info(f"Resumed from previous session: {self.profile_count} profiles completed")
            else:
                # Initialize new session
                self.profile_count = 0
                self.current_city_index = 0
                self.city_profile_count = 0
                logger.info("Starting new scraping session")
                
        except Exception as e:
            logger.error(f"Error loading progress: {e}")
            # Initialize with default values
            self.profile_count = 0
            self.current_city_index = 0
            self.city_profile_count = 0
    
    def save_progress(self):
        """Save current progress to file"""
        try:
            progress_data = {
                'profile_count': self.profile_count,
                'current_city_index': self.current_city_index,
                'city_profile_count': self.city_profile_count,
                'excel_file': self.excel_file,
                'last_updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            
            with open(self.progress_file, 'w') as f:
                json.dump(progress_data, f, indent=2)
            
            logger.debug(f"Progress saved: Profile {self.profile_count}")
            
        except Exception as e:
            logger.error(f"Error saving progress: {e}")
    
    def sanitize_city_name(self, city_name: str) -> str:
        """Sanitize city name for folder naming"""
        return re.sub(r'[^\w\s-]', '', city_name).strip().replace(' ', '_').replace(',', '')
    
    def sanitize_filename(self, name: str) -> str:
        """Sanitize name for filename"""
        # Remove special characters and limit length
        sanitized = re.sub(r'[^\w\s-]', '', name).strip().replace(' ', '_')
        # Limit length to avoid filesystem issues
        return sanitized[:20] if sanitized else "Unknown"
    
    def get_current_city(self) -> str:
        """Get current city for scraping"""
        return CITIES[self.current_city_index % len(CITIES)]
    
    def should_switch_city(self) -> bool:
        """Check if we should switch to next city"""
        return self.city_profile_count >= self.profiles_per_city
    
    def switch_city(self):
        """Switch to next city and reset counters"""
        # Mark current city as completed
        current_city = self.get_current_city()
        self.mark_city_completed(current_city)
        
        self.current_city_index += 1
        self.city_profile_count = 0
        
        # Check if we've gone through all cities
        if self.current_city_index >= len(CITIES):
            logger.warning("All cities have been processed! Starting over from the beginning.")
            self.current_city_index = 0
            self.current_city_batch += 1
        
        new_city = self.get_current_city()
        logger.info(f"Switching to city: {new_city}")
        
        # Create folder for new city if it doesn't exist
        city_folder = os.path.join(self.base_folder, self.sanitize_city_name(new_city))
        if not os.path.exists(city_folder):
            os.makedirs(city_folder)
        
        # Save progress after city switch
        self.save_progress()
        self.save_cities_progress()

    async def set_slider_value(self, page: Page, selector: str, target_x: int):
        """
        Move a slider to a target x offset.
        
        :param page: Playwright page object
        :param selector: The selector for the slider rail or thumb
        :param target_x: The horizontal position (pixels) relative to the rail
        """
        await page.goto("https://tinder.com/app/profile", timeout=60000)
        print("navigated to profile page")
        # Wait for the slider rail to be visible
        await page.wait_for_selector('[data-testid="slider-rail"]')
        slider = page.locator(selector).first
        
        # Get bounding box of the slider rail
        box = await slider.bounding_box()
        if not box:
            raise Exception("Slider element not found")

        # Calculate start and target position
        start_x = box["x"] + box["width"] / 2
        start_y = box["y"] + box["height"] / 2
        target_abs_x = box["x"] + target_x

        # Perform drag
        await page.mouse.move(start_x, start_y)
        await page.mouse.down()
        await page.mouse.move(target_abs_x, start_y, steps=10)
        await page.mouse.up()
        await asyncio.sleep(200)  # Wait for any UI updates
    
    async def change_location(self, page, city: str) -> bool:
        """Change Tinder location using Passport feature"""
        try:
            logger.info(f"Changing location to: {city}")
            
            # Navigate to Passport settings
            await page.goto("https://tinder.com/app/settings/plus/passport", 
                           wait_until='domcontentloaded', timeout=30000)
            await asyncio.sleep(5)
            
            # Handle any modal that might appear after navigation
            await self.handle_modal_popup(page)
            
            # Find and interact with location search
            search_input = await page.wait_for_selector('input[aria-label="Search a location"]', timeout=10000)
            await search_input.click()
            await search_input.fill("")
            await asyncio.sleep(1)
            await search_input.type(city, delay=100)
            await asyncio.sleep(3)
            try:
                dropdown_input = await page.query_selector_all('[aria-selected="false"]')
                await dropdown_input[6].click()
            except Exception as e:
                logger.error(f"Failed to click the location")
                await search_input.press("ArrowDown")
                await asyncio.sleep(1)
                
            # Press Enter to search
            await search_input.press('Enter')
            await asyncio.sleep(5)
            
            # Handle any modal after location search
            await self.handle_modal_popup(page)
            await asyncio.sleep(5)

            # Click on "Add new location" button
            add_button = await page.wait_for_selector('[title="Add new location"]', timeout=10000)
            await add_button.click()
            await asyncio.sleep(10)
            
            # Handle any modal after adding location
            await self.handle_modal_popup(page)
            
            logger.info(f"Successfully changed location to {city}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to change location to {city}: {e}")
            return False

    async def download_image_method_1(self, image_url: str, profile_id: str, person_name: str, image_id: int, city_name: str, page: Page) -> str:
        """Method 1: Download using aiohttp (recommended)"""
        try:
            await self.init_session()
            
            # Sanitize and prepare file path
            sanitized_name = self.sanitize_filename(person_name)
            sanitized_city = self.sanitize_city_name(city_name)
            filename = f"{profile_id}_{sanitized_name}_{image_id}.webp"
            city_folder = os.path.join(self.base_folder, sanitized_city)
            
            # Create directory if it doesn't exist
            os.makedirs(city_folder, exist_ok=True)
            file_path = os.path.join(city_folder, filename)

            logger.info(f"Downloading image from URL: {image_url}")

            # Get cookies from the page to maintain session
            cookies = await page.context.cookies()
            cookie_dict = {cookie['name']: cookie['value'] for cookie in cookies}

            async with self.session.get(image_url, cookies=cookie_dict) as response:
                if response.status == 200:
                    content = await response.read()
                    with open(file_path, "wb") as f:
                        f.write(content)
                    logger.info(f"Image saved to {file_path}")
                    return file_path
                else:
                    logger.error(f"Failed to download image. Status: {response.status}")
                    return None

        except Exception as e:
            logger.error(f"Failed to download image {image_id} for {profile_id}: {e}")
            return None

    async def download_image_method_2(self, image_url: str, profile_id: str, person_name: str, image_id: int, city_name: str, page: Page) -> str:
        """Method 2: Download using Playwright navigation (fallback)"""
        try:
            # Sanitize and prepare file path
            sanitized_name = self.sanitize_filename(person_name)
            sanitized_city = self.sanitize_city_name(city_name)
            filename = f"{profile_id}_{sanitized_name}_{image_id}.webp"
            city_folder = os.path.join(self.base_folder, sanitized_city)
            
            # Create directory if it doesn't exist
            os.makedirs(city_folder, exist_ok=True)
            file_path = os.path.join(city_folder, filename)

            logger.info(f"Downloading image via Playwright navigation: {image_url}")

            # Create a new page for image download
            context = page.context
            download_page = await context.new_page()
            
            try:
                # Navigate to the image URL
                response = await download_page.goto(image_url, timeout=30000)
                
                if response and response.status == 200:
                    # Get the image content
                    content = await response.body()
                    
                    with open(file_path, "wb") as f:
                        f.write(content)
                    
                    logger.info(f"Image saved to {file_path}")
                    return file_path
                else:
                    logger.error(f"Failed to navigate to image URL. Status: {response.status if response else 'No response'}")
                    return None
                    
            finally:
                await download_page.close()

        except Exception as e:
            logger.error(f"Failed to download image {image_id} for {profile_id}: {e}")
            return None

    async def download_image_method_3(self, image_url: str, profile_id: str, person_name: str, image_id: int, city_name: str, page: Page) -> str:
        """Method 3: Download using page.route() to intercept requests"""
        try:
            # Sanitize and prepare file path
            sanitized_name = self.sanitize_filename(person_name)
            sanitized_city = self.sanitize_city_name(city_name)
            filename = f"{profile_id}_{sanitized_name}_{image_id}.webp"
            city_folder = os.path.join(self.base_folder, sanitized_city)
            
            # Create directory if it doesn't exist
            os.makedirs(city_folder, exist_ok=True)
            file_path = os.path.join(city_folder, filename)

            logger.info(f"Downloading image via route interception: {image_url}")

            image_data = None
            
            async def handle_route(route):
                nonlocal image_data
                if route.request.url == image_url:
                    response = await route.fetch()
                    if response.status == 200:
                        image_data = await response.body()
                    await route.fulfill(response=response)
                else:
                    await route.continue_()

            # Set up route interception
            await page.route("**/*", handle_route)
            
            try:
                # Trigger the image request
                await page.evaluate(f'''
                    async () => {{
                        const img = new Image();
                        img.crossOrigin = 'anonymous';
                        img.src = "{image_url}";
                        return new Promise((resolve, reject) => {{
                            img.onload = () => resolve();
                            img.onerror = () => reject(new Error('Image failed to load'));
                            setTimeout(() => reject(new Error('Timeout')), 10000);
                        }});
                    }}
                ''')
                
                if image_data:
                    with open(file_path, "wb") as f:
                        f.write(image_data)
                    logger.info(f"Image saved to {file_path}")
                    return file_path
                else:
                    logger.error("No image data received")
                    return None
                    
            finally:
                # Remove route handler
                await page.unroute("**/*")

        except Exception as e:
            logger.error(f"Failed to download image {image_id} for {profile_id}: {e}")
            return None

    async def download_image(self, image_url: str, profile_id: str, person_name: str, image_id: int, city_name: str, page: Page = None) -> str:
        """Main download method that tries multiple approaches"""
        if not image_url:
            return None

        # Try Method 1 first (aiohttp)
        result = await self.download_image_method_1(image_url, profile_id, person_name, image_id, city_name, page)
        if result:
            return result

        # Try Method 2 (Playwright navigation)
        logger.info("Method 1 failed, trying Method 2...")
        result = await self.download_image_method_2(image_url, profile_id, person_name, image_id, city_name, page)
        if result:
            return result

        # Try Method 3 (Route interception)
        logger.info("Method 2 failed, trying Method 3...")
        result = await self.download_image_method_3(image_url, profile_id, person_name, image_id, city_name, page)
        if result:
            return result

        logger.error(f"All download methods failed for image {image_id} of {profile_id}")
        return None

    async def scrape_profile_images(self, page, profile_id: str, person_name: str, city_name: str) -> tuple:
        """Scrape all profile photo background image URLs from a Tinder profile"""
        image_paths = []
        image_urls = []  # Store original URLs

        try:
            await asyncio.sleep(3)  # Allow profile to load

            # Handle any modal that might appear on profile page
            await self.handle_modal_popup(page)

            # Check for both "Photo" and "photo" labels (case variations) to determine max photos
            photo_buttons_count = await page.evaluate('''() => {
                const photoButtons = document.querySelectorAll('button[aria-label^="Photo"]');
                const photoButtonsLower = document.querySelectorAll('button[aria-label^="photo"]');
                const fotoButtons = document.querySelectorAll('button[aria-label^="Foto"]');
                const fotoButtonsLower = document.querySelectorAll('button[aria-label^="foto"]');
                
                return Math.max(
                    photoButtons.length, 
                    photoButtonsLower.length, 
                    fotoButtons.length, 
                    fotoButtonsLower.length
                );
            }''')
            
            max_photos = int(photo_buttons_count / 2 - 1) if photo_buttons_count > 0 else 0
            logger.info(f"Found {max_photos} profile photos to iterate.")

            # Determine which label format is being used (with proper case handling for both Profile and Photo/Foto)
            label_info = await page.evaluate('''() => {
                // Check all possible combinations of Profile/profile and Photo/photo/Foto/foto
                const combinations = [
                    { selector: 'div[aria-label^="Profile Photo"]', profile: "Profile", type: "Photo" },
                    { selector: 'div[aria-label^="Profile photo"]', profile: "Profile", type: "photo" },
                    { selector: 'div[aria-label^="profile Photo"]', profile: "profile", type: "Photo" },
                    { selector: 'div[aria-label^="profile photo"]', profile: "profile", type: "photo" },
                    { selector: 'div[aria-label^="Profile Foto"]', profile: "Profile", type: "Foto" },
                    { selector: 'div[aria-label^="Profile foto"]', profile: "Profile", type: "foto" },
                    { selector: 'div[aria-label^="profile Foto"]', profile: "profile", type: "Foto" },
                    { selector: 'div[aria-label^="profile foto"]', profile: "profile", type: "foto" }
                ];
                
                for (let combo of combinations) {
                    const element = document.querySelector(combo.selector);
                    if (element) {
                        return {
                            profile: combo.profile,
                            type: combo.type,
                            full_pattern: `${combo.profile} ${combo.type}`
                        };
                    }
                }
                
                return null;
            }''')
            
            if not label_info:
                logger.warning("No Profile Photo/Foto buttons found with any case variation")
                return image_paths, image_urls

            profile_word = label_info['profile']
            label_type = label_info['type']
            full_pattern = label_info['full_pattern']

            logger.info(f"Using label pattern: {full_pattern}")

            i = 1

            while i <= max_photos:
                try:
                    # Use the detected label pattern with correct case
                    image_selector = f'[aria-label="{profile_word} {label_type} {i}"]'

                    image_element = await page.query_selector(image_selector)

                    if not image_element:
                        logger.warning(f"Profile {label_type} {i} not found, skipping...")
                        i += 1
                        continue

                    # Move mouse to reveal overlay/buttons
                    box = await image_element.bounding_box()
                    if box:
                        center_x = box["x"] + box["width"] / 2
                        center_y = box["y"] + box["height"] / 2

                        await page.mouse.move(center_x, center_y)
                        await asyncio.sleep(0.5)

                    # Extract background image
                    image_url = await image_element.evaluate(
                        'el => window.getComputedStyle(el).backgroundImage.slice(5, -2)'
                    )

                    if image_url and not image_url.startswith("none"):
                        logger.info(f"[{i}] Extracted image URL: {image_url}")
                        
                        # Store the original URL
                        image_urls.append(image_url)
                        
                        # Download the image
                        image_path = await self.download_image(image_url, profile_id, person_name, i, city_name, page=page)
                        print(f"image path: {image_path}")
                        if image_path:
                            image_paths.append(image_path)
                    else:
                        logger.warning(f"[{i}] No valid background image found.")

                    next_photo_index = i + 1
                    # Check all possible label variations for navigation buttons
                    possible_selectors = [
                        f'button[aria-label="Photo {next_photo_index}"]',
                        f'button[aria-label="photo {next_photo_index}"]',
                        f'button[aria-label="Foto {next_photo_index}"]',
                        f'button[aria-label="foto {next_photo_index}"]',
                        f'button[aria-controls="carousel-item-{next_photo_index}"]'
                    ]

                    next_image_buttons = []
                    used_selector = None

                    for selector in possible_selectors:
                        buttons = await page.query_selector_all(selector)
                        if len(buttons) > 1:
                            next_image_buttons = buttons
                            used_selector = selector
                            break

                    print("NEXT PHOTO BUTTON")
                    print(f'Used selector: {used_selector}')
                    print(f'Found {len(next_image_buttons)} buttons')

                    if len(next_image_buttons) > 1:
                        await next_image_buttons[1].click()
                        await asyncio.sleep(1)
                        # Handle any modal after clicking next image
                        await self.handle_modal_popup(page)
                    else:
                        logger.info(f"No more photo buttons found after photo {i}")
                        break  # Exit the loop if no more photos
                    i += 1

                except Exception as e:
                    logger.warning(f"Error scraping {label_type.lower()} {i}: {e}")
                    i += 1
                    continue

            return image_paths, image_urls

        except Exception as e:
            logger.error(f"Fatal error scraping profile images: {e}")
            return image_paths, image_urls
    
    async def scrape_profile_data(self, page) -> dict:
        """Scrape profile name, age, bio, and additional details using the specified selectors"""
        try:
            profile_data = {}
            
            # Handle any modal before scraping profile data
            await self.handle_modal_popup(page)
            await asyncio.sleep(2)
            # Get name using the specific selector - FIXED to use consistent selector
            try:
                name = await page.evaluate('() => document.querySelectorAll("[itemprop=\'name\']")[1]?.textContent')
                profile_data['name'] = name.strip() if name else "Unknown"
                logger.info(f"Name extracted: {profile_data['name']}")
            except:
                profile_data['name'] = "Unknown"
                logger.warning("Failed to extract name")

            # Get age using the specific selector - FIXED to use consistent selector  
            try:
                age = await page.evaluate('() => document.querySelectorAll("[itemprop=\'age\']")[1]?.textContent')
                profile_data['age'] = clean_text_for_excel(age.strip()) if age else "Unknown"
                logger.info(f"Age extracted: {profile_data['age']}")
            except:
                profile_data['age'] = "Unknown"
                logger.warning("Failed to extract age")
            await asyncio.sleep(2)

            button = page.locator("#main-content [role='button']").first
            await button.scroll_into_view_if_needed()

            # wait until it is stable
            await button.wait_for(state="visible")

            # force the click if overlays cause issues
            await button.click(force=True)
            await asyncio.sleep(2)
            
        

            # Get bio (if available) - keeping existing logic
            try:
                bio_selectors = [
                    '[data-testid="profile-bio"]',
                    '.bio',
                    '[data-cy="profile-bio"]',
                    'div[role="main"] p',
                    'div.bio-text'
                ]
                
                bio = ""
                for selector in bio_selectors:
                    bio_element = await page.query_selector(selector)
                    if bio_element:
                        bio = await bio_element.text_content()
                        if bio and bio.strip():
                            break
                
                profile_data['bio'] = bio.strip() if bio else ""
                logger.info(f"Bio extracted: {len(profile_data['bio'])} characters")
            except:
                profile_data['bio'] = ""
                logger.warning("Failed to extract bio")

            # Extract additional profile details
            try:
                # Get all sections and extract data dynamically with improved logic
                try:
                    all_section_data = await page.evaluate(r'''() => {
                        const results = [];
                        
                        // Find the profileCard container specifically
                        const profileCard = document.querySelector('.profileCard__card');
                        if (!profileCard) {
                            console.log('ProfileCard not found');
                            return [];
                        }
                        
                        const sections = profileCard.querySelectorAll('section');
                        console.log(`Found ${sections.length} sections`);
                        
                        // Skip section[0] and process from section[1] onwards as specified
                        for (let i = 1; i < sections.length; i++) {
                            const section = sections[i];
                            const sectionData = {
                                section_index: i,
                                full_text: section.textContent?.trim() || "",
                                heading: "",
                                content: "",
                                field_data: {},
                                all_list_items: []  // Add this to capture all li elements
                            };
                            
                            // Try to find section heading (h1-h6)
                            const headings = section.querySelectorAll("h1, h2, h3, h4, h5, h6");
                            if (headings.length > 0) {
                                sectionData.heading = headings[0].textContent?.trim() || "";
                            }
                            
                            // Extract ALL list items first
                            const listItems = section.querySelectorAll("li");
                            console.log(`Section ${i}: Found ${listItems.length} list items`);
                            
                            // Capture all list item texts
                            for (let li of listItems) {
                                const liText = li.textContent?.trim() || "";
                                if (liText) {
                                    sectionData.all_list_items.push(liText);
                                }
                            }
                            
                            // Special handling for interests section
                            const headingLower = sectionData.heading.toLowerCase();
                            if (headingLower.includes('interest') || headingLower.includes('hobby')) {
                                // For interests, collect all list items as a comma-separated string
                                if (sectionData.all_list_items.length > 0) {
                                    sectionData.field_data['interests'] = sectionData.all_list_items.join(', ');
                                    console.log(`Found interests: ${sectionData.field_data['interests']}`);
                                }
                            } else {
                                // For non-interests sections, use the original field extraction logic
                                for (let li of listItems) {
                                    const liText = li.textContent?.trim() || "";
                                    if (!liText) continue;
                                    
                                    // Look for any heading (h1-h6) in the list item
                                    const liHeading = li.querySelector("h1, h2, h3, h4, h5, h6");
                                    
                                    if (liHeading) {
                                        const headingText = liHeading.textContent?.trim() || "";
                                        // Get the value by removing the heading text from the full text
                                        const valueText = liText.replace(headingText, "").trim();
                                        
                                        if (headingText && valueText) {
                                            sectionData.field_data[headingText] = valueText;
                                            console.log(`Found field: ${headingText} = ${valueText}`);
                                        }
                                    } else {
                                        // If no heading found, try to extract from div structure
                                        const divs = li.querySelectorAll("div");
                                        if (divs.length >= 2) {
                                            // Sometimes the structure is: first meaningful div = heading, last div = value
                                            let potentialHeading = "";
                                            let potentialValue = "";
                                            
                                            // Find divs with meaningful text content
                                            const meaningfulDivs = [];
                                            for (let div of divs) {
                                                const divText = div.textContent?.trim() || "";
                                                if (divText && divText.length > 0 && !divText.match(/^[\s\n\r]*$/)) {
                                                    meaningfulDivs.push(divText);
                                                }
                                            }
                                            
                                            if (meaningfulDivs.length >= 2) {
                                                potentialHeading = meaningfulDivs[0];
                                                potentialValue = meaningfulDivs[meaningfulDivs.length - 1];
                                                
                                                // Check if they're different and both have content
                                                if (potentialHeading && potentialValue && potentialHeading !== potentialValue) {
                                                    sectionData.field_data[potentialHeading] = potentialValue;
                                                    console.log(`Found field from divs: ${potentialHeading} = ${potentialValue}`);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            
                            // Extract content without heading
                            if (sectionData.heading) {
                                sectionData.content = sectionData.full_text.replace(sectionData.heading, "").trim();
                            } else {
                                sectionData.content = sectionData.full_text;
                            }
                            
                            // Only add sections that have meaningful content
                            if (Object.keys(sectionData.field_data).length > 0 || sectionData.content.length > 0 || sectionData.all_list_items.length > 0) {
                                results.push(sectionData);
                            }
                        }
                        
                        console.log(`Processed ${results.length} meaningful sections`);
                        return results;
                    }''')
                    
                    logger.info(f"Found {len(all_section_data)} sections to process (skipped section[0])")
                    
                    # Process each section with the extracted field data
                    for section_data in all_section_data:
                        section_index = section_data['section_index']
                        heading = clean_text_for_excel(section_data['heading'])
                        content = clean_text_for_excel(section_data['content'])
                        field_data = section_data.get('field_data', {})
                        all_list_items = section_data.get('all_list_items', [])
                        
                        logger.info(f"Processing section {section_index}: '{heading}' with {len(field_data)} fields and {len(all_list_items)} list items")
                        
                        # Handle interests specifically from field_data
                        if 'interests' in field_data:
                            profile_data['interests'] = clean_text_for_excel(field_data['interests'])
                            logger.info(f"Extracted interests: {profile_data['interests']}")
                        
                        # Process other field data
                        for field_heading, field_value in field_data.items():
                            if field_heading != 'interests':  # Skip interests as we handled it above
                                # Clean the field heading for use as a column name
                                clean_field_name = field_heading.lower().replace(' ', '_').replace('-', '_').replace('&', 'and').replace(':', '').strip()
                                
                                # Handle special field mappings
                                if clean_field_name:
                                    # Special handling for common fields
                                    if 'distance' in field_value.lower() or 'away' in field_value.lower():
                                        profile_data['distance'] = field_value
                                    elif 'cm' in field_value and field_value.replace('cm', '').strip().isdigit():
                                        profile_data['height'] = field_value
                                    elif 'lives in' in field_value.lower():
                                        profile_data['location'] = field_value.replace('Lives in', '').strip()
                                    elif 'at ' in field_value and len(field_value.split(' at ')) == 2:
                                        # Job title at Company format
                                        parts = field_value.split(' at ')
                                        profile_data['job_title'] = clean_text_for_excel(parts[0].strip())
                                        profile_data['company'] = clean_text_for_excel(parts[1].strip())
                                    else:
                                        # Use the field as-is
                                        profile_data[clean_field_name] = clean_text_for_excel(field_value)
                                        
                                    logger.info(f"Added field: {clean_field_name} = {field_value}")
                        
                        # Add section heading as a field if it exists and has content (but not for interests)
                        if heading and content and not heading.lower().startswith('interest'):
                            clean_heading = heading.lower().replace(' ', '_').replace('-', '_').replace('&', 'and').replace(':', '').strip()
                            
                            # Special handling for different section types (excluding interests)
                            if 'looking for' in clean_heading:
                                profile_data['looking_for'] = clean_text_for_excel(content)
                            elif 'anthem' in clean_heading or 'music' in clean_heading:
                                profile_data[clean_heading] = clean_text_for_excel(content)
                            elif clean_heading and clean_heading not in ['about_me', 'essentials']:
                                profile_data[clean_heading] = clean_text_for_excel(content)

                except Exception as e:
                    logger.error(f"Error extracting enhanced section data: {e}")
                    import traceback
                    logger.error(traceback.format_exc())

                # Check for expandable content buttons and click them
                try:
                    expandable_buttons_found = await page.evaluate('''() => {
                        const buttons = document.querySelectorAll("button, span, div, [role='button']");
                        const expandTexts = ["show more", "view all", "see more", "expand", "more"];
                        let foundButtons = [];
                        
                        for (let element of buttons) {
                            const text = element.textContent?.toLowerCase() || "";
                            for (let expandText of expandTexts) {
                                if (text.includes(expandText) && text.length < 50) {  // Avoid very long text
                                    foundButtons.push({
                                        text: element.textContent?.trim(),
                                        clickable: true
                                    });
                                    break;
                                }
                            }
                        }
                        
                        return foundButtons.length > 0;
                    }''')
                    
                    if expandable_buttons_found:
                        logger.info("Expandable buttons found, attempting to click them")
                        
                        # Click all expandable buttons
                        await page.evaluate('''() => {
                            const buttons = document.querySelectorAll("button, span, div, [role='button']");
                            const expandTexts = ["show more", "view all", "see more", "expand", "more"];
                            
                            for (let element of buttons) {
                                const text = element.textContent?.toLowerCase() || "";
                                for (let expandText of expandTexts) {
                                    if (text.includes(expandText) && text.length < 50) {
                                        try {
                                            element.click();
                                            console.log(`Clicked: ${element.textContent?.trim()}`);
                                        } catch (e) {
                                            console.log(`Failed to click: ${e.message}`);
                                        }
                                        break;
                                    }
                                }
                            }
                        }''')
                        
                        await asyncio.sleep(3)  # Wait for content to load
                        await self.handle_modal_popup(page)
                        
                        # Re-extract data after expanding
                        try:
                            updated_section_data = await page.evaluate('''() => {
                                const results = [];
                                
                                const profileCard = document.querySelector('.profileCard__card');
                                if (!profileCard) return [];
                                
                                const sections = profileCard.querySelectorAll('section');
                                
                                // Skip section[0] and process from section[1] onwards
                                for (let i = 1; i < sections.length; i++) {
                                    const section = sections[i];
                                    const sectionData = {
                                        section_index: i,
                                        full_text: section.textContent?.trim() || "",
                                        heading: "",
                                        content: "",
                                        field_data: {},
                                        all_list_items: []
                                    };
                                    
                                    const headings = section.querySelectorAll("h1, h2, h3, h4, h5, h6");
                                    if (headings.length > 0) {
                                        sectionData.heading = headings[0].textContent?.trim() || "";
                                    }
                                    
                                    const listItems = section.querySelectorAll("li");
                                    
                                    // Capture all list item texts
                                    for (let li of listItems) {
                                        const liText = li.textContent?.trim() || "";
                                        if (liText) {
                                            sectionData.all_list_items.push(liText);
                                        }
                                    }
                                    
                                    // Special handling for interests section
                                    const headingLower = sectionData.heading.toLowerCase();
                                    if (headingLower.includes('interest') || headingLower.includes('hobby')) {
                                        // For interests, collect all list items as a comma-separated string
                                        if (sectionData.all_list_items.length > 0) {
                                            sectionData.field_data['interests'] = sectionData.all_list_items.join(', ');
                                            console.log(`Found interests after expand: ${sectionData.field_data['interests']}`);
                                        }
                                    } else {
                                        // For non-interests sections, use the original field extraction logic
                                        for (let li of listItems) {
                                            const liText = li.textContent?.trim() || "";
                                            if (!liText) continue;
                                            
                                            const liHeading = li.querySelector("h1, h2, h3, h4, h5, h6");
                                            
                                            if (liHeading) {
                                                const headingText = liHeading.textContent?.trim() || "";
                                                const valueText = liText.replace(headingText, "").trim();
                                                
                                                if (headingText && valueText) {
                                                    sectionData.field_data[headingText] = valueText;
                                                }
                                            } else {
                                                const divs = li.querySelectorAll("div");
                                                if (divs.length >= 2) {
                                                    const meaningfulDivs = [];
                                                    for (let div of divs) {
                                                        const divText = div.textContent?.trim() || "";
                                                        if (divText && divText.length > 0) {
                                                            meaningfulDivs.push(divText);
                                                        }
                                                    }
                                                    
                                                    if (meaningfulDivs.length >= 2) {
                                                        const potentialHeading = meaningfulDivs[0];
                                                        const potentialValue = meaningfulDivs[meaningfulDivs.length - 1];
                                                        
                                                        if (potentialHeading && potentialValue && potentialHeading !== potentialValue) {
                                                            sectionData.field_data[potentialHeading] = potentialValue;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    
                                    if (sectionData.heading) {
                                        sectionData.content = sectionData.full_text.replace(sectionData.heading, "").trim();
                                    } else {
                                        sectionData.content = sectionData.full_text;
                                    }
                                    
                                    if (Object.keys(sectionData.field_data).length > 0 || sectionData.content.length > 0 || sectionData.all_list_items.length > 0) {
                                        results.push(sectionData);
                                    }
                                }
                                
                                return results;
                            }''')
                            
                            # Process additional sections (only add new data)
                            for section_data in updated_section_data:
                                heading = clean_text_for_excel(section_data['heading'])
                                content = clean_text_for_excel(section_data['content'])
                                field_data = section_data.get('field_data', {})
                                
                                # Handle interests specifically from expanded data
                                if 'interests' in field_data and 'interests' not in profile_data:
                                    profile_data['interests'] = clean_text_for_excel(field_data['interests'])
                                    logger.info(f"Found expanded interests: {profile_data['interests']}")
                                
                                if heading and content:
                                    clean_heading = heading.lower().replace(' ', '_').replace('-', '_').replace('&', 'and').replace(':', '').strip()
                                    
                                    # Only add if we don't already have this data
                                    if clean_heading not in profile_data and not clean_heading.startswith('interest'):
                                        if 'looking for' in clean_heading:
                                            profile_data['looking_for'] = clean_text_for_excel(content)
                                        else:
                                            profile_data[clean_heading] = clean_text_for_excel(content)
                                
                                # Process additional field data
                                for field_heading, field_value in field_data.items():
                                    if field_heading != 'interests':  # Skip interests as we handled it above
                                        clean_field_name = field_heading.lower().replace(' ', '_').replace('-', '_').replace('&', 'and').replace(':', '').strip()
                                        if clean_field_name not in profile_data:  # Don't overwrite
                                            profile_data[clean_field_name] = clean_text_for_excel(field_value)
                                            logger.info(f"Found additional expanded field: {field_heading} = {field_value}")
                            
                        except Exception as e:
                            logger.warning(f"Error processing expanded sections: {e}")
                    
                except Exception as e:
                    logger.warning(f"Error handling expandable content: {e}")

            except Exception as e:
                logger.error(f"Error extracting additional profile details: {e}")
                import traceback
                logger.error(traceback.format_exc())
            
            # Log summary of extracted data
            additional_count = len([k for k in profile_data.keys() if k not in ['name', 'age', 'bio']])
            logger.info(f"Profile data extracted: name={profile_data['name']}, age={profile_data['age']}, "
                    f"additional_fields={additional_count}")
            logger.info(f"Additional fields found: {[k for k in profile_data.keys() if k not in ['name', 'age', 'bio']]}")
            
            return profile_data
            
        except Exception as e:
            logger.error(f"Error scraping profile data: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return {
                'name': 'Unknown', 
                'age': 'Unknown', 
                'bio': ''
            }
    # def save_to_excel(self):
    #     """Save all scraped data to Excel file with dynamic columns"""
    #     try:
    #         if not self.scraped_data:
    #             logger.warning("No data to save to Excel")
    #             return
            
    #         # Collect all unique column names from all profiles
    #         all_columns = set()
    #         base_columns = ['Profile_ID', 'Name', 'Age', 'City', 'Bio', 'Image_Count', 'Image_URLs', 'Local_Images', 'Scraped_Date']
    #         all_columns.update(base_columns)
            
    #         # Find all unique additional detail keys from all profiles
    #         for profile in self.scraped_data:
    #             for key in profile.keys():
    #                 if key not in ['profile_id', 'name', 'age', 'city', 'bio', 'local_image_paths', 'image_urls', 'scraped_date']:
    #                     # Convert key to proper column name
    #                     column_name = key.replace('_', ' ').title()
    #                     all_columns.add(column_name)
            
    #         # Convert to sorted list for consistent column order
    #         all_columns = sorted(list(all_columns))
            
    #         # Prepare data for Excel
    #         excel_data = []
    #         for profile in self.scraped_data:
    #             row = {
    #                 'Profile_ID': profile['profile_id'],
    #                 'Name': profile['name'],
    #                 'Age': profile['age'],
    #                 'City': profile['city'],
    #                 'Bio': profile['bio'],
    #                 'Image_Count': len(profile['local_image_paths']),
    #                 'Image_URLs': ' | '.join(profile.get('image_urls', [])),
    #                 'Local_Images': ' | '.join(profile['local_image_paths']),
    #                 'Scraped_Date': profile['scraped_date']
    #             }
                
    #             # Add all additional details as columns
    #             for key, value in profile.items():
    #                 if key not in ['profile_id', 'name', 'age', 'city', 'bio', 'local_image_paths', 'image_urls', 'scraped_date']:
    #                     column_name = key.replace('_', ' ').title()
    #                     row[column_name] = value
                
    #             # Fill missing columns with empty strings
    #             for column in all_columns:
    #                 if column not in row:
    #                     row[column] = ''
                
    #             excel_data.append(row)
            
    #         # Create DataFrame with all columns
    #         df = pd.DataFrame(excel_data)
            
    #         # Reorder columns to have base columns first, then alphabetical additional columns
    #         ordered_columns = []
    #         for base_col in base_columns:
    #             if base_col in df.columns:
    #                 ordered_columns.append(base_col)
            
    #         # Add remaining columns alphabetically
    #         remaining_columns = sorted([col for col in df.columns if col not in base_columns])
    #         ordered_columns.extend(remaining_columns)
            
    #         df = df[ordered_columns]
            
    #         # Save to Excel
    #         df.to_excel(self.excel_file, index=False, engine='openpyxl')
            
    #         logger.info(f"Data saved to Excel file: {self.excel_file}")
    #         logger.info(f"Total columns: {len(df.columns)}")
    #         logger.info(f"Columns: {list(df.columns)}")
            
    #     except Exception as e:
    #         logger.error(f"Error saving to Excel: {e}")
    #         traceback.print_exc()
    
    def save_to_excel(self):
        """Save all scraped data to Excel file with dynamic columns - APPEND mode"""
        try:
            if not self.scraped_data:
                logger.warning("No data to save to Excel")
                return
            
            # Prepare data for the new profiles
            new_excel_data = []
            all_columns = set()
            base_columns = ['Profile_ID', 'Name', 'Age', 'City', 'Bio', 'Image_Count', 'Image_URLs', 'Local_Images', 'Scraped_Date']
            all_columns.update(base_columns)
            
            # Find all unique additional detail keys from new profiles
            for profile in self.scraped_data:
                for key in profile.keys():
                    if key not in ['profile_id', 'name', 'age', 'city', 'bio', 'local_image_paths', 'image_urls', 'scraped_date']:
                        column_name = key.replace('_', ' ').title()
                        all_columns.add(column_name)
            
            # Prepare new data
            for profile in self.scraped_data:
                row = {
                    'Profile_ID': profile['profile_id'],
                    'Name': profile['name'],
                    'Age': profile['age'],
                    'City': profile['city'],
                    'Bio': profile['bio'],
                    'Image_Count': len(profile['local_image_paths']),
                    'Image_URLs': ' | '.join(profile.get('image_urls', [])),
                    'Local_Images': ' | '.join(profile['local_image_paths']),
                    'Scraped_Date': profile['scraped_date']
                }
                
                # Add all additional details as columns
                for key, value in profile.items():
                    if key not in ['profile_id', 'name', 'age', 'city', 'bio', 'local_image_paths', 'image_urls', 'scraped_date']:
                        column_name = key.replace('_', ' ').title()
                        row[column_name] = value
                
                new_excel_data.append(row)
            
            # Create DataFrame for new data
            new_df = pd.DataFrame(new_excel_data)
            
            # Check if Excel file already exists
            if os.path.exists(self.excel_file):
                try:
                    # Read existing data
                    existing_df = pd.read_excel(self.excel_file, engine='openpyxl')
                    logger.info(f"Found existing Excel file with {len(existing_df)} records")
                    
                    # Get existing Profile_IDs to avoid duplicates
                    existing_profile_ids = set(existing_df['Profile_ID'].astype(str)) if 'Profile_ID' in existing_df.columns else set()
                    
                    # Filter out profiles that already exist (prevent duplicates)
                    new_df_filtered = new_df[~new_df['Profile_ID'].astype(str).isin(existing_profile_ids)]
                    
                    if len(new_df_filtered) == 0:
                        logger.info("No new profiles to add (all profiles already exist)")
                        return
                    
                    logger.info(f"Adding {len(new_df_filtered)} new profiles (filtered out {len(new_df) - len(new_df_filtered)} duplicates)")
                    
                    # Combine all unique columns from both existing and new data
                    all_columns.update(existing_df.columns)
                    all_columns = sorted(list(all_columns))
                    
                    # Ensure both dataframes have all columns (fill missing with empty string)
                    for column in all_columns:
                        if column not in existing_df.columns:
                            existing_df[column] = ''
                        if column not in new_df_filtered.columns:
                            new_df_filtered[column] = ''
                    
                    # Reorder columns consistently
                    ordered_columns = []
                    for base_col in base_columns:
                        if base_col in all_columns:
                            ordered_columns.append(base_col)
                    
                    remaining_columns = sorted([col for col in all_columns if col not in base_columns])
                    ordered_columns.extend(remaining_columns)
                    
                    # Reorder both dataframes
                    existing_df = existing_df[ordered_columns]
                    new_df_filtered = new_df_filtered[ordered_columns]
                    
                    # Concatenate existing and new data
                    combined_df = pd.concat([existing_df, new_df_filtered], ignore_index=True)
                    
                    logger.info(f"Combined data: {len(existing_df)} existing + {len(new_df_filtered)} new = {len(combined_df)} total records")
                    
                except Exception as e:
                    logger.error(f"Error reading existing Excel file: {e}")
                    logger.info("Creating new Excel file instead")
                    # If reading fails, treat as new file
                    combined_df = new_df
                    
                    # Fill missing columns with empty strings
                    for column in all_columns:
                        if column not in combined_df.columns:
                            combined_df[column] = ''
                    
                    # Reorder columns
                    ordered_columns = []
                    for base_col in base_columns:
                        if base_col in combined_df.columns:
                            ordered_columns.append(base_col)
                    
                    remaining_columns = sorted([col for col in combined_df.columns if col not in base_columns])
                    ordered_columns.extend(remaining_columns)
                    
                    combined_df = combined_df[ordered_columns]
            else:
                # New file - use new data as is
                logger.info("Creating new Excel file")
                combined_df = new_df
                
                # Fill missing columns with empty strings
                all_columns = sorted(list(all_columns))
                for column in all_columns:
                    if column not in combined_df.columns:
                        combined_df[column] = ''
                
                # Reorder columns
                ordered_columns = []
                for base_col in base_columns:
                    if base_col in all_columns:
                        ordered_columns.append(base_col)
                
                remaining_columns = sorted([col for col in all_columns if col not in base_columns])
                ordered_columns.extend(remaining_columns)
                
                combined_df = combined_df[ordered_columns]
            
            # Save the combined data to Excel
            combined_df.to_excel(self.excel_file, index=False, engine='openpyxl')
            
            logger.info(f"Data appended to Excel file: {self.excel_file}")
            logger.info(f"Total records in file: {len(combined_df)}")
            logger.info(f"Total columns: {len(combined_df.columns)}")
            
            # Clear the scraped_data after successful save to prevent re-saving same data
            self.scraped_data.clear()
        
        except Exception as e:
            logger.error(f"Error saving to Excel: {e}")
            traceback.print_exc()
            
    async def scrape_profiles(self, page, total_profiles: int = 100):
        """Main scraping function with enhanced data collection"""
        try:
            # Navigate to Tinder
            await page.goto("https://tinder.com/app/recs", wait_until='domcontentloaded', timeout=60000)
            await asyncio.sleep(30)
            
            # Handle any initial modal popup
            await self.handle_modal_popup(page)
            await asyncio.sleep(5)
            # await self.set_slider_value(page, '[data-testid="slider-rail"]', 170)
            
            logger.info(f"Starting to scrape {total_profiles} profiles (resuming from {self.profile_count})")
            
            # Initial location change
            success = await self.change_location(page, self.get_current_city())
            await page.goto("https://tinder.com/app/recs", wait_until='domcontentloaded', timeout=60000)
            await asyncio.sleep(10)
            
            # Handle modal after navigation
            await self.handle_modal_popup(page)

            while self.profile_count < total_profiles:
                try:
                    # Check if we need to switch cities
                    if self.should_switch_city():
                        self.switch_city()
                        # Change location when switching cities
                        success = await self.change_location(page, self.get_current_city())
                        if not success:
                            logger.warning("Failed to change location, continuing with current location")
                        # Navigate back to profiles after location change
                        await page.goto("https://tinder.com/app/recs", wait_until='domcontentloaded', timeout=60000)
                        await asyncio.sleep(10)
                        # Handle modal after navigation
                        await self.handle_modal_popup(page)
                    
                    # Increment profile count first, then create profile ID
                    self.profile_count += 1
                    profile_id = f"profile_{self.profile_count:04d}"
                    current_city = self.get_current_city()
                    
                    logger.info(f"Scraping profile {self.profile_count}/{total_profiles} in {current_city}")
                    
                    # Handle modal before scraping profile
                    await self.handle_modal_popup(page)
                    
                    # Scrape profile data first to get the name and additional details
                    profile_data = await self.scrape_profile_data(page)
                    print("Profile data extracted:", profile_data)
                    
                    # Scrape images with person's name
                    local_image_paths, image_urls = await self.scrape_profile_images(page, profile_id, profile_data['name'], current_city)
                    
                    # Compile complete profile data - FIXED: Store all profile_data directly
                    complete_profile_data = {
                        'profile_id': profile_id,
                        'city': current_city,
                        'local_image_paths': local_image_paths,
                        'image_urls': image_urls,
                        'scraped_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    }
                    
                    # Add all profile data (including additional details) to the complete profile
                    complete_profile_data.update(profile_data)
                    
                    print("Complete profile data:", complete_profile_data)
                    
                    # Save data
                    self.scraped_data.append(complete_profile_data)
                    
                    # Update city profile count
                    self.city_profile_count += 1
                    
                    # Save progress and Excel after each profile
                    self.save_progress()
                    self.save_to_excel()
                    
                    # Log cities statistics periodically
                    if self.profile_count % 10 == 0:  # Every 10 profiles
                        stats = self.get_cities_stats()
                        logger.info(f"Cities Progress: {stats['completed_cities']}/{stats['total_cities']} "
                                f"({stats['completion_percentage']}%) completed")
                    
                    logger.info(f"Profile {self.profile_count}/{total_profiles} saved successfully")
                    
                    # Navigate to next profile
                    await self.navigate_to_next_profile(page)
                    
                    # Handle modal after navigation
                    await self.handle_modal_popup(page)
                    
                    # Random delay to avoid detection
                    delay = random.uniform(3, 7)
                    await asyncio.sleep(delay)
                    
                except Exception as e:
                    logger.error(f"Error scraping profile {self.profile_count}: {e}")
                    # Don't increment profile count if scraping failed
                    self.profile_count -= 1
                    # Try to recover by navigating to next profile
                    await self.navigate_to_next_profile(page)
                    await self.handle_modal_popup(page)
                    await asyncio.sleep(5)
                    continue
            
            # Final save and cities progress summary
            self.save_to_excel()
            self.save_cities_progress()
            
            final_stats = self.get_cities_stats()
            logger.info(f"Scraping completed! Total profiles scraped: {self.profile_count}")
            logger.info(f"Cities Progress Summary:")
            logger.info(f"  - Total cities available: {final_stats['total_cities']}")
            logger.info(f"  - Cities completed: {final_stats['completed_cities']}")
            logger.info(f"  - Cities remaining: {final_stats['remaining_cities']}")
            logger.info(f"  - Completion percentage: {final_stats['completion_percentage']}%")
            
        except Exception as e:
            logger.error(f"Critical error in main scraping function: {e}")
            traceback.print_exc()
            
    async def navigate_to_next_profile(self, page):
        """Navigate to next profile (swipe right or use next button)"""
        try:
            # Handle modal before navigation
            await self.handle_modal_popup(page)
            
            # Try to find and click pass/like buttons
            pass_buttons = await page.query_selector_all('.gamepad-button-wrapper button')
            if pass_buttons and len(pass_buttons) > 0:
                # Randomly choose to pass or like (use index 0 or 2 if available)
                button_index = random.choice([0, min(2, len(pass_buttons) - 1)])
                button = pass_buttons[2]
                print(f"next button : {button_index}")
                await button.click()
                await asyncio.sleep(2)
                # Handle modal after clicking
                await self.handle_modal_popup(page)
            else:
                # Try keyboard navigation
                await page.keyboard.press('ArrowRight')
                await asyncio.sleep(2)
                # Handle modal after keyboard navigation
                await self.handle_modal_popup(page)
                
        except Exception as e:
            logger.error(f"Error navigating to next profile: {e}")
            # Fallback to keyboard
            try:
                await page.keyboard.press('ArrowRight')
                await asyncio.sleep(2)
                await self.handle_modal_popup(page)
            except:
                pass

    async def init_session(self):
        """Initialize aiohttp session for image downloads"""
        if not hasattr(self, 'session'):
            import aiohttp
            self.session = aiohttp.ClientSession()

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
        logger.info(f"Environment start response: {response.text}")

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
    print("TINDER PROFILE SCRAPER - ENHANCED VERSION WITH MODAL HANDLING")
    print("="*60)
    print(f"Total profiles to scrape: {total_profiles}")
    print("Features:")
    print("- Resume functionality with progress tracking")
    print("- Profile-by-profile Excel saving")
    print(f"- Geographic diversity using {len(CITIES)} cities from CSV")
    print("- Cities progress tracking with JSON file")
    print("- Images named with person's name: profileId_personName_imageId")
    print("- Organized folder structure by city")
    print("- Local image paths saved in Excel")
    print("- Original image URLs saved in Excel")
    print("- Enhanced profile data extraction (interests, education, distance, etc.)")
    print("- Dynamic Excel columns based on available profile data")
    print("- ENHANCED: Automatic modal popup handling")
    print("- ENHANCED: Safe modal detection (won't crash if no modal)")
    print("- ENHANCED: Consistent name/age extraction using specific selectors")
    print("- Crash recovery support")
    print(f"- Cities source: {'worldcities_filtered.csv' if len(CITIES) > 30 else 'fallback list'}")
    print("="*60)
    
    try:
        result = asyncio.run(main(envId, uniqueId, total_profiles))
        
        print("\n" + "="*60)
        print("SCRAPING COMPLETED!")
        print("="*60)
        
        if isinstance(result, dict):
            if 'error' in result:
                print(f"Error occurred: {result['error']}")
            elif 'success' in result:
                print(f"Successfully scraped {result['profiles_scraped']} profiles")
                print(f"Excel file saved: {result['excel_file']}")
                print(f"Images organized in city folders under 'tinder_profiles/'")
                print(f"Progress file: tinder_profiles/scraping_progress.json")
                print(f"Cities progress file: tinder_profiles/cities_progress.json")
                print(f"Images named as: profileId_personName_imageId.jpg")
                print(f"Original image URLs saved in Excel")
                print(f"Enhanced profile data with dynamic columns")
                print(f"Using {len(CITIES)} cities for geographic diversity")
                print(f"Modal handling: Enabled and crash-safe")
                print(f"Name/Age extraction: Using consistent selectors")
        
        print("="*60)
        
    except Exception as e:
        logger.error(f"Script failed with error: {e}")
        traceback.print_exc()