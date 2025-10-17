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

    async def change_location(self, page, city: str) -> bool:
        """Change Tinder location using Passport feature"""
        try:
            logger.info(f"Changing location to: {city}")
            
            # Navigate to Passport settings
            await page.goto("https://tinder.com/app/settings/plus/passport", 
                           wait_until='domcontentloaded', timeout=30000)
            await asyncio.sleep(5)
            
            # Find and interact with location search
            search_input = await page.wait_for_selector('input[aria-label="Search a Location"]', timeout=10000)
            await search_input.click()
            await search_input.fill("")
            await asyncio.sleep(1)
            await search_input.type(city, delay=100)
            await asyncio.sleep(3)
            try:
                dropdown_input = await page.query_selector_all('[aria-selected="false"]')
                await dropdown_input[3].click()
            except Exception as e:
                logger.error(f"Failed to click the location")
                await search_input.press("ArrowDown")
                await asyncio.sleep(1)
                
            # Press Enter to search
            await search_input.press('Enter')
            await asyncio.sleep(5)
            
            # Click on "Add new location" button
            add_button = await page.wait_for_selector('[title="Add new location"]', timeout=10000)
            await add_button.click()
            await asyncio.sleep(10)
            
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

    # async def scrape_profile_images(self, page, profile_id: str, person_name: str, city_name: str) -> tuple:
    #     """Scrape all profile photo background image URLs from a Tinder profile"""
    #     image_paths = []
    #     image_urls = []  # NEW: Store original URLs

    #     try:
    #         await asyncio.sleep(3)  # Allow profile to load

    #         # Dynamically get number of photo buttons
    #         max_photos = await page.evaluate('() => document.querySelectorAll(\'button[aria-label^="Photo"]\').length / 2 - 1')
    #         max_photos = int(max_photos)
    #         logger.info(f"Found {max_photos} profile photos to iterate.")

    #         i = 1

    #         while i <= max_photos:
    #             try:
    #                 image_selector = f'[aria-label="Profile Photo {i}"]'
    #                 image_element = await page.query_selector(image_selector)

    #                 if not image_element:
    #                     logger.warning(f"Profile Photo {i} not found, skipping...")
    #                     i += 1
    #                     continue

    #                 # Move mouse to reveal overlay/buttons
    #                 box = await image_element.bounding_box()
    #                 if box:
    #                     center_x = box["x"] + box["width"] / 2
    #                     center_y = box["y"] + box["height"] / 2

    #                     await page.mouse.move(center_x, center_y)
    #                     await asyncio.sleep(0.5)

    #                 # Extract background image
    #                 image_url = await image_element.evaluate(
    #                     'el => window.getComputedStyle(el).backgroundImage.slice(5, -2)'
    #                 )

    #                 if image_url and not image_url.startswith("none"):
    #                     logger.info(f"[{i}] Extracted image URL: {image_url}")
                        
    #                     # Store the original URL
    #                     image_urls.append(image_url)
                        
    #                     # Download the image
    #                     image_path = await self.download_image(image_url, profile_id, person_name, i, city_name, page=page)
    #                     print(f"image path: {image_path}")
    #                     if image_path:
    #                         image_paths.append(image_path)
    #                 else:
    #                     logger.warning(f"[{i}] No valid background image found.")

    #                 # Try to show next image if one exists
    #                 next_photo_index = i + 1
    #                 next_image_buttons = await page.query_selector_all(f'[aria-label="Photo {next_photo_index}"]')

    #                 if len(next_image_buttons) > 1:
    #                     await next_image_buttons[1].click()
    #                     await asyncio.sleep(1)
    #                 else:
    #                     logger.info(f"No more photo buttons after photo {i}")

    #                 i += 1

    #             except Exception as e:
    #                 logger.warning(f"Error scraping photo {i}: {e}")
    #                 i += 1
    #                 continue

    #         return image_paths, image_urls  # CHANGED: Return both paths and URLs

    #     except Exception as e:
    #         logger.error(f"Fatal error scraping profile images: {e}")
    #         return image_paths, image_urls  # CHANGED: Return both even on error

    async def scrape_profile_images(self, page, profile_id: str, person_name: str, city_name: str) -> tuple:
        """Scrape all profile photo background image URLs from a Tinder profile"""
        image_paths = []
        image_urls = []  # NEW: Store original URLs

        try:
            await asyncio.sleep(3)  # Allow profile to load

            # Check for both "Photo" and "Foto" labels to determine max photos
            photo_buttons_count = await page.evaluate('''() => {
                const photoButtons = document.querySelectorAll('button[aria-label^="Photo"]');
                const fotoButtons = document.querySelectorAll('button[aria-label^="Foto"]');
                return Math.max(photoButtons.length, fotoButtons.length);
            }''')
            
            max_photos = int(photo_buttons_count / 2 - 1) if photo_buttons_count > 0 else 0
            logger.info(f"Found {max_photos} profile photos to iterate.")

            # Determine which label format is being used
            label_type = await page.evaluate('''() => {
                const hasPhoto = document.querySelector('button[aria-label^="Photo"]');
                const hasFoto = document.querySelector('button[aria-label^="Foto"]');
                if (hasPhoto) return "Photo";
                if (hasFoto) return "Foto";
                return null;
            }''')
            
            if not label_type:
                logger.warning("No Photo or Foto buttons found")
                return image_paths, image_urls

            logger.info(f"Using label type: {label_type}")

            i = 1

            while i <= max_photos:
                try:
                    # Use the detected label type
                    image_selector = f'[aria-label="Profile {label_type} {i}"]'
                    # image_selector = f'[aria-label="Foto de perfil {i}"]'

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

                    # Try to show next image if one exists
                    next_photo_index = i + 1
                    next_image_buttons = await page.query_selector_all(f'[aria-label="{label_type} {next_photo_index}"]')

                    if len(next_image_buttons) > 1:
                        await next_image_buttons[1].click()
                        await asyncio.sleep(1)
                    else:
                        logger.info(f"No more {label_type.lower()} buttons after {label_type.lower()} {i}")

                    i += 1

                except Exception as e:
                    logger.warning(f"Error scraping {label_type.lower()} {i}: {e}")
                    i += 1
                    continue

            return image_paths, image_urls  # CHANGED: Return both paths and URLs

        except Exception as e:
            logger.error(f"Fatal error scraping profile images: {e}")
            return image_paths, image_urls  # CHANGED: Return both even on error
    async def scrape_profile_data(self, page) -> dict:
        """Scrape profile name, age, and bio using the specified selectors"""
        try:
            profile_data = {}
            
            # Get name using the specified selector
            try:
                name = await page.evaluate('() => document.querySelectorAll("[itemprop=\'name\']")[1]?.textContent')
                profile_data['name'] = name.strip() if name else "Unknown"
            except:
                profile_data['name'] = "Unknown"

            try:
                age = await page.evaluate('() => document.querySelectorAll("[itemprop=\'age\']")[1]?.textContent')
                profile_data['age'] = age.strip() if age else "Unknown"
            except:
                profile_data['age'] = "Unknown"

            # Get bio (if available)
            try:
                bio_element = await page.query_selector('[data-testid="profile-bio"]')
                if bio_element:
                    profile_data['bio'] = await bio_element.text_content()
                else:
                    profile_data['bio'] = ""
            except:
                profile_data['bio'] = ""
            
            return profile_data
            
        except Exception as e:
            logger.error(f"Error scraping profile data: {e}")
            return {'name': 'Unknown', 'age': 'Unknown', 'bio': ''}

    def save_to_excel(self):
        """Save all scraped data to Excel file"""
        try:
            if not self.scraped_data:
                logger.warning("No data to save to Excel")
                return
            
            # Prepare data for Excel
            excel_data = []
            for profile in self.scraped_data:
                row = {
                    'Profile_ID': profile['profile_id'],
                    'Name': profile['name'],
                    'Age': profile['age'],
                    'City': profile['city'],
                    'Bio': profile['bio'],
                    'Image_Count': len(profile['local_image_paths']),
                    'Image_URLs': ' | '.join(profile.get('image_urls', [])),  # This will now have data
                    'Local_Images': ' | '.join(profile['local_image_paths']),
                    'Scraped_Date': profile['scraped_date']
                }
                excel_data.append(row)
            
            # Create DataFrame and save to Excel
            df = pd.DataFrame(excel_data)
            df.to_excel(self.excel_file, index=False, engine='openpyxl')
            
            logger.debug(f"Data saved to Excel file: {self.excel_file}")
            
        except Exception as e:
            logger.error(f"Error saving to Excel: {e}")

    async def scrape_profiles(self, page, total_profiles: int = 100):
        """Main scraping function"""
        try:
            # Navigate to Tinder
            await page.goto("https://tinder.com/app/recs", wait_until='domcontentloaded', timeout=60000)
            await asyncio.sleep(10)
            
            logger.info(f"Starting to scrape {total_profiles} profiles (resuming from {self.profile_count})")
            
            # Initial location change
            success = await self.change_location(page, self.get_current_city())
            await page.goto("https://tinder.com/app/recs", wait_until='domcontentloaded', timeout=60000)
            await asyncio.sleep(10)
            
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
                    
                    # Increment profile count first, then create profile ID
                    self.profile_count += 1
                    profile_id = f"profile_{self.profile_count:04d}"
                    current_city = self.get_current_city()
                    
                    logger.info(f"Scraping profile {self.profile_count}/{total_profiles} in {current_city}")
                    
                    # Scrape profile data first to get the name
                    profile_data = await self.scrape_profile_data(page)
                    print(profile_data)
                    
                    # Scrape images with person's name - CHANGED: Now returns both paths and URLs
                    local_image_paths, image_urls = await self.scrape_profile_images(page, profile_id, profile_data['name'], current_city)
                    
                    # Compile complete profile data
                    complete_profile_data = {
                        'profile_id': profile_id,
                        'name': profile_data['name'],
                        'age': profile_data['age'],
                        'bio': profile_data['bio'],
                        'city': current_city,
                        'local_image_paths': local_image_paths,
                        'image_urls': image_urls,  # NEW: Store original URLs
                        'scraped_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    }
                    print(complete_profile_data)
                    
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
                    
                    # Random delay to avoid detection
                    delay = random.uniform(3, 7)
                    await asyncio.sleep(delay)
                    
                except Exception as e:
                    logger.error(f"Error scraping profile {self.profile_count}: {e}")
                    # Don't increment profile count if scraping failed
                    self.profile_count -= 1
                    # Try to recover by navigating to next profile
                    await self.navigate_to_next_profile(page)
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
    print(f"- Geographic diversity using {len(CITIES)} cities from CSV")
    print("- Cities progress tracking with JSON file")
    print("- Images named with person's name: profileId_personName_imageId")
    print("- Organized folder structure by city")
    print("- Local image paths saved in Excel")
    print("- Original image URLs saved in Excel")
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
                print(f"❌ Error occurred: {result['error']}")
            elif 'success' in result:
                print(f"✅ Successfully scraped {result['profiles_scraped']} profiles")
                print(f"📊 Excel file saved: {result['excel_file']}")
                print(f"📁 Images organized in city folders under 'tinder_profiles/'")
                print(f"📄 Progress file: tinder_profiles/scraping_progress.json")
                print(f"🏙️  Cities progress file: tinder_profiles/cities_progress.json")
                print(f"🖼️  Images named as: profileId_personName_imageId.jpg")
                print(f"🔗 Original image URLs saved in Excel")
                print(f"🌍 Using {len(CITIES)} cities for geographic diversity")
        
        print("="*60)
        
    except Exception as e:
        logger.error(f"Script failed with error: {e}")
        traceback.print_exc()