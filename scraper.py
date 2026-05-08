"""
Canadian Insurance Brokerage Web Scraper - AUTOMATED VERSION
Automatically scrapes public directories to find hundreds of brokerages

Requirements:
pip install requests beautifulsoup4 pandas lxml
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import re
from urllib.parse import unquote
from typing import List, Dict


class BrokerageScraper:
    def __init__(self):
        self.brokerages = []
        self.seen_names = set()
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }

    def clean_phone(self, phone: str) -> str:
        """Clean phone number to standard format"""
        if not phone:
            return ""
        digits = re.sub(r'\D', '', phone)
        if len(digits) == 10:
            return f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"
        elif len(digits) == 11 and digits[0] == '1':
            return f"({digits[1:4]}) {digits[4:7]}-{digits[7:]}"
        return phone

    def clean_postal(self, postal: str) -> str:
        """Clean postal code to standard format"""
        if not postal:
            return ""
        postal = postal.replace(" ", "").upper()
        if len(postal) == 6 and re.match(r'^[A-Z]\d[A-Z]\d[A-Z]\d$', postal):
            return f"{postal[:3]} {postal[3:]}"
        return postal

    def is_duplicate(self, name: str) -> bool:
        """Check if brokerage name already exists"""
        clean_name = name.lower().strip()
        if clean_name in self.seen_names:
            return True
        self.seen_names.add(clean_name)
        return False

    def scrape_canada411(self, city: str, province: str) -> int:
        """
        Scrape Canada411 business directory
        """
        print(f"  Searching Canada411 for {city}, {province}...")
        found = 0

        try:
            city_formatted = city.replace(' ', '+')
            prov_formatted = province.upper()
            url = f"https://www.canada411.ca/search/?stype=si&what=insurance+brokers&where={city_formatted}%2C+{prov_formatted}"

            response = requests.get(url, headers=self.headers, timeout=15)
            response.raise_for_status()

            soup = BeautifulSoup(response.content, 'html.parser')

            listings = soup.find_all('div', class_='c411Listing')
            if not listings:
                listings = soup.find_all('article')
            if not listings:
                listings = soup.find_all('div', class_='listing')

            for listing in listings[:20]:
                try:
                    name = None
                    name_elem = listing.find('h3') or listing.find('h2') or listing.find('a', class_=re.compile('name', re.I))
                    if name_elem:
                        name = name_elem.get_text(strip=True)

                    if not name or len(name) < 3:
                        continue

                    if self.is_duplicate(name):
                        continue

                    brokerage = {
                        'name': name,
                        'address': '',
                        'city': city,
                        'province': province.upper(),
                        'postal': '',
                        'phone': '',
                        'email': '',
                        'website': '',
                        'specialties': ['Personal Insurance', 'Commercial Insurance'],
                        'notes': f'Source: Canada411 - {city}'
                    }

                    # Extract phone
                    phone_elem = listing.find('a', href=re.compile(r'tel:'))
                    if not phone_elem:
                        phone_elem = listing.find(string=re.compile(r'\(\d{3}\)\s*\d{3}-\d{4}'))
                    if not phone_elem:
                        phone_elem = listing.find(class_=re.compile('phone', re.I))

                    if phone_elem:
                        phone_text = phone_elem if isinstance(phone_elem, str) else phone_elem.get_text(strip=True)
                        brokerage['phone'] = self.clean_phone(phone_text)

                    # Extract address
                    address_elem = listing.find(class_=re.compile('address', re.I)) or listing.find('address')
                    if address_elem:
                        address_text = address_elem.get_text(strip=True)
                        postal_match = re.search(r'([A-Z]\d[A-Z]\s*\d[A-Z]\d)', address_text)
                        if postal_match:
                            brokerage['postal'] = self.clean_postal(postal_match.group(1))
                            address_text = address_text.replace(postal_match.group(0), '').strip()
                        brokerage['address'] = address_text

                    # Extract website
                    website_elem = listing.find('a', href=re.compile(r'http'))
                    if website_elem and 'canada411' not in website_elem.get('href', ''):
                        brokerage['website'] = website_elem.get('href')

                    # Extract email
                    email_elem = listing.find('a', href=re.compile(r'mailto:'))
                    if email_elem:
                        brokerage['email'] = email_elem.get('href').replace('mailto:', '')

                    self.brokerages.append(brokerage)
                    found += 1
                    print(f"    ✓ {name}")

                except Exception as e:
                    continue

            time.sleep(2)

        except Exception as e:
            print(f"    ✗ Error: {str(e)[:50]}")

        return found

    def scrape_yellowpages(self, city: str, province: str) -> int:
        """
        Scrape Yellow Pages Canada
        """
        print(f"  Searching Yellow Pages for {city}, {province}...")
        found = 0

        province_codes = {
            'ON': 'on', 'BC': 'bc', 'AB': 'ab', 'QC': 'qc',
            'MB': 'mb', 'SK': 'sk', 'NB': 'nb', 'NS': 'ns',
            'PE': 'pe', 'NL': 'nl', 'NT': 'nt', 'NU': 'nu', 'YT': 'yt'
        }

        try:
            prov = province_codes.get(province.upper(), 'on')
            city_formatted = city.lower().replace(' ', '-')

            url = f"https://www.yellowpages.ca/search/si/1/insurance+brokers/{city_formatted}+{prov}"

            response = requests.get(url, headers=self.headers, timeout=15)
            response.raise_for_status()

            soup = BeautifulSoup(response.content, 'lxml')

            listings = soup.find_all('div', class_='listing')
            if not listings:
                listings = soup.find_all('div', class_='listing__content')
            if not listings:
                listings = soup.find_all('article', class_='listing')
            if not listings:
                listings = soup.find_all('div', class_=re.compile('listing'))

            for listing in listings[:20]:
                try:
                    # Extract name
                    name = None
                    name_elem = listing.find('h3') or listing.find('h2') or listing.find('a', class_=re.compile('name'))
                    if name_elem:
                        name = name_elem.get_text(strip=True)

                    if not name or len(name) < 3:
                        continue

                    if self.is_duplicate(name):
                        continue

                    brokerage = {
                        'name': name,
                        'address': '',
                        'city': city,
                        'province': province.upper(),
                        'postal': '',
                        'phone': '',
                        'email': '',
                        'website': '',
                        'specialties': ['Personal Insurance', 'Commercial Insurance'],
                        'notes': f'Source: Yellow Pages - {city}'
                    }

                    # Extract phone
                    phone_elem = listing.find('a', href=re.compile(r'tel:'))
                    if not phone_elem:
                        phone_elem = listing.find(class_=re.compile('phone', re.I))
                    if not phone_elem:
                        phone_match = listing.find(string=re.compile(r'\d{3}[-\.\s]?\d{3}[-\.\s]?\d{4}'))
                        if phone_match:
                            phone_elem = phone_match

                    if phone_elem:
                        phone_text = phone_elem if isinstance(phone_elem, str) else phone_elem.get_text(strip=True)
                        # Also check data-phone attribute
                        if hasattr(phone_elem, 'get') and phone_elem.get('data-phone'):
                            phone_text = phone_elem.get('data-phone')
                        brokerage['phone'] = self.clean_phone(phone_text)

                    # If phone still empty, try data-phone attribute directly
                    if not brokerage['phone']:
                        data_phone_elem = listing.find(attrs={'data-phone': True})
                        if data_phone_elem:
                            brokerage['phone'] = self.clean_phone(data_phone_elem.get('data-phone'))

                    # Extract address
                    address_elem = listing.find('span', itemprop='streetAddress')
                    if not address_elem:
                        address_elem = listing.find(class_=re.compile('address', re.I))
                    if not address_elem:
                        address_elem = listing.find('address')

                    if address_elem:
                        address_text = address_elem.get_text(strip=True)
                        postal_match = re.search(r'([A-Z]\d[A-Z]\s*\d[A-Z]\d)', address_text)
                        if postal_match:
                            brokerage['postal'] = self.clean_postal(postal_match.group(1))
                        brokerage['address'] = address_text

                    # Extract postal code separately if not found in address
                    if not brokerage['postal']:
                        postal_elem = listing.find('span', itemprop='postalCode')
                        if postal_elem:
                            brokerage['postal'] = self.clean_postal(postal_elem.get_text(strip=True))

                    # Extract website from YP redirect link
                    website_elem = listing.find('a', class_='mlr__item__cta', href=re.compile(r'redirect='))
                    if website_elem:
                        href = website_elem.get('href', '')
                        redirect_match = re.search(r'redirect=([^&]+)', href)
                        if redirect_match:
                            brokerage['website'] = unquote(redirect_match.group(1))

                    # Extract email
                    email_elem = listing.find('a', href=re.compile(r'mailto:'))
                    if email_elem:
                        brokerage['email'] = email_elem.get('href').replace('mailto:', '')
                    else:
                        email_match = listing.find(string=re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'))
                        if email_match:
                            email_pattern = re.search(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', email_match)
                            if email_pattern:
                                brokerage['email'] = email_pattern.group(0)

                    self.brokerages.append(brokerage)
                    found += 1
                    print(f"    ✓ {name}")

                except Exception as e:
                    continue

            time.sleep(2)

        except Exception as e:
            print(f"    ✗ Error: {str(e)[:50]}")

        return found

    def scrape_google_search(self, city: str, province: str) -> int:
        """
        Simple Google search scraping (limited but finds some)
        """
        print(f"  Searching Google for {city}, {province}...")
        found = 0

        try:
            query = f"insurance broker {city} {province} Canada"
            url = f"https://www.google.com/search?q={query.replace(' ', '+')}"

            response = requests.get(url, headers=self.headers, timeout=15)
            response.raise_for_status()

            soup = BeautifulSoup(response.content, 'html.parser')

            results = soup.find_all('h3')

            for result in results[:10]:
                try:
                    text = result.get_text(strip=True)

                    if 'insurance' not in text.lower():
                        continue

                    name = text.split('|')[0].split('-')[0].strip()

                    if len(name) < 5 or len(name) > 100:
                        continue

                    if self.is_duplicate(name):
                        continue

                    brokerage = {
                        'name': name,
                        'address': '',
                        'city': city,
                        'province': province.upper(),
                        'postal': '',
                        'phone': '',
                        'email': '',
                        'website': '',
                        'specialties': ['Personal Insurance', 'Commercial Insurance'],
                        'notes': f'Source: Google Search - {city} (verify details)'
                    }

                    self.brokerages.append(brokerage)
                    found += 1
                    print(f"    ✓ {name}")

                except Exception as e:
                    continue

            time.sleep(3)

        except Exception as e:
            print(f"    ✗ Error: {str(e)[:50]}")

        return found

    def scrape_all_cities(self, cities_provinces: List[tuple]) -> None:
        """
        Scrape multiple cities using multiple sources
        """
        total_found = 0

        for i, (city, province) in enumerate(cities_provinces, 1):
            print(f"\n[{i}/{len(cities_provinces)}] Processing {city}, {province}")
            print("-" * 60)

            city_total = 0

            city_total += self.scrape_canada411(city, province)
            city_total += self.scrape_yellowpages(city, province)

            # Uncomment to also use Google
            # city_total += self.scrape_google_search(city, province)

            total_found += city_total
            print(f"  → Found {city_total} brokerages in {city}")

            if i < len(cities_provinces):
                time.sleep(3)

        print(f"\n{'='*60}")
        print(f"✓ TOTAL FOUND: {total_found} brokerages across {len(cities_provinces)} cities")
        print(f"{'='*60}")

    def export_to_csv(self, filename: str = 'canadian_brokerages_scraped.csv'):
        """Export collected brokerages to CSV"""
        if not self.brokerages:
            print("\n⚠ No brokerages found to export!")
            return

        df = pd.DataFrame(self.brokerages)

        df['specialties'] = df['specialties'].apply(lambda x: '; '.join(x) if isinstance(x, list) else x)

        columns_order = ['name', 'city', 'province', 'address', 'postal', 'phone', 'email', 'website', 'specialties', 'notes']
        df = df[columns_order]

        df.to_csv(filename, index=False, encoding='utf-8')
        print(f"\n✓ Exported {len(self.brokerages)} brokerages to: {filename}")

    def print_summary(self):
        """Print summary of collected data"""
        if not self.brokerages:
            print("\n⚠ No brokerages collected")
            return

        df = pd.DataFrame(self.brokerages)
        print(f"\n{'='*60}")
        print(f"DATABASE SUMMARY")
        print(f"{'='*60}")
        print(f"Total Brokerages: {len(self.brokerages)}")
        print(f"\nBy Province:")
        print(df['province'].value_counts().to_string())
        print(f"\nTop 10 Cities:")
        print(df['city'].value_counts().head(10).to_string())
        print(f"\nWith Phone Numbers: {df['phone'].ne('').sum()}")
        print(f"With Addresses: {df['address'].ne('').sum()}")
        print(f"With Websites: {df['website'].ne('').sum()}")


def main():
    print("\n" + "="*60)
    print("AUTOMATED CANADIAN INSURANCE BROKERAGE SCRAPER")
    print("="*60)
    print("\nThis will scrape multiple cities across Canada.")
    print("Estimated time: 10-20 minutes")
    print("Expected results: 100-300+ brokerages\n")

    scraper = BrokerageScraper()

    cities_to_scrape = [
        # Ontario (largest market)
        ('Toronto', 'ON'),
        ('Mississauga', 'ON'),
        ('Ottawa', 'ON'),
        ('Brampton', 'ON'),
        ('Hamilton', 'ON'),
        ('London', 'ON'),
        ('Markham', 'ON'),
        ('Vaughan', 'ON'),
        ('Kitchener', 'ON'),
        ('Windsor', 'ON'),
        ('Richmond Hill', 'ON'),
        ('Oakville', 'ON'),
        ('Burlington', 'ON'),
        ('Barrie', 'ON'),
        ('Oshawa', 'ON'),
        ('St Catharines', 'ON'),
        ('Cambridge', 'ON'),
        ('Kingston', 'ON'),
        ('Guelph', 'ON'),
        ('Whitby', 'ON'),

        # British Columbia
        ('Vancouver', 'BC'),
        ('Surrey', 'BC'),
        ('Burnaby', 'BC'),
        ('Richmond', 'BC'),
        ('Abbotsford', 'BC'),
        ('Coquitlam', 'BC'),
        ('Kelowna', 'BC'),
        ('Victoria', 'BC'),
        ('Langley', 'BC'),
        ('Delta', 'BC'),
        ('Kamloops', 'BC'),
        ('Nanaimo', 'BC'),

        # Alberta
        ('Calgary', 'AB'),
        ('Edmonton', 'AB'),
        ('Red Deer', 'AB'),
        ('Lethbridge', 'AB'),
        ('St Albert', 'AB'),
        ('Medicine Hat', 'AB'),
        ('Grande Prairie', 'AB'),
        ('Airdrie', 'AB'),

        # Quebec
        ('Montreal', 'QC'),
        ('Quebec City', 'QC'),
        ('Laval', 'QC'),
        ('Gatineau', 'QC'),
        ('Longueuil', 'QC'),
        ('Sherbrooke', 'QC'),
        ('Trois-Rivieres', 'QC'),

        # Manitoba
        ('Winnipeg', 'MB'),
        ('Brandon', 'MB'),

        # Saskatchewan
        ('Saskatoon', 'SK'),
        ('Regina', 'SK'),

        # Nova Scotia
        ('Halifax', 'NS'),
        ('Dartmouth', 'NS'),

        # New Brunswick
        ('Moncton', 'NB'),
        ('Saint John', 'NB'),
        ('Fredericton', 'NB'),

        # Newfoundland
        ('St Johns', 'NL'),

        # PEI
        ('Charlottetown', 'PE'),
    ]

    print(f"Will search {len(cities_to_scrape)} cities across Canada\n")

    scraper.scrape_all_cities(cities_to_scrape)
    scraper.print_summary()
    scraper.export_to_csv('canadian_brokerages_scraped.csv')

    print("\n" + "="*60)
    print("✓ SCRAPING COMPLETE")
    print("="*60)
    print("\nNEXT STEPS:")
    print("1. Open 'canadian_brokerages_scraped.csv'")
    print("2. Review and clean the data")
    print("3. Run enrichment_layer2.py to enrich with website data")
    print("4. Start your outreach!\n")


if __name__ == "__main__":
    main()
