import requests
from korean_romanizer.romanizer import Romanizer
from logger import logger

class APIHandler:
    def __init__(self, client_id, client_secret, ncp_client_id=None, ncp_client_secret=None):
        self.client_id = client_id
        self.client_secret = client_secret
        self.ncp_client_id = ncp_client_id or client_id
        self.ncp_client_secret = ncp_client_secret or client_secret
        logger.info("APIHandler initialized")

    def get_romanized_text(self, text, is_company=False):
        """Converts Korean text to English Romanization with business naming rules."""
        try:
            if not text or not isinstance(text, str) or text.lower() == 'nan':
                return ""
            
            processed_text = text
            if is_company:
                # Business naming rules
                if processed_text.endswith("점"):
                    processed_text = processed_text[:-1] + " Branch"
                if processed_text == "본점":
                    processed_text = "Headquarters"
                elif "본점" in processed_text:
                    processed_text = processed_text.replace("본점", " Headquarters")

            r = Romanizer(processed_text)
            romanized = r.romanize()
            
            # Capitalize every word
            if romanized:
                # Use a custom title case to avoid lowercasing existing uppercase letters like ISCC
                romanized = ' '.join(word[0].upper() + word[1:] if word else "" for word in romanized.split())
                
            logger.debug(f"Romanized (is_company={is_company}): {text} -> {romanized}")
            return romanized
        except Exception as e:
            logger.error(f"Romanization error for '{text}': {e}", exc_info=True)
            return text

    def get_naver_data(self, address, company_name=""):
        """Fetches Phone, Zip Code, and English Address from Naver APIs."""
        logger.info(f"Fetching Naver data for address: {address} (Company: {company_name})")
        phone = ""
        zip_code = ""
        english_address = ""

        # 1. Search API for Phone (Search by "Address + Company" for better accuracy)
        try:
            search_query = f"{address} {company_name}" if company_name else address
            search_url = "https://openapi.naver.com/v1/search/local.json"
            headers = {
                "X-Naver-Client-Id": self.client_id,
                "X-Naver-Client-Secret": self.client_secret
            }
            params = {"query": search_query, "display": 1}
            
            res = requests.get(search_url, headers=headers, params=params, timeout=5)
            if res.status_code == 200:
                items = res.json().get('items', [])
                if items:
                    phone = items[0].get('telephone', '')
                    logger.debug(f"Found phone for '{search_query}': {phone}")
            else:
                logger.warning(f"Naver Search failed: {res.text}")
        except Exception as e:
            logger.error(f"Naver Search API error: {e}")

        # 2. Geocoding API for Zip Code and English Address
        try:
            # Note: Geocoding API usually uses NCP credentials
            geo_url = "https://naveropenapi.apigw.ntruss.com/map-geocode/v2/geocode"
            geo_headers = {
                "X-NCP-APIGW-API-KEY-ID": self.ncp_client_id,
                "X-NCP-APIGW-API-KEY": self.ncp_client_secret
            }
            geo_params = {"query": address}
            
            res_geo = requests.get(geo_url, headers=geo_headers, params=geo_params, timeout=5)
            if res_geo.status_code == 200:
                data = res_geo.json()
                addresses = data.get('addresses', [])
                if addresses:
                    # Extract English Address
                    english_address = addresses[0].get('englishAddress', '')
                    logger.debug(f"Found English address: {english_address}")

                    # Extract Zip Code
                    for element in addresses[0].get('addressElements', []):
                        if 'POSTAL_CODE' in element.get('types', []):
                            zip_code = element.get('longName', '')
                            logger.debug(f"Found zip code: {zip_code}")
                            break
            else:
                logger.warning(f"Naver Geocode failed: {res_geo.text}")
        except Exception as e:
            logger.error(f"Naver Geocode API error: {e}")

        return phone, zip_code, english_address
