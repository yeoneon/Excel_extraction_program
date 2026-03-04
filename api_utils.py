import requests
from korean_romanizer.romanizer import Romanizer
from logger import logger

class APIHandler:
    def __init__(self, ncp_client_id, ncp_client_secret, kakao_api_key):
        self.ncp_client_id = ncp_client_id
        self.ncp_client_secret = ncp_client_secret
        self.kakao_api_key = kakao_api_key
        logger.info("APIHandler initialized (NCP + Kakao)")

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
                romanized = ' '.join(word[0].upper() + word[1:] if word else "" for word in romanized.split())
                
            logger.debug(f"Romanized (is_company={is_company}): {text} -> {romanized}")
            return romanized
        except Exception as e:
            logger.error(f"Romanization error for '{text}': {e}", exc_info=True)
            return text

    def get_enriched_data(self, address, company_name=""):
        """Fetches Enriched Data: Phone (Kakao), Zip Code (NCP), English Address (NCP)."""
        logger.info(f"Enriching data for: {address} (Company: {company_name})")
        phone = ""
        zip_code = ""
        english_address = ""

        # 1. Geocoding API (NCP) for Zip Code and English Address
        try:
            geo_url = "https://maps.apigw.ntruss.com/map-geocode/v2/geocode"
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
                    english_address = addresses[0].get('englishAddress', '')
                    # Remove country suffix for cleaner format
                    for suffix in [", Republic of Korea", ", South Korea"]:
                        if english_address.endswith(suffix):
                            english_address = english_address[:-len(suffix)].strip()
                            break
                    
                    for element in addresses[0].get('addressElements', []):
                        if 'POSTAL_CODE' in element.get('types', []):
                            zip_code = element.get('longName', '')
                            break
            else:
                logger.warning(f"NCP Geocode failed ({res_geo.status_code}): {res_geo.text}")
        except Exception as e:
            logger.error(f"NCP Geocode error: {e}")

        # 2. Kakao Local Search for Phone
        if company_name:
            try:
                # Extract city/province from address (usually first part)
                city = address.split()[0] if address else ""
                
                kakao_url = "https://dapi.kakao.com/v2/local/search/keyword.json"
                kakao_headers = {"Authorization": f"KakaoAK {self.kakao_api_key}"}
                
                # Try strategies in order
                kakao_strategies = [
                    f"{company_name} {city}".strip(), # Strategy 1: Company + City
                    company_name                      # Strategy 2: Company only
                ]

                for query in kakao_strategies:
                    kakao_params = {"query": query, "size": 1}
                    res_kakao = requests.get(kakao_url, headers=kakao_headers, params=kakao_params, timeout=5)
                    logger.debug(f"Kakao Search ({query}) response: {res_kakao.status_code}")
                    
                    if res_kakao.status_code == 200:
                        documents = res_kakao.json().get('documents', [])
                        if documents:
                            phone = documents[0].get('phone', '')
                            logger.info(f"Found phone via Kakao ('{query}'): {phone}")
                            break # Found results, stop
                        else:
                            logger.info(f"No results for Kakao search: {query}")
                    else:
                        logger.warning(f"Kakao Search failed ({res_kakao.status_code}): {res_kakao.text}")
                        if res_kakao.status_code in [401, 403]:
                            break
            except Exception as e:
                logger.error(f"Kakao Search error: {e}")

        return phone, zip_code, english_address

class KakaoAPIHandler:
    def __init__(self, kakao_api_key):
        self.kakao_api_key = kakao_api_key
        logger.info("KakaoAPIHandler initialized")

    def _format_coord(self, val):
        """Formats coordinate to 6 decimal places, replacing trailing 0 with 1."""
        if not val:
            return ""
        try:
            s = "{:.6f}".format(float(val))
            if s.endswith('0'):
                s = s[:-1] + '1'
            return s
        except (ValueError, TypeError):
            return str(val)

    def get_kakao_data(self, address, company_name=""):
        """Fetches Phone, Zip Code, Latitude, and Longitude using ONLY Kakao API."""
        logger.info(f"Enriching data (Kakao Only) for: {address} (Company: {company_name})")
        phone = ""
        zip_code = ""
        longitude = ""
        latitude = ""
        raw_longitude = ""
        raw_latitude = ""

        if not self.kakao_api_key:
            logger.error("Kakao API key is missing.")
            return phone, zip_code, longitude, latitude

        kakao_headers = {"Authorization": f"KakaoAK {self.kakao_api_key}"}

        # 1. Kakao Keyword Search: Strictly for Phone Number
        try:
            city = address.split()[0] if address else ""
            kakao_url = "https://dapi.kakao.com/v2/local/search/keyword.json"
            kakao_strategies = [f"{company_name} {city}".strip(), company_name]

            for query in kakao_strategies:
                if not query: continue
                params = {"query": query, "size": 1}
                res = requests.get(kakao_url, headers=kakao_headers, params=params, timeout=5)
                if res.status_code == 200:
                    docs = res.json().get('documents', [])
                    if docs:
                        phone = docs[0].get('phone', '')
                        if phone:
                            logger.info(f"Found phone via Kakao Keyword ('{query}'): {phone}")
                            break
        except Exception as e:
            logger.error(f"Kakao Keyword Search error: {e}")

        # 2. Kakao Address Search: Strictly for Zip Code, Longitude (x), Latitude (y)
        if address:
            try:
                addr_url = "https://dapi.kakao.com/v2/local/search/address.json"
                params = {"query": address, "size": 1}
                res = requests.get(addr_url, headers=kakao_headers, params=params, timeout=5)
                if res.status_code == 200:
                    docs = res.json().get('documents', [])
                    if docs:
                        doc = docs[0]
                        
                        # Get Coordinates from the documents root (most reliable for address)
                        raw_longitude = doc.get('x', '')
                        raw_latitude = doc.get('y', '')
                        
                        # Get Zip Code (zone_no) from road_address if available
                        road_addr = doc.get('road_address')
                        if road_addr:
                            zip_code = road_addr.get('zone_no', '')
                            # Sometimes road_address has more specific coordinates
                            if not raw_longitude: raw_longitude = road_addr.get('x', '')
                            if not raw_latitude: raw_latitude = road_addr.get('y', '')
                        
                        # Fallback for zip_code from 'address' object if road_address is missing
                        if not zip_code:
                            jibun_addr = doc.get('address')
                            if jibun_addr:
                                pass # No zone_no in jibun_addr usually

                        longitude = self._format_coord(raw_longitude)
                        latitude = self._format_coord(raw_latitude)
                        
                        logger.info(f"Found Spatial data via Kakao Address: Zip={zip_code}, Lat={latitude}, Long={longitude}")
            except Exception as e:
                logger.error(f"Kakao Address Search error: {e}")

        return phone, zip_code, longitude, latitude
