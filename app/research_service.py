"""
Research Service for collecting data from websites
"""
import requests
from bs4 import BeautifulSoup
from typing import Dict, List, Optional, Any
from datetime import datetime
import json
from urllib.parse import urljoin, urlparse
import time

class ResearchService:
    """Service for collecting research data from websites"""
    
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
    
    def collect_data(self, url: str, category: Optional[str] = None, 
                    extract_text: bool = True, extract_links: bool = False,
                    extract_images: bool = False) -> Dict[str, Any]:
        """
        Collect data from a website
        
        Args:
            url: URL to collect data from
            category: Optional category for the data
            extract_text: Whether to extract text content
            extract_links: Whether to extract links
            extract_images: Whether to extract image URLs
            
        Returns:
            Dictionary containing collected data
        """
        try:
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Extract title
            title = soup.find('title')
            title_text = title.get_text(strip=True) if title else None
            
            # Extract main content
            content = ""
            if extract_text:
                # Remove script and style elements
                for script in soup(["script", "style"]):
                    script.decompose()
                
                # Try to find main content areas
                main_content = soup.find('main') or soup.find('article') or soup.find('div', class_='content')
                if main_content:
                    content = main_content.get_text(separator=' ', strip=True)
                else:
                    # Fallback to body text
                    body = soup.find('body')
                    if body:
                        content = body.get_text(separator=' ', strip=True)
            
            # Extract links
            links = []
            if extract_links:
                for link in soup.find_all('a', href=True):
                    href = link['href']
                    absolute_url = urljoin(url, href)
                    link_text = link.get_text(strip=True)
                    links.append({
                        'url': absolute_url,
                        'text': link_text
                    })
            
            # Extract images
            images = []
            if extract_images:
                for img in soup.find_all('img', src=True):
                    src = img['src']
                    absolute_url = urljoin(url, src)
                    alt_text = img.get('alt', '')
                    images.append({
                        'url': absolute_url,
                        'alt': alt_text
                    })
            
            # Build metadata
            metadata = {
                'status_code': response.status_code,
                'content_type': response.headers.get('Content-Type', ''),
                'content_length': len(response.content),
                'collected_at': datetime.now().isoformat(),
            }
            
            if extract_links:
                metadata['links_count'] = len(links)
                metadata['links'] = links[:50]  # Limit to first 50 links
            
            if extract_images:
                metadata['images_count'] = len(images)
                metadata['images'] = images[:50]  # Limit to first 50 images
            
            return {
                'url': url,
                'title': title_text,
                'content': content[:10000] if content else None,  # Limit content length
                'category': category,
                'metadata': metadata,
                'status': 'collected'
            }
            
        except requests.exceptions.RequestException as e:
            return {
                'url': url,
                'title': None,
                'content': None,
                'category': category,
                'metadata': {
                    'error': str(e),
                    'collected_at': datetime.now().isoformat()
                },
                'status': 'error'
            }
    
    def collect_multiple(self, urls: List[str], category: Optional[str] = None,
                         delay: float = 1.0) -> List[Dict[str, Any]]:
        """
        Collect data from multiple URLs
        
        Args:
            urls: List of URLs to collect from
            category: Optional category for all data
            delay: Delay between requests in seconds
            
        Returns:
            List of collected data dictionaries
        """
        results = []
        for url in urls:
            result = self.collect_data(url, category=category)
            results.append(result)
            time.sleep(delay)  # Be respectful with rate limiting
        return results

