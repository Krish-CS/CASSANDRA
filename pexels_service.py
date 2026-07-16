"""
Pexels API Service for Cassandra
Fetches background images for PPT templates
"""

import requests
import os
from typing import List, Dict, Optional

# Pexels API configuration
PEXELS_API_KEY = "DcOz5wUlomPoKtscbUJ0MJ7btjS3SXnDUJpCczd2DrlBzPeIRqKasWQ2"
PEXELS_API_URL = "https://api.pexels.com/v1/search"

# Supported colors for filtering
SUPPORTED_COLORS = [
    "red", "orange", "yellow", "green", "turquoise", 
    "blue", "violet", "pink", "brown", "black", "gray", "white"
]


def fetch_backgrounds(
    color: str = None,
    query: str = "abstract background",
    per_page: int = 30
) -> List[Dict]:
    """
    Fetch landscape background images from Pexels API
    
    Args:
        color: Optional color filter (e.g., 'pink', 'blue', 'violet')
        query: Search query (default: 'abstract background')
        per_page: Number of images to fetch (max 80, default 20)
    
    Returns:
        List of template objects with id, url, thumb_url, photographer
    """
    
    headers = {
        "Authorization": PEXELS_API_KEY
    }
    
    # Build search query - include color for better results
    search_query = query
    if color and color.lower() in SUPPORTED_COLORS:
        # Add color to query for more accurate results
        search_query = f"{color} {query}"
    
    params = {
        "query": search_query,
        "orientation": "landscape",  # 16:9 ratio for PPT slides
        "size": "large",  # High quality images
        "per_page": min(per_page, 80)  # API limit is 80
    }
    
    # Also add color filter if valid (for extra filtering)
    if color and color.lower() in SUPPORTED_COLORS:
        params["color"] = color.lower()

    
    try:
        response = requests.get(PEXELS_API_URL, headers=headers, params=params, timeout=10)
        response.raise_for_status()
        
        data = response.json()
        photos = data.get("photos", [])
        
        # Transform to simplified format
        templates = []
        for photo in photos:
            templates.append({
                "id": photo["id"],
                "url": photo["src"]["large2x"],  # High-res for PPT
                "thumb_url": photo["src"]["medium"],  # Thumbnail for preview
                "small_url": photo["src"]["small"],  # Small for quick load
                "photographer": photo["photographer"],
                "alt": photo.get("alt", f"Background by {photo['photographer']}")
            })
        
        return templates
        
    except requests.exceptions.RequestException as e:
        print(f"Pexels API error: {e}")
        return []




def search_thank_you_images(per_page: int = 80, max_results: int = 100) -> List[Dict]:
    """
    Search Pexels for 'thank you' themed images
    
    Args:
        per_page: Images per API request (max 80)
        max_results: Maximum total images to return (will fetch multiple pages)
    
    Returns:
        List of image objects with id, url, thumb_url, photographer
    """
    headers = {
        "Authorization": PEXELS_API_KEY
    }
    
    all_images = []
    pages_needed = min((max_results + per_page - 1) // per_page, 3)  # Max 3 pages for performance
    
    try:
        for page in range(1, pages_needed + 1):
            params = {
                "query": "thank you gratitude appreciation",
                "orientation": "landscape",
                "size": "large",
                "per_page": min(per_page, 80),
                "page": page
            }
            
            response = requests.get(PEXELS_API_URL, headers=headers, params=params, timeout=10)
            response.raise_for_status()
            
            data = response.json()
            photos = data.get("photos", [])
            
            for photo in photos:
                if len(all_images) >= max_results:
                    break
                    
                all_images.append({
                    "id": photo["id"],
                    "url": photo["src"]["large2x"],  # High-res download
                    "thumb_url": photo["src"]["medium"],  # Gallery display
                    "small_url": photo["src"]["small"],  # Quick preview
                    "photographer": photo["photographer"],
                    "alt": photo.get("alt", f"Thank you image by {photo['photographer']}")
                })
            
            if len(all_images) >= max_results:
                break
        
        return all_images[:max_results]
        
    except requests.exceptions.RequestException as e:
        print(f"Pexels API error (thank you search): {e}")
        return []


def get_supported_colors() -> List[Dict]:

    """
    Get list of supported colors with display info
    
    Returns:
        List of color objects with name and hex value
    """
    color_hex_map = {
        "pink": "#ff69b4",
        "violet": "#8a2be2",
        "blue": "#4169e1",
        "turquoise": "#40e0d0",
        "green": "#32cd32",
        "yellow": "#ffd700",
        "orange": "#ff8c00",
        "red": "#dc143c",
        "white": "#f5f5f5",
        "gray": "#808080",
        "brown": "#8b4513",
        "black": "#2d2d2d"
    }
    
    return [
        {"name": color, "hex": color_hex_map.get(color, "#ffffff")}
        for color in SUPPORTED_COLORS
    ]


if __name__ == "__main__":
    # Quick test
    print("Testing Pexels API...")
    templates = fetch_backgrounds(color="pink", per_page=3)
    print(f"Fetched {len(templates)} templates")
    for t in templates:
        print(f"  - {t['id']}: {t['alt']}")
