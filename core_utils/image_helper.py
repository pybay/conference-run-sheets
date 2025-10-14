"""
Image downloading, caching, and Excel insertion utilities.

Handles speaker profile photo downloads from Sessionize, caches locally to avoid
repeated downloads, and inserts images into Excel worksheets.

Why Pillow is Required:
- Source images have inconsistent DPI (72, 240, 350 DPI) causing variable sizing in Excel
- xlsxwriter renders images based on embedded DPI metadata, not just pixel dimensions
- Pillow normalizes all images to 96 DPI (Excel standard) for consistent display
- Also handles PNG transparency, format conversion, and exact dimension control

Design:
- Downloads images from Sessionize (or other sources)
- Normalizes with Pillow: RGB conversion, 96 DPI, exact dimensions, JPEG format
- Caches using original filename + size suffix (no hashing)
- Validates filenames for filesystem safety
- Uses concurrent downloads via ThreadPoolExecutor
- Correctly handles "Not Provided" text placeholders for Speakers who did not provide a photo and invalid URLs
"""

from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
import re
import requests

# Cache directory at project root
CACHE_DIR = Path(__file__).parent.parent / 'images_cache'


def get_image_extension(url: str) -> str:
    """
    Extract image extension from URL. Only handles formats supported by xlsxwriter and Excel.

    Supported formats: PNG, JPG/JPEG, GIF, BMP
    """
    match = re.search(r'\.([a-zA-Z0-9]+)(\?|$)', url)
    if not match:
        return 'jpg'

    ext = match.group(1).lower()

    # Validate Excel-supported formats using match/case
    match ext:
        case 'png' | 'jpg' | 'jpeg' | 'gif' | 'bmp':
            return 'jpg' if ext == 'jpeg' else ext
        case _:
            return 'jpg'  # Default fallback


def get_cache_path_from_url(url: str, normalized_size: tuple[int, int] | None = None) -> Path | None:
    """
    Generate cache file path from URL using the filename from the URL path.

    Args:
        url: Image URL (e.g., https://sessionize.com/image/abc-400o400o1-xyz.png)
        normalized_size: If provided, adds size suffix (e.g., (200, 200) → filename_200x200.jpg)

    Returns:
        Path to cache file, or None if URL is invalid
    """
    from urllib.parse import urlparse

    try:
        return _build_cache_filename(urlparse, url, normalized_size)
    except Exception as e:
        print(f"⚠ Could not parse URL: {url} ({e})")
        return None


def _build_cache_filename(urlparse, url: str, normalized_size: tuple[int, int] | None) -> Path | None:
    """
    Parse URL and build cache filename with optional size suffix.

    Args:
        urlparse: urllib.parse.urlparse function
        url: Image URL to parse
        normalized_size: Optional (width, height) to append to filename

    Returns:
        Path to cache file, or None if URL is invalid
    """
    parsed = urlparse(url)
    filename = Path(parsed.path).name

    # Sanitize: only allow alphanumeric, dash, underscore, dot
    # This catches any weird characters that might break filesystem
    if not re.match(r'^[\w\-.]+$', filename):
        print(f"⚠ Invalid filename characters in URL: {url}")
        return None

    # Validate has an extension
    if '.' not in filename:
        print(f"⚠ No file extension in URL: {url}")
        return None

    if not normalized_size:
        return CACHE_DIR / filename

    width, height = normalized_size
    base_name = Path(filename).stem
    cache_filename = f"{base_name}_{width}x{height}.jpg"  # Normalized images are always JPG
    return CACHE_DIR / cache_filename


def is_image_cached(url: str, normalized_size: tuple[int, int] | None = None) -> bool:
    """Check if image already exists in cache."""
    cache_path = get_cache_path_from_url(url, normalized_size)
    return cache_path is not None and cache_path.exists()


def download_and_cache_image(url: str, timeout: int = 10, target_size: tuple[int, int] = (200, 200)) -> Path | None:
    """
    Download image from URL, normalize and resize it, then cache locally.

    Normalization creates consistent images:
    - Converts to RGB (removes alpha channel from PNG)
    - Sets DPI to 96 (Excel standard)
    - Resizes to target dimensions
    - Saves as JPEG with consistent quality

    Args:
        url: Image URL to download
        timeout: Request timeout in seconds
        target_size: Target (width, height) for resized image (default 200x200)

    Returns:
        Path to cached normalized image file, or None if download/cache fails
    """
    from PIL import Image
    from io import BytesIO

    # Handle "Not Provided" and invalid URLs
    if not url or url in {'Not Provided', 'nan', ''}:
        return None

    # Use normalized cache path with size suffix
    cache_path = get_cache_path_from_url(url, normalized_size=target_size)
    if cache_path is None:
        return None

    # Return immediately if already cached
    if cache_path.exists():
        return cache_path

    try:
        # Ensure cache directory exists
        CACHE_DIR.mkdir(exist_ok=True)

        # Download image
        response = requests.get(url, timeout=timeout)
        response.raise_for_status()

        # Normalize image: consistent DPI, format, color mode, and size
        img = Image.open(BytesIO(response.content))

        # Convert to RGB (removes alpha channel if present)
        if img.mode in ('RGBA', 'LA', 'P'):
            # Create white background
            background = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            background.paste(img, mask=img.split()[-1] if img.mode in ('RGBA', 'LA') else None)
            img = background
        elif img.mode != 'RGB':
            img = img.convert('RGB')

        # Resize to target dimensions
        if img.size != target_size:
            img = img.resize(target_size, Image.Resampling.LANCZOS)

        # Save with consistent settings: JPEG, 96 DPI, quality 85
        img.save(cache_path, 'JPEG', quality=85, dpi=(96, 96))

        return cache_path

    except Exception as e:
        print(f"⚠ Failed to download/normalize {url}: {e}")
        return None


def batch_download_and_resize_images(urls: list[str], target_sizes: list[tuple[int, int]]) -> dict[str, dict[tuple[int, int], Path | None]]:
    """
    Download images once, then create multiple resized/normalized versions.

    More efficient than calling batch_download_images multiple times - downloads
    each image only once from the network, then creates multiple local sizes.

    Args:
        urls: List of image URLs to download
        target_sizes: List of (width, height) tuples for different versions

    Returns:
        Dict mapping URL to dict of {size: path} for each size
    """
    from PIL import Image
    from io import BytesIO
    import time

    start_time = time.time()

    # Filter URLs
    valid_urls = [url for url in urls if url and str(url) not in ['Not Provided', 'nan', '']]

    # Check which images need downloading (any size not cached)
    urls_to_download = []
    for url in valid_urls:
        needs_download = False
        for target_size in target_sizes:
            if not is_image_cached(url, normalized_size=target_size):
                needs_download = True
                break
        if needs_download:
            urls_to_download.append(url)

    cached_count = len(valid_urls) - len(urls_to_download)

    if not urls_to_download:
        print(f"✓ Skipped download - all {len(valid_urls)} speaker images already cached in {len(target_sizes)} sizes")
        return {}

    print("\nStarting speaker image download and resize...")
    print(f"  {len(urls_to_download)} images to download, creating {len(target_sizes)} sizes each")
    print(f"  {cached_count} already fully cached")

    results = {}
    successful = 0
    failed = 0

    # Download images in parallel
    with ThreadPoolExecutor() as executor:
        def download_and_create_sizes(url):
            try:
                # Download once
                response = requests.get(url, timeout=10)
                response.raise_for_status()

                # Open image
                img = Image.open(BytesIO(response.content))

                # Create all requested sizes
                size_paths = {}
                for target_size in target_sizes:
                    cache_path = get_cache_path_from_url(url, normalized_size=target_size)
                    if cache_path and not cache_path.exists():
                        # Ensure cache dir exists
                        CACHE_DIR.mkdir(exist_ok=True)

                        # Normalize: RGB, resize, save with 96 DPI
                        img_resized = img.copy()

                        # Convert to RGB
                        if img_resized.mode in ('RGBA', 'LA', 'P'):
                            background = Image.new('RGB', img_resized.size, (255, 255, 255))
                            if img_resized.mode == 'P':
                                img_resized = img_resized.convert('RGBA')
                            background.paste(img_resized, mask=img_resized.split()[-1] if img_resized.mode in ('RGBA', 'LA') else None)
                            img_resized = background
                        elif img_resized.mode != 'RGB':
                            img_resized = img_resized.convert('RGB')

                        # Resize
                        if img_resized.size != target_size:
                            img_resized = img_resized.resize(target_size, Image.Resampling.LANCZOS)

                        # Save
                        img_resized.save(cache_path, 'JPEG', quality=85, dpi=(96, 96))
                        size_paths[target_size] = cache_path

                return url, size_paths, True
            except Exception as e:
                print(f"  ✗ Failed: {url[:60]}... ({e})")
                return url, {}, False

        futures = [executor.submit(download_and_create_sizes, url) for url in urls_to_download]

        for future in futures:
            url, size_paths, success = future.result()
            results[url] = size_paths
            if success:
                successful += 1
            else:
                failed += 1

    elapsed = time.time() - start_time
    total_images = successful * len(target_sizes)
    print(f"✓ Image processing complete: {successful} downloads, {total_images} images created, "
          f"{failed} failed in {elapsed:.2f}s\n")

    return results


def batch_download_images(urls: list[str], target_size: tuple[int, int] = (200, 200)) -> dict[str, Path | None]:
    """
    Download multiple images in parallel, skipping already-cached images.

    Uses ThreadPoolExecutor with automatic worker count based on system capabilities.
    Filters out already-cached images before starting downloads for efficiency.

    Args:
        urls: List of image URLs to download

    Returns:
        Dict mapping URL to cached file path (or None if failed)
    """
    import time

    start_time = time.time()

    # Filter out already-cached images BEFORE starting any downloads
    urls_to_download = [url for url in urls
                       if url and str(url) not in ['Not Provided', 'nan', '']
                       and not is_image_cached(url, normalized_size=target_size)]

    # Quick return if everything is cached
    cached_count = len(urls) - len(urls_to_download)
    if not urls_to_download:
        print(f"✓ Skipped download - all {len(urls)} speaker images already cached ({target_size[0]}x{target_size[1]})")
        return {url: get_cache_path_from_url(url, normalized_size=target_size) for url in urls}

    print("\nStarting speaker image download...")
    print(f"  {cached_count} already cached, {len(urls_to_download)} to download")

    results = {}
    successful = 0
    failed = 0

    # Download missing images in parallel (ThreadPoolExecutor uses smart defaults)
    with ThreadPoolExecutor() as executor:
        future_to_url = {
            executor.submit(download_and_cache_image, url, 10, target_size): url
            for url in urls_to_download
        }

        for future in future_to_url:
            url = future_to_url[future]
            try:
                results[url] = future.result()
                if results[url]:
                    successful += 1
                else:
                    failed += 1
            except Exception as e:
                print(f"  ✗ Failed: {url[:60]}... ({e})")
                results[url] = None
                failed += 1

    # Add already-cached images to results
    for url in urls:
        if url not in results:
            results[url] = get_cache_path_from_url(url, normalized_size=target_size)

    # Report completion with timing
    elapsed = time.time() - start_time
    print(f"✓ Image download complete: {successful} downloaded, {failed} failed, "
          f"{cached_count} cached in {elapsed:.2f}s\n")

    return results


def insert_image_to_worksheet(
    worksheet,
    row: int,
    col: int,
    url: str,
    fallback_format,
    target_size: tuple[int, int] = (100, 100)
) -> bool:
    """
    Insert cached/normalized image into Excel worksheet.

    Images are pre-normalized to target size with consistent DPI (96) and format (JPEG).
    No scaling needed - images are inserted at actual size (1:1).

    Args:
        worksheet: xlsxwriter worksheet object
        row: Row index for image
        col: Column index for image
        url: Image URL
        fallback_format: Format to use if image insertion fails (writes URL as text)
        target_size: Target image dimensions (width, height) - default 100x100

    Returns:
        True if image inserted successfully, False if fallback to URL text
    """
    # Handle invalid/missing URLs
    if not url or url in {'Not Provided', 'nan', ''}:
        worksheet.write(row, col, '', fallback_format)
        return False

    try:
        # Get cached/normalized image (downloads and normalizes if not cached)
        image_path = download_and_cache_image(url, timeout=10, target_size=target_size)

        if image_path is None:
            # Download failed, write URL as fallback
            worksheet.write(row, col, url, fallback_format)
            return False

        # Insert image at actual size (no scaling needed - already normalized to target_size)
        worksheet.insert_image(
            row, col, str(image_path),
            {
                'x_offset': 5,
                'y_offset': 5,
                'positioning': 2  # Move but don't size with cells
            }
        )
        return True

    except Exception as e:
        print(f"⚠ Failed to insert image from {url}: {e}")
        worksheet.write(row, col, url, fallback_format)
        return False
