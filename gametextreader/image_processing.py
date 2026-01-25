"""
Image processing functions for OCR preprocessing
"""
from PIL import Image, ImageEnhance, ImageFilter
import numpy as np


def preprocess_image(image, brightness=1.0, contrast=1.0, saturation=1.0, sharpness=1.0, blur=0.0, threshold=None, hue=0.0, exposure=1.0, color_mask_enabled=False, color_mask_color='#FF0000', color_mask_tolerance=30, color_mask_background='black', color_mask_position='after', color_mask_text_mode=False, color_mask_preserve_edges=False, color_mask_enhance_contrast=False):
    """Apply various image preprocessing filters to improve OCR accuracy"""

    # Apply color mask FIRST if position is "before"
    if color_mask_enabled and color_mask_position == 'before':
        image = apply_color_mask(image, color_mask_color, color_mask_tolerance, color_mask_background, color_mask_text_mode, color_mask_preserve_edges, color_mask_enhance_contrast)

    # Batch apply PIL enhancements for better performance
    enhancer_params = []
    if brightness != 1.0:
        enhancer_params.append(('brightness', brightness))
    if contrast != 1.0:
        enhancer_params.append(('contrast', contrast))
    if saturation != 1.0:
        enhancer_params.append(('color', saturation))
    if sharpness != 1.0:
        enhancer_params.append(('sharpness', sharpness))
    
    # Apply enhancements in sequence
    for param_type, value in enhancer_params:
        if param_type == 'brightness':
            enhancer = ImageEnhance.Brightness(image)
        elif param_type == 'contrast':
            enhancer = ImageEnhance.Contrast(image)
        elif param_type == 'color':
            enhancer = ImageEnhance.Color(image)
        elif param_type == 'sharpness':
            enhancer = ImageEnhance.Sharpness(image)
        image = enhancer.enhance(value)

    # Apply blur if needed
    if blur > 0:
        image = image.filter(ImageFilter.GaussianBlur(blur))

    # Apply threshold if not None
    if threshold is not None:
        image = image.point(lambda p: p > threshold and 255)

    # Combine hue and exposure operations to reduce conversions
    if hue != 0.0 or exposure != 1.0:
        # Convert to HSV once
        image = image.convert('HSV')
        channels = list(image.split())
        
        # Apply hue shift if needed
        if hue != 0.0:
            channels[0] = channels[0].point(lambda p: (p + int(hue * 255)) % 256)
        
        # Convert back to RGB
        image = Image.merge('HSV', channels).convert('RGB')
        
        # Apply exposure (brightness) after HSV conversion
        if exposure != 1.0:
            enhancer = ImageEnhance.Brightness(image)
            image = enhancer.enhance(exposure)

    # Apply color mask LAST if position is "after"
    if color_mask_enabled and color_mask_position == 'after':
        image = apply_color_mask(image, color_mask_color, color_mask_tolerance, color_mask_background, color_mask_text_mode, color_mask_preserve_edges, color_mask_enhance_contrast)

    return image


def apply_color_mask(image, target_color, tolerance, background='black', text_mode=False, preserve_edges=False, enhance_contrast=False):
    """Apply color mask to image using optimized numpy operations"""
    try:
        
        # Convert hex color to RGB
        hex_color = target_color.lstrip('#')
        target_r = int(hex_color[0:2], 16)
        target_g = int(hex_color[2:4], 16)
        target_b = int(hex_color[4:6], 16)
        
        
        # Convert image to RGB if not already
        if image.mode != 'RGB':
            image = image.convert('RGB')
        
        # Convert to numpy array for vectorized operations
        img_array = np.array(image)
        
        # Calculate color distance vectorized
        target_rgb = np.array([target_r, target_g, target_b])
        # Use squared distance to avoid sqrt for better performance
        distances_squared = np.sum((img_array - target_rgb) ** 2, axis=2)
        tolerance_squared = tolerance ** 2
        
        # Create boolean mask for pixels within tolerance
        mask = distances_squared <= tolerance_squared
        matching_pixels = np.sum(mask)
        
        
        # Text-specific optimizations
        if text_mode:
            
            # Apply morphological operations to clean up text
            try:
                from scipy import ndimage
                
                # Convert mask to binary for morphological operations
                binary_mask = mask.astype(np.uint8) * 255
                
                # Remove small noise (helps with font artifacts)
                if preserve_edges:
                    # Use opening to remove small noise while preserving edges
                    structure = np.ones((2, 2))
                    cleaned_mask = ndimage.binary_opening(binary_mask > 0, structure=structure)
                    
                    # Close small gaps in characters
                    structure = np.ones((3, 3))
                    cleaned_mask = ndimage.binary_closing(cleaned_mask, structure=structure)
                    
                    mask = cleaned_mask
                
                # Enhance contrast for better OCR
                if enhance_contrast:
                    # Apply adaptive contrast enhancement to text regions
                    text_regions = img_array[mask]
                    if len(text_regions) > 0:
                        # Calculate local contrast and enhance
                        mean_intensity = np.mean(text_regions)
                        std_intensity = np.std(text_regions)
                        
                        if std_intensity > 0:
                            # Enhance contrast by stretching the dynamic range
                            enhanced_text = ((text_regions - mean_intensity) * 1.5 + mean_intensity).clip(0, 255)
                            img_array[mask] = enhanced_text
            except ImportError:
                pass
            except Exception as e:
                pass
        
        # Apply background color to non-matching pixels
        if background == 'white':
            img_array[~mask] = [255, 255, 255]
        else:  # black
            img_array[~mask] = [0, 0, 0]
        
        # Convert back to PIL Image
        masked_image = Image.fromarray(img_array.astype(np.uint8))
        
        return masked_image
        
    except Exception as e:
        print(f"[ERROR] Error applying color mask: {e}")
        import traceback
        traceback.print_exc()
        return image


def filter_by_color(image, target_color, tolerance=30):
    """
    Filter image to only keep pixels matching the target color within tolerance
    Returns a black and white image where matching pixels are white
    
    Args:
        image: PIL Image to filter
        target_color: Hex color string (e.g., "#FF0000") or RGB tuple
        tolerance: Color distance tolerance (0-100)
    
    Returns:
        PIL Image in grayscale mode with filtered pixels
    """
    print(f"Filtering image by color {target_color} with tolerance {tolerance}")
    
    # Convert to RGB if not already
    if image.mode != 'RGB':
        image = image.convert('RGB')
    
    # Parse target color
    if isinstance(target_color, str) and target_color.startswith('#'):
        # Convert hex to RGB
        hex_color = target_color.lstrip('#')
        target_r = int(hex_color[0:2], 16)
        target_g = int(hex_color[2:4], 16)
        target_b = int(hex_color[4:6], 16)
    elif isinstance(target_color, (tuple, list)) and len(target_color) == 3:
        target_r, target_g, target_b = target_color
    else:
        raise ValueError(f"Invalid color format: {target_color}")
    
    # Create a new image for the filtered result (grayscale)
    filtered = Image.new('L', image.size, 0)  # Black background
    
    # Get pixel data
    pixels = image.load()
    filtered_pixels = filtered.load()
    
    # Calculate maximum distance squared for tolerance
    max_distance_squared = (tolerance * 4.41) ** 2  # 4.41 ≈ sqrt(255^2 + 255^2 + 255^2) / 100
    
    for y in range(image.height):
        for x in range(image.width):
            r, g, b = pixels[x, y]
            # Calculate Euclidean distance squared in RGB space
            distance_squared = (r - target_r)**2 + (g - target_g)**2 + (b - target_b)**2
            if distance_squared <= max_distance_squared:
                filtered_pixels[x, y] = 255  # White for matching pixels
    
    return filtered

