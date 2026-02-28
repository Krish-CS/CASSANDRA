"""
Utility functions for the AI Report Builder
"""

import logging
from pathlib import Path
from datetime import datetime


def setup_logger(name: str = "ai_report_builder") -> logging.Logger:
    """
    Setup and configure logger
    
    Args:
        name: Name of the logger
        
    Returns:
        Configured logger instance
    """
    # Create logs directory
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    
    # Create logger
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    
    # Remove existing handlers
    if logger.handlers:
        logger.handlers.clear()
    
    # Create formatters
    file_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    console_formatter = logging.Formatter(
        '%(levelname)s - %(message)s'
    )
    
    # File handler
    log_file = log_dir / f"app_{datetime.now().strftime('%Y%m%d')}.log"
    file_handler = logging.FileHandler(log_file)
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(file_formatter)
    
    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(console_formatter)
    
    # Add handlers
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger


def sanitize_filename(filename: str) -> str:
    """
    Sanitize filename to remove invalid characters
    
    Args:
        filename: Original filename
        
    Returns:
        Sanitized filename
    """
    # Replace invalid characters
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    
    return filename


def format_size(bytes_size: int) -> str:
    """
    Format bytes to human-readable size
    
    Args:
        bytes_size: Size in bytes
        
    Returns:
        Formatted size string
    """
    for unit in ['B', 'KB', 'MB', 'GB']:
        if bytes_size < 1024.0:
            return f"{bytes_size:.2f} {unit}"
        bytes_size /= 1024.0
    return f"{bytes_size:.2f} TB"


def validate_docx_file(file_path: str) -> bool:
    """
    Validate if file is a valid .docx file
    
    Args:
        file_path: Path to the file
        
    Returns:
        True if valid, False otherwise
    """
    path = Path(file_path)
    
    # Check if file exists
    if not path.exists():
        return False
    
    # Check extension
    if path.suffix.lower() != '.docx':
        return False
    
    # Check if it's a file
    if not path.is_file():
        return False
    
    return True


def get_project_info() -> dict:
    """
    Get project information
    
    Returns:
        Dictionary with project metadata
    """
    return {
        "name": "AI Report Builder",
        "version": "1.0.0",
        "description": "Automated technical report generation using AI",
        "author": "AI Assistant",
        "created": datetime.now().isoformat()
    }
