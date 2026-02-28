"""
Cassandra AI - PPT Generation Module

This module handles automated PowerPoint presentation generation from:
- PPT templates with presentation overview
- LLM-powered content generation

Main Components:
- PPTAnalyzer: Analyzes template and extracts overview topics
- SlideGenerator: Generates slide content using LLM
- LayoutManager: Applies template styling to generated slides

Usage:
    from presentation.slide_generator import SlideGenerator
    
    generator = SlideGenerator()
    result = await generator.generate_ppt_content(
        toc_structure=toc,
        code_content="",
        project_name="My Project"
    )
"""

try:
    from presentation.ppt_analyzer import PPTAnalyzer
    from presentation.slide_generator import SlideGenerator
    from presentation.layout_manager import LayoutManager
except ImportError:
    # Fallback for different import contexts
    pass

__all__ = [
    'PPTAnalyzer',
    'SlideGenerator', 
    'LayoutManager'
]

__version__ = '2.0.0'
__author__ = 'K.KRISHKANTH'
__project__ = 'Cassandra AI'

