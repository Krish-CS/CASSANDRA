"""
Cassandra AI - Slide Generator
Generates comprehensive slide content for topic-based PPT generation
"""

import os
from typing import Dict, List, Any
import logging
from dotenv import load_dotenv
import re
import requests
import json

load_dotenv('.env.ppt')

logger = logging.getLogger(__name__)


class SlideGenerator:
    """Generates comprehensive slide content for topic-based PPT generation"""
    
    def __init__(self):
        self.api_client = None
        self.api_type = None
        self._initialize_api()
    
    def _initialize_api(self):
        """Initialize PPT API from .env.ppt"""
        try:
            ppt_api_type = os.getenv("PPT_API_TYPE", "").lower()
            
            if ppt_api_type == "groq" or os.getenv("PPT_GROQ_API_KEY"):
                from groq import Groq
                api_key = os.getenv("PPT_GROQ_API_KEY")
                if api_key:
                    self.api_client = Groq(api_key=api_key)
                    self.api_type = "groq"
                    print("   Using Groq API")
                    return
            
            if os.getenv("PPT_USE_CEREBRAS", "").lower() == "true":
                from cerebras.cloud.sdk import Cerebras
                api_key = os.getenv("PPT_CEREBRAS_API_KEY")
                if api_key:
                    self.api_client = Cerebras(api_key=api_key)
                    self.api_type = "cerebras"
                    return
            
            if os.getenv("GROQ_API_KEY"):
                from groq import Groq
                self.api_client = Groq(api_key=os.getenv("GROQ_API_KEY"))
                self.api_type = "groq"
                return
        
        except Exception as e:
            logger.error(f"Error initializing API: {str(e)}")
    
    # ========================================================================
    # TOPIC PARSING - Generate 15+ topic-specific slides
    # ========================================================================
    
    async def analyze_overview(self, overview_text: str, project_name: str = "", num_slides: int = 15) -> Dict[str, Any]:
        """Generate topic-specific slide topics"""
        print(f"\n   Generating {num_slides} slides for: {project_name or overview_text[:30]}")
        
        overview_text = overview_text.replace('\t', ' ').replace('\r\n', '\n')
        
        try:
            parsed = await self._parse_overview_with_llm(overview_text, project_name, num_slides)
            slide_topics = parsed.get("slides", [])
            print(f"   Generated {len(slide_topics)} topics")
            return {"success": True, "slides": slide_topics, "project_name": project_name}
        except Exception as e:
            print(f"   Error: {str(e)}")
            return {"success": False, "slides": self._fallback_topics(project_name, num_slides), "error": str(e)}
    
    async def refine_user_titles(self, titles: List[str], project_name: str) -> List[str]:
        """Refine user-provided titles to fix typos and professionalize them"""
        print(f"   Refining {len(titles)} user titles...")
        
        prompt = f"""I have a list of slide titles for a presentation on "{project_name}".
Some might have typos or be informal. Refine them to be professional slide titles.
Keep the SAME NUMBER of slides and roughly the same meaning.

User Input: {json.dumps(titles)}

Return ONLY valid JSON: ["Title 1", "Title 2", ...]"""

        try:
            response = self._call_llm(prompt, 600)
            match = re.search(r'\[.*\]', response, re.DOTALL)
            if match:
                refined = json.loads(match.group(0))
                if isinstance(refined, list) and len(refined) == len(titles):
                    return refined
            return titles # Fallback to original
        except Exception as e:
            print(f"Error refining titles: {e}")
            return titles

    async def _parse_overview_with_llm(self, overview_text: str, topic: str, num_slides: int) -> Dict[str, Any]:
        """Generate topic-specific slide titles using LLM"""
        
        prompt = f"""You are creating a professional presentation about "{topic or overview_text}".

Generate EXACTLY {num_slides} slide topics that DEEPLY explore this subject.

IMPORTANT RULES:
1. First 2 slides: INTRODUCTION and ABSTRACT (always include these)
2. Middle slides: Topic-specific content that dives deep into the subject
   - For technology topics: History, How it works, Syntax/Structure, Components, Implementation, Use Cases
   - For concepts: Definition, Principles, Types, Methodology, Examples, Case Studies
   - For products/tools: Features, Architecture, Installation, Usage, Best Practices
3. Last 4 slides: ADVANTAGES, DISADVANTAGES, FUTURE SCOPE, CONCLUSION (always include these)

EXAMPLE for "Python Programming":
["INTRODUCTION TO PYTHON", "ABSTRACT", "HISTORY OF PYTHON", "PYTHON SYNTAX AND STRUCTURE", "DATA TYPES IN PYTHON", "CONTROL FLOW STATEMENTS", "FUNCTIONS AND MODULES", "OBJECT ORIENTED PROGRAMMING", "FILE HANDLING", "LIBRARIES AND FRAMEWORKS", "APPLICATIONS OF PYTHON", "ADVANTAGES", "DISADVANTAGES", "FUTURE SCOPE", "CONCLUSION"]

EXAMPLE for "Machine Learning":
["INTRODUCTION TO MACHINE LEARNING", "ABSTRACT", "TYPES OF MACHINE LEARNING", "SUPERVISED LEARNING", "UNSUPERVISED LEARNING", "NEURAL NETWORKS", "DEEP LEARNING FUNDAMENTALS", "TRAINING AND TESTING", "POPULAR ML ALGORITHMS", "ML FRAMEWORKS AND TOOLS", "REAL WORLD APPLICATIONS", "ADVANTAGES", "DISADVANTAGES", "FUTURE SCOPE", "CONCLUSION"]

Now generate {num_slides} slide topics for "{topic or overview_text}":
Return ONLY valid JSON: {{"slides": ["SLIDE1", "SLIDE2", ...]}}"""

        try:
            response = self._call_llm(prompt, 600)
            match = re.search(r'\{.*\}', response, re.DOTALL)
            if match:
                result = json.loads(match.group(0))
                slides = result.get("slides", [])
                # Ensure we have the right number
                if len(slides) >= num_slides - 2:
                    slides = slides[:num_slides]
                    
                    # ENSURE CONCLUSION IS LAST - post-process
                    slides = self._ensure_conclusion_last(slides, topic)
                    return {"slides": slides}
            return {"slides": self._fallback_topics(topic, num_slides)}
        except:
            return {"slides": self._fallback_topics(topic, num_slides)}
    
    def _ensure_conclusion_last(self, slides: List[str], topic: str) -> List[str]:
        """Ensure CONCLUSION is always the last slide"""
        # Find and remove any existing conclusion slide
        conclusion_idx = -1
        for i, slide in enumerate(slides):
            if "CONCLUSION" in slide.upper():
                conclusion_idx = i
                break
        
        if conclusion_idx >= 0:
            # Remove from current position
            conclusion_slide = slides.pop(conclusion_idx)
            # Add at the end
            slides.append(conclusion_slide)
        else:
            # No conclusion found, add it
            slides.append("CONCLUSION")
        
        return slides

    
    def _fallback_topics(self, topic: str, num_slides: int = 15) -> List[str]:
        """Default topic-specific slide topics - always ends with CONCLUSION"""
        # Fixed start slides (first 2)
        start_slides = [
            f"INTRODUCTION TO {topic.upper()}",
            "ABSTRACT",
        ]
        
        # Fixed end slides (last 4) - ALWAYS included
        end_slides = [
            "ADVANTAGES",
            "DISADVANTAGES",
            "FUTURE SCOPE",
            "CONCLUSION"
        ]
        
        # Middle content slides
        middle_slides = [
            f"HISTORY AND BACKGROUND",
            f"KEY CONCEPTS",
            f"CORE COMPONENTS",
            f"HOW IT WORKS",
            f"TYPES AND CATEGORIES",
            f"IMPLEMENTATION DETAILS",
            f"TOOLS AND TECHNOLOGIES",
            f"PRACTICAL EXAMPLES",
            f"REAL WORLD APPLICATIONS",
        ]
        
        # Calculate how many middle slides we need
        middle_needed = num_slides - len(start_slides) - len(end_slides)
        
        # Extend middle if needed
        while len(middle_slides) < middle_needed:
            middle_slides.append(f"TOPIC {len(middle_slides) + 1}")
        
        # Take only what we need from middle
        middle_slides = middle_slides[:middle_needed]
        
        # Combine: start + middle + end (CONCLUSION always last)
        return start_slides + middle_slides + end_slides

    
    # ========================================================================
    # CONTENT GENERATION - 8 bullet points, 8-9 line paragraphs
    # ========================================================================
    
    async def generate_ppt_content(
        self,
        toc_structure: Dict[str, Any],
        code_content: str,
        project_name: str,
        content_mode: str = "cassandra"
    ) -> Dict[str, Any]:
        """Generate comprehensive PPT content"""
        print(f"\n   Generating content for: {project_name} (Mode: {content_mode})")
        
        generated_content = {
            "project_name": project_name,
            "abstract": "",
            "chapters": []
        }
        
        # Generate Abstract
        generated_content["abstract"] = await self._generate_abstract(project_name)
        
        # Generate each slide
        total = len(toc_structure.get("chapters", []))
        for idx, chapter in enumerate(toc_structure.get("chapters", []), 1):
            title = chapter.get("title", "")
            print(f"   Slide {idx}/{total}: {title}")
            
            chapter_content = {
                "chapter_number": chapter.get("chapter_number", idx),
                "title": title,
                "sections": []
            }
            
            for section in chapter.get("sections", []):
                section_title = section.get("title", title)
                content = await self._generate_section(section_title, project_name, content_mode)
                
                chapter_content["sections"].append({
                    "number": section.get("number", ""),
                    "title": section_title,
                    "content": content
                })
            
            generated_content["chapters"].append(chapter_content)
        
        print(f"   Content generation complete")
        return generated_content
    
    async def _generate_abstract(self, topic: str) -> str:
        """Generate 8-9 line abstract (paragraph format)"""
        prompt = f"""Write a comprehensive ABSTRACT about "{topic}" for a professional presentation.

REQUIREMENTS:
- 8-9 sentences (180-220 words)
- Professional academic tone
- Cover: What it is, why it matters, key features, applications
- NO bullet points, just paragraph format

Write the abstract:"""

        try:
            content = self._call_llm(prompt, 400)
            return self._clean_paragraph(content)
        except:
            return f"{topic} is a significant advancement in modern technology with wide-ranging applications across various industries. It provides innovative solutions to complex problems through its unique approach and methodology. The fundamental principles underlying {topic} enable efficient and effective implementation in diverse scenarios. Organizations and individuals leverage {topic} to achieve better outcomes and improved performance. The field continues to evolve with new developments and innovations. Research and development efforts are driving continuous improvements. This presentation explores the key aspects, benefits, and practical applications of {topic}. Understanding these concepts is essential for professionals in this domain."
    
    async def _generate_section(self, section_title: str, topic: str, content_mode: str = "cassandra") -> str:
        """Generate content based on section type and content mode"""
        
        section_upper = section_title.upper()
        
        # Determine strict style based on mode
        use_paragraph = False
        
        if content_mode == 'para':
            use_paragraph = True
        elif content_mode == 'point':
            use_paragraph = False
        else:
            # Cassandra Mode (Default)
            # PARAGRAPH sections (intro/conclusion/abstract only)
            if any(word in section_upper for word in ["INTRODUCTION", "CONCLUSION", "ABSTRACT"]):
                use_paragraph = True
        
        if use_paragraph:
            return await self._generate_paragraph(section_title, topic)
        else:
            return await self._generate_bullets(section_title, topic)
    
    async def _generate_paragraph(self, section: str, topic: str) -> str:
        """Generate paragraph content (10-11 sentences)"""
        prompt = f"""Write a comprehensive paragraph about "{section}" for a presentation on "{topic}".

REQUIREMENTS:
- 10-11 sentences (220-280 words)
- Professional academic tone
- Informative and detailed
- NO bullet points

Write the paragraph:"""

        try:
            content = self._call_llm(prompt, 500)
            return self._clean_paragraph(content)
        except:
            return f"This section provides a comprehensive overview of {section.lower()} in the context of {topic}. Understanding these fundamentals is essential for effective implementation and utilization. The concepts presented here form the foundation for advanced topics covered in subsequent sections. Practical applications and real-world examples demonstrate the relevance and importance of this subject matter. The field has evolved significantly over the years with continuous innovations. Modern approaches incorporate best practices from various domains. By mastering these concepts, professionals can leverage {topic} to achieve significant improvements in their respective domains. This knowledge is crucial for anyone working in this field. The ongoing research and development continues to drive new discoveries. Organizations worldwide are investing in these technologies to stay competitive."

    
    async def _generate_bullets(self, section: str, topic: str) -> str:
        """Generate exactly 8 crisp bullet points"""
        
        prompt = f"""Generate exactly 8 bullet points about "{section}" for a presentation on "{topic}".

CRITICAL RULES:
1. Each bullet point must be ONE clear sentence (10-15 words)
2. Each point must END with a period
3. Be specific and informative
4. NO sub-points, NO colons in the middle
5. Points must be relevant to the section topic

FORMAT (exactly like this):
Provides efficient data processing capabilities for large scale applications.
Enables seamless integration with existing enterprise systems.
Supports multiple programming languages and development frameworks.
Offers robust security features for data protection.
Facilitates real-time analytics and decision making processes.
Ensures high availability and fault tolerance mechanisms.
Delivers comprehensive monitoring and logging capabilities.
Enables rapid deployment and scaling of applications.

Now generate 8 bullet points about "{section}" for "{topic}":"""

        try:
            content = self._call_llm(prompt, 500)
            return self._format_bullets(content)
        except:
            return self._default_bullets(section, topic)
    
    def _format_bullets(self, content: str) -> str:
        """Clean and format bullet points - ensure 8 points"""
        lines = content.strip().split('\n')
        bullets = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Remove any existing bullet markers
            line = re.sub(r'^[\s\-\*\•\➢\➤\►\▶\→\d\.\)\:]+\s*', '', line)
            line = line.strip()
            
            if len(line) < 15:
                continue
            
            # Truncate if too long (max 90 chars for slide fit)
            if len(line) > 90:
                cut = line[:90].rfind(' ')
                if cut > 50:
                    line = line[:cut]
            
            # Ensure ends with period
            if line and line[-1] not in '.!?':
                line = line + '.'
            
            # Capitalize first letter
            if line:
                line = line[0].upper() + line[1:]
            
            bullets.append(line)
            
            if len(bullets) >= 8:
                break
        
        # Ensure we have 8 bullets
        while len(bullets) < 8:
            bullets.append(f"Provides essential capabilities for effective {bullets[0].split()[0].lower() if bullets else 'implementation'}.")
        
        return '\n'.join(bullets[:8])
    
    def _clean_paragraph(self, content: str) -> str:
        """Clean paragraph content"""
        # Remove markdown
        content = re.sub(r'\*\*(.+?)\*\*', r'\1', content)
        content = re.sub(r'^\#+\s+', '', content, flags=re.MULTILINE)
        content = re.sub(r'^[\-\*\•]\s+', '', content, flags=re.MULTILINE)
        
        # Join into single paragraph
        content = ' '.join(content.split())
        
        # Ensure minimum length (10-11 sentences needs ~500+ chars)
        if len(content) < 500:
            content = content + " This aspect plays a crucial role in the overall implementation and effectiveness of the solution. Understanding these concepts is essential for successful application. The ongoing developments in this field continue to expand possibilities. Professionals benefit greatly from staying updated with these advancements."
        
        # Limit maximum length (10-11 sentences needs ~800 chars max)
        if len(content) > 800:
            cut = content[:800].rfind('.')
            if cut > 500:
                content = content[:cut+1]
        
        return content.strip()

    
    def _default_bullets(self, section: str, topic: str) -> str:
        """Fallback bullet points (8 points)"""
        return f"""Provides fundamental capabilities for {topic} implementation.
Enables efficient processing and management of resources.
Supports scalable solutions for various requirements.
Ensures reliable performance across different scenarios.
Facilitates integration with existing systems and workflows.
Offers comprehensive documentation and support resources.
Delivers consistent results in production environments.
Enables rapid development and deployment cycles."""
    
    def _call_llm(self, prompt: str, max_tokens: int = 500) -> str:
        """Call LLM API"""
        try:
            if self.api_type == "groq":
                response = self.api_client.chat.completions.create(
                    messages=[{"role": "user", "content": prompt}],
                    model=os.getenv("PPT_GROQ_MODEL", "llama-3.3-70b-versatile"),
                    max_tokens=max_tokens,
                    temperature=0.7
                )
                return response.choices[0].message.content
            
            elif self.api_type == "cerebras":
                response = self.api_client.chat.completions.create(
                    messages=[{"role": "user", "content": prompt}],
                    model=os.getenv("PPT_CEREBRAS_MODEL", "llama-3.3-70b"),
                    max_tokens=max_tokens,
                    temperature=0.7
                )
                return response.choices[0].message.content
            
            return ""
        except Exception as e:
            logger.error(f"LLM call failed: {str(e)}")
            return ""
    
    # ========================================================================
    # REFINE SLIDE - Regenerate content for a specific slide
    # ========================================================================
    
    async def refine_slide(self, slide_title: str, current_content: str, topic: str, style: str = "bullet") -> str:
        """
        Refine/regenerate content for a specific slide.
        
        Args:
            slide_title: Title of the slide being refined
            current_content: Current content (for context)
            topic: Main presentation topic
            style: 'paragraph' or 'bullet'
            
        Returns:
            New refined content in the same style
        """
        print(f"   Refining slide: {slide_title} (style: {style})")
        
        if style == "paragraph":
            return await self._refine_paragraph(slide_title, current_content, topic)
        else:
            return await self._refine_bullets(slide_title, current_content, topic)
    
    async def _refine_paragraph(self, slide_title: str, current_content: str, topic: str) -> str:
        """Refine paragraph content"""
        prompt = f"""You are refining a slide about "{slide_title}" for a presentation on "{topic}".

Current content: {current_content[:200]}...

Write a NEW, IMPROVED paragraph about "{slide_title}".

REQUIREMENTS:
- 8-9 sentences (180-220 words)
- Professional academic tone
- More detailed and informative than before
- NO bullet points

Write the improved paragraph:"""

        try:
            content = self._call_llm(prompt, 400)
            return self._clean_paragraph(content)
        except:
            return current_content
    
    async def _refine_bullets(self, slide_title: str, current_content: str, topic: str) -> str:
        """Refine bullet point content - generate COMPLETELY NEW points"""
        prompt = f"""You are creating NEW content for a slide about "{slide_title}" in a presentation on "{topic}".

The current slide has some points, but generate COMPLETELY DIFFERENT and NEW points.
DO NOT rephrase or modify the existing points - create FRESH NEW information.

Generate 8 COMPLETELY NEW bullet points about "{slide_title}".

CRITICAL RULES:
1. Each point must be ONE clear sentence (10-15 words)
2. Each point must END with a period
3. Cover DIFFERENT aspects than before
4. Be specific and informative
5. NO sub-points, NO colons, NO numbering

Write 8 fresh new bullet points:"""


        try:
            content = self._call_llm(prompt, 500)
            return self._format_bullets(content)
        except:
            return current_content
