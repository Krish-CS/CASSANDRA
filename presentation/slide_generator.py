"""
Cassandra AI - Slide Generator
Generates comprehensive slide content for topic-based PPT generation
"""

import sys
import os
from typing import Dict, List, Any
import logging
from dotenv import load_dotenv
import re
import requests
import json

# Ensure stdout handles UTF-8 safely without UnicodeEncodeError
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

# Load environment variables: check .env.ppt, .env, and fall back to system environment
base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ppt_env_path = os.path.join(base_dir, '.env.ppt')
std_env_path = os.path.join(base_dir, '.env')

if os.path.exists(ppt_env_path):
    load_dotenv(ppt_env_path)
if os.path.exists(std_env_path):
    load_dotenv(std_env_path)
load_dotenv()

logger = logging.getLogger(__name__)


class SlideGenerator:
    """Generates comprehensive slide content for topic-based PPT generation"""
    
    def __init__(self):
        self.api_client = None
        self.api_type = None
        self.nvidia_api_key = None
        self._initialize_api()
    
    def _initialize_api(self):
        """Initialize PPT API from environment variables or .env.ppt"""
        try:
            ppt_api_type = os.getenv("PPT_API_TYPE", "").lower()
            groq_key = os.getenv("PPT_GROQ_API_KEY") or os.getenv("GROQ_API_KEY")
            nvidia_key = os.getenv("PPT_NVIDIA_API_KEY") or os.getenv("NVIDIA_API_KEY")
            
            if nvidia_key:
                self.nvidia_api_key = nvidia_key

            if ppt_api_type == "groq" or groq_key:
                if groq_key:
                    from groq import Groq
                    self.api_client = Groq(api_key=groq_key)
                    self.api_type = "groq"
                    print("   Using Groq API")
                    return
            
            if ppt_api_type == "nvidia" or nvidia_key:
                if nvidia_key:
                    try:
                        from openai import OpenAI
                        self.api_client = OpenAI(
                            base_url="https://integrate.api.nvidia.com/v1",
                            api_key=nvidia_key
                        )
                    except Exception:
                        self.api_client = None
                    self.api_type = "nvidia"
                    print("   Using NVIDIA NIM API")
                    return
            
            if os.getenv("PPT_USE_CEREBRAS", "").lower() == "true":
                api_key = os.getenv("PPT_CEREBRAS_API_KEY")
                if api_key:
                    from cerebras.cloud.sdk import Cerebras
                    self.api_client = Cerebras(api_key=api_key)
                    self.api_type = "cerebras"
                    return
        
        except Exception as e:
            logger.error(f"Error initializing API: {str(e)}")
    
    # ========================================================================
    # TOPIC PARSING - Generate 15+ topic-specific slides
    # ========================================================================
    
    async def analyze_overview(self, overview_text: str = "", project_name: str = "", num_slides: int = 15) -> Dict[str, Any]:
        """Generate topic-specific slide topics"""
        safe_topic = (overview_text or project_name or "Topic").strip()
        safe_project_name = (project_name or safe_topic).strip()
        print(f"\n   Generating {num_slides} slides for: {safe_project_name[:30]}")
        
        cleaned_text = safe_topic.replace('\t', ' ').replace('\r\n', '\n')
        
        try:
            parsed = await self._parse_overview_with_llm(cleaned_text, safe_project_name, num_slides)
            slide_topics = (parsed or {}).get("slides", [])
            if not slide_topics:
                slide_topics = self._fallback_topics(safe_project_name, num_slides)
            print(f"   Generated {len(slide_topics)} topics")
            return {"success": True, "slides": slide_topics, "project_name": safe_project_name}
        except Exception as e:
            print(f"   Notice: {str(e)[:100]}, using fallback slide structure")
            return {"success": False, "slides": self._fallback_topics(safe_project_name, num_slides), "error": str(e)}
    
    async def refine_user_titles(self, titles: List[str], project_name: str) -> List[str]:
        """Refine user-provided titles to fix typos and professionalize them"""
        if not titles:
            return []
        safe_project_name = (project_name or "Presentation").strip()
        print(f"   Refining {len(titles)} user titles...")
        
        prompt = f"""I have a list of slide titles for a presentation on "{safe_project_name}".
Some might have typos or be informal. Refine them to be professional slide titles.
Keep the SAME NUMBER of slides and roughly the same meaning.

User Input: {json.dumps(titles)}

Return ONLY valid JSON: ["Title 1", "Title 2", ...]"""

        try:
            response = self._call_llm(prompt, 600)
            if response:
                match = re.search(r'\[.*\]', response, re.DOTALL)
                if match:
                    refined = json.loads(match.group(0))
                    if isinstance(refined, list) and len(refined) == len(titles):
                        return [str(t) for t in refined]
            return titles  # Fallback to original
        except Exception as e:
            print(f"Error refining titles: {e}")
            return titles

    async def _parse_overview_with_llm(self, overview_text: str, topic: str, num_slides: int) -> Dict[str, Any]:
        """Generate topic-specific slide titles using LLM"""
        safe_topic = (topic or overview_text or "Topic").strip()
        prompt = f"""You are creating a professional presentation about "{safe_topic}".

Generate EXACTLY {num_slides} slide topics that DEEPLY explore this subject.

IMPORTANT RULES:
1. First 2 slides: INTRODUCTION and ABSTRACT (always include these)
2. Middle slides: Topic-specific content that dives deep into the subject
   - For technology topics: History, How it works, Syntax/Structure, Components, Implementation, Use Cases
   - For concepts: Definition, Principles, Types, Methodology, Examples, Case Studies
   - For products/tools: Features, Architecture, Installation, Usage, Best Practices
3. Last 4 slides: ADVANTAGES, DISADVANTAGES, FUTURE SCOPE, CONCLUSION (always include these)

Now generate {num_slides} slide topics for "{safe_topic}":
Return ONLY valid JSON: {{"slides": ["SLIDE1", "SLIDE2", ...]}}"""

        try:
            response = self._call_llm(prompt, 600)
            if response:
                match = re.search(r'\{.*\}', response, re.DOTALL)
                if match:
                    result = json.loads(match.group(0))
                    slides = result.get("slides", [])
                    if isinstance(slides, list) and len(slides) >= max(3, num_slides - 2):
                        slides = slides[:num_slides]
                        slides = self._ensure_conclusion_last(slides, safe_topic)
                        return {"slides": slides}
            return {"slides": self._fallback_topics(safe_topic, num_slides)}
        except Exception:
            return {"slides": self._fallback_topics(safe_topic, num_slides)}
    
    def _ensure_conclusion_last(self, slides: List[str], topic: str) -> List[str]:
        """Ensure CONCLUSION is always the last slide"""
        if not slides:
            return [f"INTRODUCTION TO {(topic or 'TOPIC').upper()}", "ABSTRACT", "CONCLUSION"]
        
        conclusion_idx = -1
        for i, slide in enumerate(slides):
            if slide and "CONCLUSION" in str(slide).upper():
                conclusion_idx = i
                break
        
        if conclusion_idx >= 0:
            conclusion_slide = slides.pop(conclusion_idx)
            slides.append(conclusion_slide)
        else:
            slides.append("CONCLUSION")
        
        return slides

    def _fallback_topics(self, topic: str, num_slides: int = 15) -> List[str]:
        """Default topic-specific slide topics - always ends with CONCLUSION"""
        safe_topic = (topic or "TOPIC").strip().upper()
        # Fixed start slides (first 2)
        start_slides = [
            f"INTRODUCTION TO {safe_topic}",
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
        
        middle_needed = max(0, num_slides - len(start_slides) - len(end_slides))
        while len(middle_slides) < middle_needed:
            middle_slides.append(f"TOPIC {len(middle_slides) + 1}")
        
        middle_slides = middle_slides[:middle_needed]
        return start_slides + middle_slides + end_slides

    # ========================================================================
    # CONTENT GENERATION - 8 bullet points, 8-9 line paragraphs
    # ========================================================================
    
    async def generate_ppt_content(
        self,
        toc_structure: Dict[str, Any],
        code_content: str = "",
        project_name: str = "",
        content_mode: str = "cassandra"
    ) -> Dict[str, Any]:
        """Generate comprehensive PPT content"""
        safe_project_name = (project_name or "Presentation").strip()
        print(f"\n   Generating content for: {safe_project_name} (Mode: {content_mode})")
        
        generated_content = {
            "project_name": safe_project_name,
            "abstract": "",
            "chapters": []
        }
        
        # Generate Abstract
        generated_content["abstract"] = await self._generate_abstract(safe_project_name)
        
        chapters = (toc_structure or {}).get("chapters", [])
        total = len(chapters)
        for idx, chapter in enumerate(chapters, 1):
            title = (chapter.get("title") or f"Slide {idx}").strip()
            print(f"   Slide {idx}/{total}: {title}")
            
            chapter_content = {
                "chapter_number": chapter.get("chapter_number", idx),
                "title": title,
                "sections": []
            }
            
            sections = chapter.get("sections") or [{"title": title, "number": f"{idx}.1"}]
            for section in sections:
                section_title = (section.get("title") or title).strip()
                content = await self._generate_section(section_title, safe_project_name, content_mode)
                
                chapter_content["sections"].append({
                    "number": section.get("number", f"{idx}.1"),
                    "title": section_title,
                    "content": content
                })
            
            generated_content["chapters"].append(chapter_content)
        
        print(f"   Content generation complete")
        return generated_content
    
    async def _generate_abstract(self, topic: str) -> str:
        """Generate 8-9 line abstract (paragraph format)"""
        safe_topic = (topic or "Presentation").strip()
        prompt = f"""Write a comprehensive ABSTRACT about "{safe_topic}" for a professional presentation.

REQUIREMENTS:
- 8-9 sentences (180-220 words)
- Professional academic tone
- Cover: What it is, why it matters, key features, applications
- NO bullet points, just paragraph format

Write the abstract:"""

        try:
            content = self._call_llm(prompt, 400)
            cleaned = self._clean_paragraph(content)
            if cleaned and len(cleaned) > 50:
                return cleaned
        except Exception:
            pass
        return f"{safe_topic} is a significant advancement in modern technology with wide-ranging applications across various industries. It provides innovative solutions to complex problems through its unique approach and methodology. The fundamental principles underlying {safe_topic} enable efficient and effective implementation in diverse scenarios. Organizations and individuals leverage {safe_topic} to achieve better outcomes and improved performance. The field continues to evolve with new developments and innovations. Research and development efforts are driving continuous improvements. This presentation explores the key aspects, benefits, and practical applications of {safe_topic}. Understanding these concepts is essential for professionals in this domain."
    
    async def _generate_section(self, section_title: str, topic: str, content_mode: str = "cassandra") -> str:
        """Generate content based on section type and content mode"""
        safe_section = (section_title or "Section").strip()
        safe_topic = (topic or "Topic").strip()
        section_upper = safe_section.upper()
        
        use_paragraph = False
        if content_mode == 'para':
            use_paragraph = True
        elif content_mode == 'point':
            use_paragraph = False
        else:
            if any(word in section_upper for word in ["INTRODUCTION", "CONCLUSION", "ABSTRACT"]):
                use_paragraph = True
        
        if use_paragraph:
            return await self._generate_paragraph(safe_section, safe_topic)
        else:
            return await self._generate_bullets(safe_section, safe_topic)
    
    async def _generate_paragraph(self, section: str, topic: str) -> str:
        """Generate paragraph content (10-11 sentences)"""
        safe_section = (section or "Overview").strip()
        safe_topic = (topic or "Topic").strip()
        prompt = f"""Write a comprehensive paragraph about "{safe_section}" for a presentation on "{safe_topic}".

REQUIREMENTS:
- 10-11 sentences (220-280 words)
- Professional academic tone
- Informative and detailed
- NO bullet points

Write the paragraph:"""

        try:
            content = self._call_llm(prompt, 500)
            cleaned = self._clean_paragraph(content)
            if cleaned and len(cleaned) > 50:
                return cleaned
        except Exception:
            pass
        return f"This section provides a comprehensive overview of {safe_section.lower()} in the context of {safe_topic}. Understanding these fundamentals is essential for effective implementation and utilization. The concepts presented here form the foundation for advanced topics covered in subsequent sections. Practical applications and real-world examples demonstrate the relevance and importance of this subject matter. The field has evolved significantly over the years with continuous innovations. Modern approaches incorporate best practices from various domains. By mastering these concepts, professionals can leverage {safe_topic} to achieve significant improvements in their respective domains. This knowledge is crucial for anyone working in this field. The ongoing research and development continues to drive new discoveries. Organizations worldwide are investing in these technologies to stay competitive."

    async def _generate_bullets(self, section: str, topic: str) -> str:
        """Generate exactly 8 crisp bullet points"""
        safe_section = (section or "Key Aspects").strip()
        safe_topic = (topic or "Topic").strip()
        
        prompt = f"""Generate exactly 8 bullet points about "{safe_section}" for a presentation on "{safe_topic}".

CRITICAL RULES:
1. Each bullet point must be ONE clear sentence (10-15 words)
2. Each point must END with a period
3. Be specific and informative
4. NO sub-points, NO colons in the middle
5. Points must be relevant to the section topic

Now generate 8 bullet points about "{safe_section}" for "{safe_topic}":"""

        try:
            content = self._call_llm(prompt, 500)
            return self._format_bullets(content, safe_section, safe_topic)
        except Exception:
            return self._default_bullets(safe_section, safe_topic)
    
    def _format_bullets(self, content: str, section: str = "Key Aspects", topic: str = "Presentation") -> str:
        """Clean and format bullet points - ensure 8 points"""
        if not content:
            return self._default_bullets(section, topic)
            
        lines = str(content).strip().split('\n')
        bullets = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Remove markdown bold tags
            line = re.sub(r'\*\*(.+?)\*\*', r'\1', line)
            
            line_upper = line.upper()
            if line.endswith(':') or any(line_upper.startswith(word) for word in ["HERE ARE", "HERE IS", "SURE", "BELOW IS", "PRESENTATION ON"]):
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
            
            if line and line[-1] not in '.!?':
                line = line + '.'
            
            if line:
                line = line[0].upper() + line[1:]
            
            bullets.append(line)
            if len(bullets) >= 8:
                break
        
        if not bullets:
            return self._default_bullets(section, topic)
            
        while len(bullets) < 8:
            bullets.append(f"Provides essential capabilities for effective {bullets[0].split()[0].lower() if bullets else 'implementation'}.")
        
        return '\n'.join(bullets[:8])
    
    def _clean_paragraph(self, content: str) -> str:
        """Clean paragraph content safely"""
        if not content:
            return ""
        content = str(content)
        content = re.sub(r'\*\*(.+?)\*\*', r'\1', content)
        content = re.sub(r'^\#+\s+', '', content, flags=re.MULTILINE)
        content = re.sub(r'^[\-\*\•]\s+', '', content, flags=re.MULTILINE)
        content = ' '.join(content.split())
        
        if len(content) < 500:
            content = content + " This aspect plays a crucial role in the overall implementation and effectiveness of the solution. Understanding these concepts is essential for successful application. The ongoing developments in this field continue to expand possibilities. Professionals benefit greatly from staying updated with these advancements."
        
        if len(content) > 800:
            cut = content[:800].rfind('.')
            if cut > 500:
                content = content[:cut+1]
        
        return content.strip()
    
    def _default_bullets(self, section: str = "Key Aspects", topic: str = "Presentation") -> str:
        """Fallback bullet points (8 points)"""
        return f"""Provides fundamental capabilities for {topic} implementation.
Enables efficient processing and management of resources.
Supports scalable solutions for various requirements.
Ensures reliable performance across different scenarios.
Facilitates integration with existing systems and workflows.
Offers comprehensive documentation and support resources.
Delivers consistent results in production environments.
Enables rapid development and deployment cycles."""
    
    def _sanitize_text(self, text: str) -> str:
        """Sanitize unicode characters that Windows cp1252 cannot encode"""
        if not text:
            return ""
        text = str(text)
        replacements = {
            '\u2011': '-',
            '\u2012': '-',
            '\u2013': '-',
            '\u2014': '-',
            '\u2015': '-',
            '\u2018': "'",
            '\u2019': "'",
            '\u201c': '"',
            '\u201d': '"',
            '\u2022': '*',
            '\u2026': '...',
            '\u00a0': ' ',
        }
        for char, replacement in replacements.items():
            text = text.replace(char, replacement)
        return text.encode('ascii', errors='ignore').decode('ascii')
    
    def _call_nvidia_direct(self, prompt: str, api_key: str, max_tokens: int = 500) -> str:
        """Direct HTTP call to NVIDIA NIM API without requiring third-party SDK quirks"""
        model = os.getenv("PPT_NVIDIA_MODEL", "meta/llama-3.2-11b-vision-instruct")
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        payload = {
            "model": model,
            "messages": [{"role": "user", "content": prompt}],
            "max_tokens": max_tokens,
            "temperature": 0.7
        }
        res = requests.post(
            "https://integrate.api.nvidia.com/v1/chat/completions",
            headers=headers,
            json=payload,
            timeout=30
        )
        if res.status_code == 200:
            data = res.json()
            choices = data.get("choices", [])
            if choices and "message" in choices[0]:
                content = choices[0]["message"].get("content", "")
                return self._sanitize_text(content or "")
        raise Exception(f"NVIDIA API status {res.status_code}: {res.text[:150]}")

    def _call_llm(self, prompt: str, max_tokens: int = 500) -> str:
        """Call LLM API with active models and multi-model fallback support"""
        # Primary API: Groq
        if self.api_type == "groq" and self.api_client:
            # Active working models on Groq in priority order
            groq_models = [
                os.getenv("PPT_GROQ_MODEL", "qwen/qwen3.8-27b"),
                "openai/gpt-oss-120b",
                "openai/gpt-oss-20b",
                "groq/compound-mini"
            ]
            for model_name in groq_models:
                try:
                    response = self.api_client.chat.completions.create(
                        messages=[{"role": "user", "content": prompt}],
                        model=model_name,
                        max_tokens=max_tokens,
                        temperature=0.7
                    )
                    if response.choices and response.choices[0].message and response.choices[0].message.content:
                        return self._sanitize_text(response.choices[0].message.content)
                except Exception as g_err:
                    logger.warning(f"Groq model {model_name} failed: {str(g_err)[:100]}")
                    continue
        
        # Primary API: NVIDIA NIM
        elif self.api_type == "nvidia":
            if self.nvidia_api_key:
                try:
                    return self._call_nvidia_direct(prompt, self.nvidia_api_key, max_tokens)
                except Exception as n_err:
                    logger.warning(f"NVIDIA direct call failed: {str(n_err)[:100]}")
            
            if self.api_client:
                try:
                    response = self.api_client.chat.completions.create(
                        messages=[{"role": "user", "content": prompt}],
                        model=os.getenv("PPT_NVIDIA_MODEL", "meta/llama-3.2-11b-vision-instruct"),
                        max_tokens=max_tokens,
                        temperature=0.7
                    )
                    if response.choices and response.choices[0].message and response.choices[0].message.content:
                        return self._sanitize_text(response.choices[0].message.content)
                except Exception as n_err:
                    logger.warning(f"NVIDIA SDK call failed: {str(n_err)[:100]}")

        elif self.api_type == "cerebras" and self.api_client:
            try:
                response = self.api_client.chat.completions.create(
                    messages=[{"role": "user", "content": prompt}],
                    model=os.getenv("PPT_CEREBRAS_MODEL", "llama-3.3-70b"),
                    max_tokens=max_tokens,
                    temperature=0.7
                )
                if response.choices and response.choices[0].message and response.choices[0].message.content:
                    return self._sanitize_text(response.choices[0].message.content)
            except Exception as c_err:
                logger.warning(f"Cerebras call failed: {str(c_err)[:100]}")

        # Secondary Fallback: try NVIDIA NIM if Groq failed and we have NVIDIA key
        nvidia_key = self.nvidia_api_key or os.getenv("PPT_NVIDIA_API_KEY") or os.getenv("NVIDIA_API_KEY")
        if nvidia_key:
            try:
                logger.info("Falling back to NVIDIA NIM API...")
                return self._call_nvidia_direct(prompt, nvidia_key, max_tokens)
            except Exception as fb_err:
                logger.warning(f"Fallback to NVIDIA NIM failed: {str(fb_err)[:100]}")

        raise Exception("All LLM API calls failed or no valid API key configured.")

    # ========================================================================
    # REFINE SLIDE - Regenerate content for a specific slide
    # ========================================================================
    
    async def refine_slide(self, slide_title: str, current_content: str, topic: str, style: str = "bullet") -> str:
        """
        Refine/regenerate content for a specific slide.
        """
        safe_title = (slide_title or "Slide").strip()
        safe_content = (current_content or "").strip()
        safe_topic = (topic or "Topic").strip()
        print(f"   Refining slide: {safe_title} (style: {style})")
        
        if style == "paragraph":
            return await self._refine_paragraph(safe_title, safe_content, safe_topic)
        else:
            return await self._refine_bullets(safe_title, safe_content, safe_topic)
    
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
            cleaned = self._clean_paragraph(content)
            if cleaned and len(cleaned) > 50:
                return cleaned
        except Exception:
            pass
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
            return self._format_bullets(content, slide_title, topic)
        except Exception:
            return current_content

