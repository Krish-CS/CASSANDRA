"""
Cassandra AI PPT Generator
Simple Flask backend for topic-based PPT generation
"""

from flask import Flask, render_template, jsonify, request, send_file, redirect, url_for
import os
import time
import threading
from pathlib import Path
from datetime import datetime

app = Flask(__name__)

# Directories
DATA_DIR = Path("data")
OUTPUT_DIR = Path("output")
DATA_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# File Cleanup Configuration (Minutes)
FILE_LIFETIME_MINUTES = 30

def cleanup_old_files():
    """Background thread to delete generated PPT files older than FILE_LIFETIME_MINUTES"""
    while True:
        try:
            now = time.time()
            for filename in os.listdir(OUTPUT_DIR):
                if filename.endswith(".pptx") and filename != "template_blank.pptx":
                    filepath = (OUTPUT_DIR / filename)
                    if os.path.exists(filepath):
                        file_age_minutes = (now - os.path.getmtime(filepath)) / 60
                        if file_age_minutes > FILE_LIFETIME_MINUTES:
                            os.remove(filepath)
                            print(f"🧹 Auto-cleaned old file: {filename} (Age: {int(file_age_minutes)}m)")
        except Exception as e:
            print(f"⚠️ Cleanup thread error: {e}")
        time.sleep(600)  # Check every 10 minutes

# Start the background cleanup thread
cleanup_thread = threading.Thread(target=cleanup_old_files, daemon=True)
cleanup_thread.start()

# ============================================================================
# ROUTES
# ============================================================================

@app.route('/')
def login():
    """Serve login page"""
    return render_template('login.html')

@app.route('/dashboard')
def dashboard():
    """Serve dashboard page"""
    return render_template('dashboard.html')

@app.route('/preview')
def preview():
    """Serve preview page"""
    topic = request.args.get('topic', 'Sample Topic')
    template = request.args.get('template', 'modern')
    template_url = request.args.get('templateUrl', '')
    mode = request.args.get('mode', 'flash')
    content_mode = request.args.get('content_mode', 'cassandra')
    return render_template('preview.html', topic=topic, template=template, template_url=template_url, mode=mode, content_mode=content_mode)

@app.route('/thank-you')
def thank_you():
    """Serve thank you page after PPT generation"""
    return render_template('thankyou.html')

@app.route('/logout')
def logout():
    """Logout - redirect to login page. Client-side localStorage will be cleared."""
    return redirect(url_for('login'))

@app.route('/ping')
def ping():
    """Health check endpoint to prevent Render cold starts (e.g., via UptimeRobot)"""
    return jsonify({"status": "ok", "message": "I'm alive!"})



# ============================================================================
# API ENDPOINTS
# ============================================================================

@app.route('/api/templates')
def get_templates():
    """
    Fetch background templates from Pexels API
    Query params:
        - color: Filter by color (pink, blue, violet, etc.)
        - query: Search query (default: abstract background)
        - count: Number of images (default: 12)
    """
    try:
        from pexels_service import fetch_backgrounds
        
        color = request.args.get('color', 'pink')
        query = request.args.get('query', 'abstract background')
        count = int(request.args.get('count', 12))
        
        templates = fetch_backgrounds(color=color, query=query, per_page=count)
        
        return jsonify({
            "success": True,
            "color": color,
            "count": len(templates),
            "templates": templates
        })
        
    except Exception as e:
        print(f"Error fetching templates: {e}")
        return jsonify({"success": False, "error": str(e)}), 500


@app.route('/api/template-colors')
def get_template_colors():
    """Get list of supported template colors"""
    try:
        from pexels_service import get_supported_colors
        colors = get_supported_colors()
        return jsonify({"success": True, "colors": colors})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/pexels/thank-you-images')
def get_thank_you_images():
    """Fetch Thank You themed images from Pexels"""
    try:
        from pexels_service import search_thank_you_images
        images = search_thank_you_images(per_page=80, max_results=100)
        return jsonify({"success": True, "count": len(images), "images": images})
    except Exception as e:
        print(f"Error fetching thank you images: {e}")
        return jsonify({"success": False, "error": str(e)}), 500


@app.route('/api/generate-topics', methods=['POST'])
def generate_topics():
    """
    Generate slide topics from a main topic using AI
    This is what makes Cassandra different - takes just a topic,
    AI auto-generates the slide structure
    """
    try:
        data = request.get_json() or request.form
        topic = data.get('topic', '')
        num_slides = int(data.get('num_slides', 8))
        
        if not topic:
            return jsonify({"success": False, "error": "Topic is required"}), 400
        
        # Try to use AI for topic generation
        try:
            from presentation.slide_generator import SlideGenerator
            generator = SlideGenerator()
            
            # Generate slide topics based on main topic
            import asyncio
            result = asyncio.run(generator.analyze_overview(
                overview_text=topic,
                project_name=topic
            ))
            
            slides = result.get("slides", [])
            
            if not slides:
                # Use default topic structure
                slides = _get_default_slides(topic)
                
        except Exception as e:
            print(f"AI not available: {e}, using default slides")
            slides = _get_default_slides(topic)
        
        return jsonify({
            "success": True,
            "topic": topic,
            "slides": slides
        })
        
    except Exception as e:
        print(f"Error generating topics: {e}")
        return jsonify({"success": False, "error": str(e)}), 500


@app.route('/api/generate-preview', methods=['POST'])
def generate_preview():
    """
    Generate AI-powered slide content for Decide Mode preview.
    Returns slides with titles AND content generated by Groq.
    """
    try:
        data = request.get_json() or request.form
        topic = data.get('topic', '')
        num_slides = int(data.get('num_slides', 15))
        content_mode = data.get('content_mode', 'cassandra')
        
        # Clamp slide count between 10 and 30
        num_slides = max(10, min(30, num_slides))
        
        if not topic:
            return jsonify({"success": False, "error": "Topic is required"}), 400
        
        print(f"\n" + "="*60)
        print(f"   DECIDE MODE - AI PREVIEW GENERATION")
        print(f"   Topic: {topic}")
        print(f"   Slides: {num_slides}")
        print("="*60)
        
        try:
            from presentation.slide_generator import SlideGenerator
            generator = SlideGenerator()
            import asyncio
            
            # CHECK FOR CUSTOM TITLES (Decide Mode - Create with Mine)
            user_titles = data.get('user_titles')
            
            if user_titles and isinstance(user_titles, list) and len(user_titles) > 0:
                print(f"   Using {len(user_titles)} user-provided titles")
                # Refine titles to fix typos
                slide_titles = asyncio.run(generator.refine_user_titles(user_titles, topic))
                print(f"   Refined titles: {slide_titles}")
            else:
                # Step 1: Generate slide topics using AI
                import asyncio
                topics_result = asyncio.run(generator.analyze_overview(
                    overview_text=topic,
                    project_name=topic,
                    num_slides=num_slides
                ))
                slide_titles = topics_result.get("slides", _get_default_slides(topic))

            
            # Step 2: Build TOC structure
            toc_structure = {
                'project_name': topic,
                'chapters': []
            }
            
            for idx, slide_topic in enumerate(slide_titles):
                toc_structure['chapters'].append({
                    'chapter_number': idx + 1,
                    'title': slide_topic,
                    'sections': [{'number': f'{idx + 1}.1', 'title': slide_topic}]
                })
            
            # Step 3: Generate content for each slide
            generated_content = asyncio.run(generator.generate_ppt_content(
                toc_structure=toc_structure,
                code_content="",  # Topic-based, no code
                project_name=topic,
                content_mode=content_mode
            ))
            
            # Convert to preview format - all slides come from LLM
            slides = []
            
            # Add chapter slides (LLM already generates proper topics including abstract/intro/conclusion)
            for chapter in generated_content.get("chapters", []):
                chapter_title = chapter.get("title", "")
                for section in chapter.get("sections", []):
                    content = section.get("content", "")
                    if content:
                        # Determine type based on title or content_mode
                        if content_mode == 'para':
                            slide_type = "paragraph"
                        elif content_mode == 'point':
                            slide_type = "bullet"
                        else:
                            # Cassandra Mode (Auto)
                            title_upper = chapter_title.upper()
                            is_paragraph = any(word in title_upper for word in ["INTRODUCTION", "CONCLUSION", "ABSTRACT", "SUMMARY"])
                            slide_type = "paragraph" if is_paragraph else "bullet"
                        
                        # Add bullet symbols to bullet content for preview display
                        if slide_type == "bullet":
                            lines = content.strip().split('\n')
                            formatted_lines = []
                            for line in lines:
                                line = line.strip()
                                if line and not line.startswith('➣'):
                                    line = f"➣ {line}"
                                formatted_lines.append(line)
                            content = '\n'.join(formatted_lines)
                        
                        slides.append({
                            "title": chapter_title,
                            "content": content,
                            "type": slide_type
                        })

            
            print(f"✅ Generated {len(slides)} slides with AI content")
            
            return jsonify({
                "success": True,
                "topic": topic,
                "slides": slides,
                "ai_generated": True
            })
            
        except Exception as e:
            print(f"⚠️ AI generation failed: {e}, using fallback")
            import traceback
            traceback.print_exc()
            
            # Fallback to default content
            return jsonify({
                "success": True,
                "topic": topic,
                "slides": _get_default_preview_slides(topic),
                "ai_generated": False
            })
        
    except Exception as e:
        print(f"Error generating preview: {e}")
        return jsonify({"success": False, "error": str(e)}), 500


def _get_default_preview_slides(topic):
    """Generate default preview slides with content"""
    return [
        {"title": f"Introduction to {topic}", "content": f"{topic} represents a significant advancement in its field. It encompasses various methodologies and approaches that have evolved over time. The fundamental principles underlying {topic} provide a strong foundation for understanding its applications.", "type": "paragraph"},
        {"title": f"Overview of {topic}", "content": f"• {topic} is a comprehensive framework that addresses modern challenges.\n• It integrates multiple components to provide effective solutions.\n• The core principles are designed for scalability and efficiency.\n• Understanding the fundamentals enables better implementation.", "type": "bullet"},
        {"title": "Key Concepts", "content": f"• Foundation principles form the backbone of implementation.\n• Core terminology and definitions establish clear understanding.\n• Theoretical frameworks guide practical applications.\n• Component relationships enable system integration.", "type": "bullet"},
        {"title": "Core Principles", "content": f"• Modularity ensures flexible component design.\n• Scalability considerations enable growth and adaptation.\n• Efficiency optimization reduces resource consumption.\n• Reliability measures guarantee consistent performance.", "type": "bullet"},
        {"title": "Applications & Use Cases", "content": f"• Industry applications demonstrate practical value.\n• Research applications advance scientific understanding.\n• Everyday use cases show accessibility to users.\n• Future possibilities reveal untapped potential.", "type": "bullet"},
        {"title": "Advantages", "content": f"• Enhanced efficiency improves overall performance.\n• Cost-effectiveness reduces operational expenses.\n• Scalability allows adaptation to requirements.\n• User-friendly design ensures easy adoption.", "type": "bullet"},
        {"title": "Disadvantages", "content": f"• Initial implementation may require investment.\n• Learning curve can be steep for complex uses.\n• Compatibility issues may arise with legacy systems.\n• Maintenance needs ongoing attention.", "type": "bullet"},
        {"title": "Limitations", "content": f"• Technical constraints may limit certain applications.\n• Resource requirements can be substantial.\n• Knowledge gaps exist in specific areas.\n• Environmental factors may affect performance.", "type": "bullet"},
        {"title": "Future Scope", "content": f"• Emerging trends indicate growing adoption.\n• Research explores new application domains.\n• Technological advances enable enhanced capabilities.\n• Industry evolution creates new opportunities.", "type": "bullet"},
        {"title": "Conclusion", "content": f"In conclusion, {topic} offers significant value across multiple dimensions. The advantages clearly outweigh the limitations when proper implementation strategies are followed. Continued research and practical application will unlock further potential.", "type": "paragraph"}
    ]


@app.route('/api/refine-slide', methods=['POST'])
def refine_slide():
    """
    Refine/regenerate content for a specific slide.
    Preserves the style (paragraph or bullet points).
    """
    try:
        data = request.get_json() or request.form
        topic = data.get('topic', '')
        slide_title = data.get('slide_title', '')
        current_content = data.get('current_content', '')
        style = data.get('style', 'bullet')  # 'paragraph' or 'bullet'
        bullet_symbol = data.get('bullet_symbol', '➣')  # Use selected bullet style
        
        if not topic or not slide_title:
            return jsonify({"success": False, "error": "Topic and slide title are required"}), 400
        
        print(f"\n   REFINE SLIDE")
        print(f"   Title: {slide_title}")
        print(f"   Style: {style}")
        print(f"   Bullet: {bullet_symbol}")

        
        try:
            from presentation.slide_generator import SlideGenerator
            import asyncio
            
            generator = SlideGenerator()
            new_content = asyncio.run(generator.refine_slide(
                slide_title=slide_title,
                current_content=current_content,
                topic=topic,
                style=style
            ))
            
            # Add bullet symbols if bullet style - strip any existing bullets first
            if style == "bullet":
                lines = new_content.strip().split('\n')
                formatted_lines = []
                for line in lines:
                    line = line.strip()
                    if not line:
                        continue
                    # Strip ALL existing bullet markers first
                    import re
                    line = re.sub(r'^[\s\-\*\•\➢\➣\➤\►\▶\→\>\d\.\)\:]+\s*', '', line)
                    line = line.strip()
                    if line:
                        line = f"{bullet_symbol} {line}"  # Use selected bullet symbol
                        formatted_lines.append(line)
                new_content = '\n'.join(formatted_lines)


            
            print(f"   Refine complete")
            
            return jsonify({
                "success": True,
                "content": new_content,
                "style": style
            })
            
        except Exception as e:
            print(f"   Refine error: {e}")
            import traceback
            traceback.print_exc()
            return jsonify({"success": False, "error": str(e)}), 500
        
    except Exception as e:
        print(f"Error refining slide: {e}")
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/generate-ppt', methods=['POST'])
def generate_ppt():
    """
    Generate PPT from topic
    Cassandra flow: Topic → AI generates slides → Creates PPT
    
    For Flash Mode: Returns PPT file directly
    For Decide Mode: Returns JSON with slide data for preview
    """
    try:
        data = request.get_json() or request.form
        topic = data.get('topic', '')
        slides_data = data.get('slides', [])
        template_style = data.get('template', 'modern')
        template_url = data.get('templateUrl', '')  # Pexels background URL
        mode = data.get('mode', 'flash')  # 'flash' or 'decide'
        
        if not topic:
            return jsonify({"success": False, "error": "Topic is required"}), 400
        
        # Generate slide content with AI
        try:
            from presentation.slide_generator import SlideGenerator
            generator = SlideGenerator()
            
            # Check if slides_data contains full content (from Decide Mode) or just titles
            if slides_data and isinstance(slides_data[0], dict) and 'content' in slides_data[0]:
                # Decide Mode: user has provided full content
                print("   📝 Using user-edited slide content from preview...")
                generated_content = {
                    "project_name": topic,
                    "abstract": "",  # Will be first slide content if intro-type
                    "chapters": []
                }
                
                for idx, slide in enumerate(slides_data):
                    slide_title = slide.get('title', f'Slide {idx + 1}')
                    slide_content = slide.get('content', '')
                    slide_type = slide.get('type', 'bullet')
                    
                    # Add ALL slides to chapters - preserve slide order from preview
                    generated_content["chapters"].append({
                        "chapter_number": idx + 1,
                        "title": slide_title.upper(),
                        "sections": [{
                            "number": f"{idx + 1}.1",
                            "title": slide_title,
                            "content": slide_content,
                            "style": slide_type  # Preserve the style
                        }]
                    })

            else:
                # Flash Mode or just titles provided - generate content
                if not slides_data:
                    import asyncio
                    result = asyncio.run(generator.analyze_overview(
                        overview_text=topic,
                        project_name=topic
                    ))
                    slides_data = result.get("slides", _get_default_slides(topic))
                
                # Create TOC structure for content generation
                toc_structure = {
                    'project_name': topic,
                    'chapters': [],
                    'slides': slides_data
                }
                
                for idx, slide_topic in enumerate(slides_data):
                    toc_structure['chapters'].append({
                        'chapter_number': idx + 1,
                        'title': slide_topic,
                        'sections': [{
                            'number': f'{idx + 1}.1',
                            'title': slide_topic
                        }]
                    })
                
                # Generate content (no code analysis for Cassandra)
                import asyncio
                generated_content = asyncio.run(generator.generate_ppt_content(
                    toc_structure=toc_structure,
                    code_content="",  # No code for Cassandra
                    project_name=topic
                ))
            
        except Exception as e:
            print(f"AI content generation failed: {e}")
            # Fallback: create basic content
            if slides_data and isinstance(slides_data[0], dict):
                # User-edited content fallback
                generated_content = {
                    "project_name": topic,
                    "abstract": slides_data[0].get('content', '') if slides_data else '',
                    "chapters": [
                        {
                            "chapter_number": idx + 1,
                            "title": s.get('title', '').upper(),
                            "sections": [{
                                "number": f"{idx + 1}.1",
                                "title": s.get('title', ''),
                                "content": s.get('content', '')
                            }]
                        }
                        for idx, s in enumerate(slides_data[1:], 1)
                    ]
                }
            else:
                generated_content = _create_fallback_content(topic, slides_data or _get_default_slides(topic))
        
        # Build PPT
        try:
            from ppt_generator import PPTGenerator
            ppt_gen = PPTGenerator()
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_filename = f"cassandra_{topic.replace(' ', '_')}_{timestamp}.pptx"
            output_path = OUTPUT_DIR / output_filename
            
            # Create a simple template if none exists
            template_path = _create_simple_template()
            
            # Extract thank you image URL if provided
            thank_you_image_url = data.get('thankYouImageUrl', '')
            
            ppt_gen.generate_ppt(
                template_path=str(template_path),
                project_name=topic,
                generated_content=generated_content,
                sections_config={
                    "sections": {}, 
                    "bullet_symbol": data.get('bulletSymbol', '➣'),
                    "background_url": data.get('templateUrl', ''),
                    "thank_you_image_url": thank_you_image_url
                },
                output_path=str(output_path)
            )
            
            if output_path.exists():
                @app.after_request
                def remove_file(response):
                    try:
                        os.remove(output_path)
                        print(f"🧹 Cleaned up file: {output_path}")
                    except Exception as error:
                        print(f"⚠️ Error removing file: {error}")
                    return response

                return send_file(
                    str(output_path),
                    mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                    as_attachment=True,
                    download_name=output_filename
                )
            else:
                return jsonify({"success": False, "error": "Failed to create PPT"}), 500
                
        except ImportError as e:
            print(f"PPT generator not available: {e}")
            return jsonify({
                "success": False, 
                "error": "PPT generator not fully configured. Missing dependencies.",
                "slides": slides_data
            }), 500
            
    except Exception as e:
        print(f"Error generating PPT: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"success": False, "error": str(e)}), 500


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def _get_default_slides(topic):
    """Generate default slide structure for a topic (10 slides)"""
    return [
        f"Introduction to {topic}",
        f"Overview of {topic}",
        "Key Concepts",
        "Core Principles",
        "Applications & Use Cases",
        "Advantages",
        "Disadvantages",
        "Limitations",
        "Future Scope",
        "Conclusion"
    ]


def _create_fallback_content(topic, slides):
    """Create proper content with 4 bullet points per slide (except intro/conclusion as paragraphs)"""
    content = {
        "project_name": topic,
        "abstract": f"This presentation provides a comprehensive overview of {topic}, covering its key concepts, applications, advantages, limitations, and future scope. Understanding {topic} is essential for professionals and enthusiasts seeking to leverage its potential in various domains.",
        "chapters": []
    }
    
    # Content templates for different slide types
    slide_content_templates = {
        "introduction": f"{topic} represents a significant advancement in its field. It encompasses various methodologies and approaches that have evolved over time. The fundamental principles underlying {topic} provide a strong foundation for understanding its applications. This presentation explores the key aspects that make {topic} relevant in today's context.",
        
        "overview": f"• {topic} is a comprehensive framework that addresses modern challenges.\n• It integrates multiple components to provide effective solutions.\n• The core principles are designed for scalability and efficiency.\n• Understanding the fundamentals enables better implementation strategies.",
        
        "key concepts": f"• Foundation principles form the backbone of {topic} implementation.\n• Core terminology and definitions establish clear understanding.\n• Theoretical frameworks guide practical applications.\n• Relationship between components enables system integration.",
        
        "core principles": f"• Principle of modularity ensures flexible component design.\n• Scalability considerations enable growth and adaptation.\n• Efficiency optimization reduces resource consumption.\n• Reliability measures guarantee consistent performance.",
        
        "applications": f"• Industry applications demonstrate practical value in real scenarios.\n• Research applications advance scientific understanding.\n• Everyday use cases show accessibility to general users.\n• Future possibilities reveal untapped potential areas.",
        
        "advantages": f"• Enhanced efficiency improves overall system performance.\n• Cost-effectiveness reduces operational expenses significantly.\n• Scalability allows adaptation to varying requirements.\n• User-friendly design ensures easy adoption and learning.",
        
        "disadvantages": f"• Initial implementation may require significant investment.\n• Learning curve can be steep for complex applications.\n• Compatibility issues may arise with legacy systems.\n• Maintenance requirements need ongoing attention and resources.",
        
        "limitations": f"• Technical constraints may limit certain applications.\n• Resource requirements can be substantial for large-scale use.\n• Knowledge gaps exist in specific implementation areas.\n• Environmental factors may affect performance outcomes.",
        
        "future scope": f"• Emerging trends indicate growing adoption across sectors.\n• Research directions explore new application domains.\n• Technological advances enable enhanced capabilities.\n• Industry evolution creates new opportunities for innovation.",
        
        "conclusion": f"In conclusion, {topic} offers significant value across multiple dimensions. The advantages clearly outweigh the limitations when proper implementation strategies are followed. As technology continues to evolve, {topic} will play an increasingly important role in shaping future developments. Continued research and practical application will unlock further potential."
    }
    
    for idx, slide_topic in enumerate(slides):
        slide_lower = slide_topic.lower()
        
        # Find matching content template
        content_text = None
        for key, template in slide_content_templates.items():
            if key in slide_lower:
                content_text = template
                break
        
        # Default content with 4 points if no match
        if not content_text:
            content_text = f"• Key aspect of {slide_topic} relates to core functionality.\n• Implementation involves specific methodologies and approaches.\n• Benefits include improved efficiency and effectiveness.\n• Future developments will enhance current capabilities."
        
        content["chapters"].append({
            "chapter_number": idx + 1,
            "title": slide_topic.upper(),
            "sections": [{
                "number": f"{idx + 1}.1",
                "title": slide_topic,
                "content": content_text
            }]
        })
    
    return content


def _create_simple_template():
    """Create a simple blank PPT template if needed"""
    template_path = DATA_DIR / "template_blank.pptx"
    
    if not template_path.exists():
        try:
            from pptx import Presentation
            prs = Presentation()
            prs.save(str(template_path))
        except:
            pass
    
    return template_path


# ============================================================================
# RUN
# ============================================================================

if __name__ == '__main__':
    print("\n" + "="*60)
    print("🌸 CASSANDRA AI PPT GENERATOR")
    print("="*60)
    print("\n🌐 Open: http://localhost:5000")
    print("💡 Just enter a topic and let AI generate your PPT!")
    print("\n" + "="*60 + "\n")
    
    app.run(debug=True, host='0.0.0.0', port=5000)