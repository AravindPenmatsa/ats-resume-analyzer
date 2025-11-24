import logging
import os
import sys
from datetime import datetime
from pathlib import Path
from jinja2 import Template
from app.core.config import GENERATED_DIR
from app.services.resume_service import parse_resume_to_structure, clean_resume_formatting_issues, clean_professional_experience_bullets

# ReportLab imports
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY

logger = logging.getLogger("app")

# Force ReportLab for consistent PDF formatting across environments
USE_REPORTLAB_ONLY = True

def generate_resume_html(resume_data: dict) -> str:
    template = """
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            body { font-family: Calibri, Arial, sans-serif; margin: 0.7in; line-height: 1.2; font-size: 11pt; }
            .header { }
            .header h1 { margin: 0; font-size: 20pt; }
            .header h2 { margin: 0; font-size: 14pt; font-weight: normal; }
            .header p { margin: 2px 0; font-size: 10pt; }
            .section { margin-top: 15px; }
            .section-title { font-size: 14pt; font-weight: bold; color: #4F81BD; border-bottom: 1px solid #B0C4DE; padding-bottom: 3px; margin-bottom: 8px; }
            
            .experience-entry { margin-bottom: 18px; }
            .experience-header { 
                margin: 0; 
                padding: 0; 
                display: flex; 
                justify-content: space-between; 
                width: 100%; 
            }
            .experience-header .company-location { 
                margin-right: auto; 
                padding-right: 20px; 
                font-weight: bold; 
            }
            .experience-header .duration { 
                text-align: right; 
                flex-shrink: 0; 
                white-space: nowrap; 
                font-weight: bold;
                letter-spacing: normal;
                word-spacing: normal;
                min-width: 150px;
            }
            .experience-role { 
                margin: 5px 0 5px 0; 
                font-style: italic; 
                font-weight: bold; 
            }
            .project-company {
                font-weight: bold;
                margin: 0;
                padding: 0;
            }
            .project-duration {
                margin: 2px 0;
                padding: 0;
                font-size: 10pt;
                font-weight: bold;
            }
            .project-role {
                font-weight: bold;
                font-style: italic;
                margin: 2px 0 8px 0;
                padding: 0;
            }
            .bullet-list { list-style-position: outside; padding-left: 22px; margin: 0; }
            .bullet-list li { margin-bottom: 6px; }
            .project-environment { margin-top: 8px; padding: 4px; background-color: #F2F2F2; font-size: 9pt; font-style: italic; }
            .project-environment b { font-style: normal; }
            .plain-text-block-content { 
                margin-top: 0; 
                margin-bottom: 10px; 
                white-space: pre-wrap; 
                line-height: 1.3; 
                padding-left: 0;
                text-align: left;
            }
            .skills-table { 
                width: 100%; 
                border-collapse: collapse; 
                margin-bottom: 10px; 
            }
            .skills-table td { 
                padding: 3px 5px; 
                vertical-align: top; 
            }
            .skills-category { 
                font-weight: bold; 
                width: 25%; 
                white-space: nowrap; 
            }
            .skills-list { 
                width: 75%; 
            }
        </style>
    </head>
    <body>
        <div class="header">
            <h1>{{ name }}</h1>
            <h2>{{ title }}</h2>
            <p>{{ location }} | {{ phone }} | {{ email }}</p>
            <p><a href="{{ linkedin }}">{{ linkedin }}</a></p>
        </div>

        {% if summary %}
        <div class="section">
            <div class="section-title">PROFESSIONAL SUMMARY</div>
            <ul class="bullet-list">
                {% for bullet in summary %}
                <li>{{ bullet }}</li>
                {% endfor %}
            </ul>
        </div>
        {% endif %}

        {% if skills %}
        <div class="section">
            <div class="section-title">TECHNICAL SKILLS</div>
            <table class="skills-table">
                {% for category, skill_list in skills.items() %}
                <tr>
                    <td class="skills-category">{{ category }}</td>
                    <td class="skills-list">{{ skill_list }}</td>
                </tr>
                {% endfor %}
            </table>
        </div>
        {% endif %}

        {% if experience %}
        <div class="section">
            <div class="section-title">PROFESSIONAL EXPERIENCE</div>
            {% for job in experience %}
            <div class="experience-entry">
                <div class="experience-header">
                    <span class="company-location">{{ job.company }} | {{ job.location }}</span>
                    <span class="duration">{{ job.duration }}</span>
                </div>
                <div class="experience-role">{{ job.role }}</div>
                <ul class="bullet-list">
                    {% for bullet in job.responsibilities %}
                    <li>{{ bullet }}</li>
                    {% endfor %}
                </ul>
            </div>
            {% endfor %}
        </div>
        {% endif %}

        {% if projects %}
        <div class="section">
            <div class="section-title">PROJECTS</div>
            {% for project in projects %}
            <div class="experience-entry">
                <div class="project-company">{{ project.name }}</div>
                <div class="project-duration">{{ project.duration }}</div>
                <div class="project-role">Role: {{ project.role }}</div>
                <ul class="bullet-list">
                    {% for bullet in project.details %}
                    <li>{{ bullet }}</li>
                    {% endfor %}
                </ul>
                {% if project.environment %}
                <div class="project-environment"><b>Environment:</b> {{ project.environment }}</div>
                {% endif %}
            </div>
            {% endfor %}
        </div>
        {% endif %}

        {% if education %}
        <div class="section">
            <div class="section-title">EDUCATION</div>
            {% for edu in education %}
            <div class="experience-entry">
                <div class="experience-header">
                    <span class="company-location">{{ edu.degree }}</span>
                    <span class="duration">{{ edu.year }}</span>
                </div>
                <div>{{ edu.university }}</div>
            </div>
            {% endfor %}
        </div>
        {% endif %}
        
        {% if certifications %}
        <div class="section">
            <div class="section-title">CERTIFICATIONS</div>
            <ul class="bullet-list">
                {% for cert in certifications %}
                <li>{{ cert }}</li>
                {% endfor %}
            </ul>
        </div>
        {% endif %}

    </body>
    </html>
    """
    
    t = Template(template)
    return t.render(resume_data)

def generate_formatted_resume_pdf(filename: str, enhanced_text: str, user_info: dict) -> str:
    """
    Generate a professionally formatted resume PDF with enhanced content.
    """
    
    # Parse the enhanced text into structured data
    resume_data = parse_resume_to_structure(enhanced_text, user_info)
    
    # Generate HTML content first
    html_content = generate_resume_html(resume_data)
    
    # CRITICAL: Apply comprehensive cleanup to the HTML content to fix all formatting issues
    logger.info("🧹 Applying comprehensive formatting fixes to HTML content...")
    html_content = clean_resume_formatting_issues(html_content)
    html_content = clean_professional_experience_bullets(html_content)
    
    logger.info("Attempting to generate PDF resume.")
    # Set library path for macOS if necessary
    if 'darwin' in str(sys.platform):
        os.environ['DYLD_LIBRARY_PATH'] = '/opt/homebrew/lib:' + os.environ.get('DYLD_LIBRARY_PATH', '')
    
    output_path = os.path.join(GENERATED_DIR, f"{os.path.splitext(filename)[0]}_formatted.pdf")
    
    # Determine PDF generation method based on configuration
    if USE_REPORTLAB_ONLY:
        # Directly use ReportLab for PDF generation (pure Python, no external deps)
        # Create PDF using ReportLab with enhanced formatting
        doc = SimpleDocTemplate(output_path, pagesize=letter, 
                              rightMargin=0.7*inch, leftMargin=0.7*inch,
                              topMargin=0.7*inch, bottomMargin=0.7*inch)
        
        styles = getSampleStyleSheet()
        
        # Enhanced custom styles to match WeasyPrint
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Title'],
            fontSize=20,
            fontName='Helvetica-Bold',
            spaceAfter=2,
            alignment=TA_LEFT,
            textColor=colors.black
        )
        
        subtitle_style = ParagraphStyle(
            'CustomSubtitle',
            parent=styles['Normal'],
            fontSize=14,
            fontName='Helvetica',
            spaceAfter=2,
            alignment=TA_LEFT,
            textColor=colors.black
        )
        
        contact_style = ParagraphStyle(
            'ContactStyle',
            parent=styles['Normal'],
            fontSize=10,
            fontName='Helvetica',
            spaceAfter=15,
            alignment=TA_LEFT,
            textColor=colors.black
        )
        
        section_title_style = ParagraphStyle(
            'SectionTitle',
            parent=styles['Heading2'],
            fontSize=14,
            fontName='Helvetica-Bold',
            spaceBefore=15,
            spaceAfter=8,
            textColor=colors.HexColor('#4F81BD'),
            borderWidth=1,
            borderColor=colors.HexColor('#B0C4DE'),
            borderPadding=3
        )
        
        company_style = ParagraphStyle(
            'CompanyStyle',
            parent=styles['Normal'],
            fontSize=11,
            fontName='Helvetica-Bold',
            spaceBefore=8,
            spaceAfter=2,
            alignment=TA_LEFT
        )
        
        duration_style = ParagraphStyle(
            'DurationStyle',
            parent=styles['Normal'],
            fontSize=10,
            fontName='Helvetica-Bold',
            spaceAfter=2,
            alignment=TA_LEFT
        )
        
        role_style = ParagraphStyle(
            'RoleStyle',
            parent=styles['Normal'],
            fontSize=11,
            fontName='Helvetica-BoldOblique',
            spaceAfter=8,
            alignment=TA_LEFT
        )
        
        bullet_style = ParagraphStyle(
            'BulletStyle',
            parent=styles['Normal'],
            fontSize=11,
            fontName='Helvetica',
            spaceBefore=3,
            spaceAfter=3,
            leftIndent=22,
            bulletIndent=10,
            alignment=TA_JUSTIFY
        )
        
        environment_style = ParagraphStyle(
            'EnvironmentStyle',
            parent=styles['Normal'],
            fontSize=9,
            fontName='Helvetica-Oblique',
            spaceBefore=8,
            spaceAfter=5,
            leftIndent=10,
            rightIndent=10,
            backColor=colors.HexColor('#F2F2F2'),
            borderWidth=0.5,
            borderColor=colors.HexColor('#E0E0E0'),
            borderPadding=4
        )
        
        content = []
        
        # Enhanced header section
        content.append(Paragraph(resume_data['name'], title_style))
        if resume_data.get('title'):
            content.append(Paragraph(resume_data['title'], subtitle_style))
        if resume_data.get('subtitle'):
            content.append(Paragraph(resume_data['subtitle'], contact_style))
        
        # Enhanced contact info
        contact_parts = []
        if resume_data.get('email'):
            contact_parts.append(resume_data['email'])
        if resume_data.get('phone'):
            contact_parts.append(resume_data['phone'])
        if resume_data.get('location'):
            contact_parts.append(resume_data['location'])
        if resume_data.get('linkedin'):
            contact_parts.append(resume_data['linkedin'])
        
        if contact_parts:
            content.append(Paragraph(' | '.join(contact_parts), contact_style))
        
        content.append(Spacer(1, 0.1*inch))
        
        # Enhanced sections with professional styling
        for section in resume_data['sections']:
            # Section title with underline
            section_title = section['name']
            if section_title == 'Projects':
                section_title = 'Professional Experience'
            
            content.append(Paragraph(section_title, section_title_style))
            content.append(HRFlowable(width="100%", thickness=1, 
                                    color=colors.HexColor('#B0C4DE'), 
                                    spaceBefore=3, spaceAfter=8))
            
            if section['type'] == 'professional_experience':
                for entry in section['content']:
                    # Use the transformed data structure with company, location, duration
                    company = entry.get('company', '')
                    location = entry.get('location', '')
                    duration = entry.get('duration', '')
                    
                    # Format company and location together
                    if company and location:
                        company_location_text = f"{company} | {location}"
                    elif company:
                        company_location_text = company
                    else:
                        company_location_text = ''
                    
                    if company_location_text:
                        content.append(Paragraph(company_location_text, company_style))
                    if duration:
                        content.append(Paragraph(duration, duration_style))
                    if entry.get('role'):
                        content.append(Paragraph(entry['role'], role_style))
                    
                    # Enhanced responsibilities with proper bullet formatting
                    for resp in entry.get('responsibilities', []):
                        content.append(Paragraph(f"• {resp}", bullet_style))
                    
                    # Environment section if present
                    if entry.get('environment'):
                        env_text = f"<b>Environment:</b> {entry['environment']}"
                        content.append(Paragraph(env_text, environment_style))
                    
                    content.append(Spacer(1, 0.15*inch))
            
            elif section['type'] == 'bullets':
                for bullet in section['content']:
                    content.append(Paragraph(f"• {bullet}", bullet_style))
                content.append(Spacer(1, 0.1*inch))
            
            elif section['type'] == 'plain_text_block':
                # Handle Technical Skills as table if it contains colons
                if section['name'] == 'Technical Skills' and ':' in section['content']:
                    lines = section['content'].split('\n')
                    skills_data = []
                    for line in lines:
                        if line.strip() and ':' in line:
                            parts = line.split(':', 1)
                            skills_data.append([f"{parts[0].strip()}:", parts[1].strip()])
                        elif line.strip():
                            skills_data.append([line.strip(), ''])
                    
                    if skills_data:
                        skills_table = Table(skills_data, colWidths=[2*inch, 4*inch])
                        skills_table.setStyle(TableStyle([
                            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                            ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
                            ('FONTSIZE', (0, 0), (-1, -1), 11),
                            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                            ('LEFTPADDING', (0, 0), (-1, -1), 0),
                            ('RIGHTPADDING', (0, 0), (0, -1), 15),
                            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                        ]))
                        content.append(skills_table)
                else:
                    # Regular plain text
                    lines = section['content'].split('\n')
                    for line in lines:
                        if line.strip():
                            content.append(Paragraph(line.strip(), bullet_style))
                content.append(Spacer(1, 0.1*inch))
            
            elif section['type'] == 'projects':
                for project in section['content']:
                    header = project.get('header', [])
                    if isinstance(header, list) and header:
                        # Company
                        content.append(Paragraph(header[0], company_style))
                        # Duration (usually the last element)
                        if len(header) > 1:
                            content.append(Paragraph(header[-1], duration_style))
                        # Role (middle element if exists)
                        if len(header) > 2:
                            content.append(Paragraph(header[1], role_style))
                    
                    # Enhanced responsibilities
                    for resp in project.get('responsibilities', []):
                        content.append(Paragraph(f"• {resp}", bullet_style))
                    
                    # Environment section
                    if project.get('environment'):
                        env_text = f"<b>Environment:</b> {project['environment']}"
                        content.append(Paragraph(env_text, environment_style))
                    
                    content.append(Spacer(1, 0.15*inch))
        
        # Build PDF content
        doc.build(content)
        logger.info(f"✅ PDF generated successfully using ReportLab (forced mode): {output_path}")
        return output_path
    else:
        # Try WeasyPrint first with enhanced error handling
        try:
            # Set additional environment variables for WeasyPrint if not set
            if not os.getenv('FONTCONFIG_PATH'):
                os.environ['FONTCONFIG_PATH'] = '/nix/store/*/etc/fonts:/usr/share/fonts'
            
            from weasyprint import HTML
            logger.info("WeasyPrint import successful, attempting PDF generation...")
            
            # Test WeasyPrint with a simple document first
            try:
                HTML(string="<html><body><h1>Test</h1></body></html>").write_pdf("/tmp/test.pdf")
                logger.info("WeasyPrint test document generation successful")
            except Exception as test_error:
                logger.warning(f"WeasyPrint test failed: {test_error}")
            
            # Generate actual PDF using cleaned HTML content
            HTML(string=html_content).write_pdf(output_path)
            logger.info(f"✅ PDF generated successfully using WeasyPrint: {output_path}")
            return output_path
        except ImportError as import_error:
            logger.warning(f"WeasyPrint import failed: {import_error}. Trying ReportLab as fallback...")
        except Exception as weasyprint_error:
            logger.warning(f"WeasyPrint failed: {weasyprint_error}. Trying ReportLab as fallback...")
        
        # Fallback to ReportLab (pure Python, no system dependencies)
        try:
            # Create PDF using ReportLab with enhanced formatting
            doc = SimpleDocTemplate(output_path, pagesize=letter,
                                      rightMargin=0.7*inch, leftMargin=0.7*inch,
                                      topMargin=0.7*inch, bottomMargin=0.7*inch)
            
            styles = getSampleStyleSheet()
            
            # Enhanced custom styles to match WeasyPrint
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Title'],
                fontSize=20,
                fontName='Helvetica-Bold',
                spaceAfter=2,
                alignment=TA_LEFT,
                textColor=colors.black
            )
            
            subtitle_style = ParagraphStyle(
                'CustomSubtitle',
                parent=styles['Normal'],
                fontSize=14,
                fontName='Helvetica',
                spaceAfter=2,
                alignment=TA_LEFT,
                textColor=colors.black
            )
            
            contact_style = ParagraphStyle(
                'ContactStyle',
                parent=styles['Normal'],
                fontSize=10,
                fontName='Helvetica',
                spaceAfter=15,
                alignment=TA_LEFT,
                textColor=colors.black
            )
            
            section_title_style = ParagraphStyle(
                'SectionTitle',
                parent=styles['Heading2'],
                fontSize=14,
                fontName='Helvetica-Bold',
                spaceBefore=15,
                spaceAfter=8,
                textColor=colors.HexColor('#4F81BD'),
                borderWidth=1,
                borderColor=colors.HexColor('#B0C4DE'),
                borderPadding=3
            )
            
            company_style = ParagraphStyle(
                'CompanyStyle',
                parent=styles['Normal'],
                fontSize=11,
                fontName='Helvetica-Bold',
                spaceBefore=8,
                spaceAfter=2,
                alignment=TA_LEFT
            )
            
            duration_style = ParagraphStyle(
                'DurationStyle',
                parent=styles['Normal'],
                fontSize=10,
                fontName='Helvetica-Bold',
                spaceAfter=2,
                alignment=TA_LEFT
            )
            
            role_style = ParagraphStyle(
                'RoleStyle',
                parent=styles['Normal'],
                fontSize=11,
                fontName='Helvetica-BoldOblique',
                spaceAfter=8,
                alignment=TA_LEFT
            )
            
            bullet_style = ParagraphStyle(
                'BulletStyle',
                parent=styles['Normal'],
                fontSize=11,
                fontName='Helvetica',
                spaceBefore=3,
                spaceAfter=3,
                leftIndent=22,
                bulletIndent=10,
                alignment=TA_JUSTIFY
            )
            
            environment_style = ParagraphStyle(
                'EnvironmentStyle',
                parent=styles['Normal'],
                fontSize=9,
                fontName='Helvetica-Oblique',
                spaceBefore=8,
                spaceAfter=5,
                leftIndent=10,
                rightIndent=10,
                backColor=colors.HexColor('#F2F2F2'),
                borderWidth=0.5,
                borderColor=colors.HexColor('#E0E0E0'),
                borderPadding=4
            )
            
            content = []
            
            # Enhanced header section
            content.append(Paragraph(resume_data['name'], title_style))
            if resume_data.get('title'):
                content.append(Paragraph(resume_data['title'], subtitle_style))
            if resume_data.get('subtitle'):
                content.append(Paragraph(resume_data['subtitle'], contact_style))
            
            # Enhanced contact info
            contact_parts = []
            if resume_data.get('email'):
                contact_parts.append(resume_data['email'])
            if resume_data.get('phone'):
                contact_parts.append(resume_data['phone'])
            if resume_data.get('location'):
                contact_parts.append(resume_data['location'])
            if resume_data.get('linkedin'):
                contact_parts.append(resume_data['linkedin'])
            
            if contact_parts:
                content.append(Paragraph(' | '.join(contact_parts), contact_style))
            
            content.append(Spacer(1, 0.1*inch))
            
            # Enhanced sections with professional styling
            for section in resume_data['sections']:
                # Section title with underline
                section_title = section['name']
                if section_title == 'Projects':
                    section_title = 'Professional Experience'
                
                content.append(Paragraph(section_title, section_title_style))
                content.append(HRFlowable(width="100%", thickness=1, 
                                        color=colors.HexColor('#B0C4DE'), 
                                        spaceBefore=3, spaceAfter=8))
                
                
                if section['type'] == 'professional_experience':
                    for entry in section['content']:
                        # Use the transformed data structure with company, location, duration
                        company = entry.get('company', '')
                        location = entry.get('location', '')
                        duration = entry.get('duration', '')
                        
                        # Format company and location together
                        if company and location:
                            company_location_text = f"{company} | {location}"
                        elif company:
                            company_location_text = company
                        else:
                            company_location_text = ''
                        
                        if company_location_text:
                            content.append(Paragraph(company_location_text, company_style))
                        if duration:
                            content.append(Paragraph(duration, duration_style))
                        if entry.get('role'):
                            content.append(Paragraph(entry['role'], role_style))
                        
                        # Enhanced responsibilities with proper bullet formatting
                        for resp in entry.get('responsibilities', []):
                            content.append(Paragraph(f"• {resp}", bullet_style))
                        
                        # Environment section if present
                        if entry.get('environment'):
                            env_text = f"<b>Environment:</b> {entry['environment']}"
                            content.append(Paragraph(env_text, environment_style))
                        
                        content.append(Spacer(1, 0.15*inch))

                
                elif section['type'] == 'bullets':
                    for bullet in section['content']:
                        content.append(Paragraph(f"• {bullet}", bullet_style))
                    content.append(Spacer(1, 0.1*inch))
                
                elif section['type'] == 'plain_text_block':
                    # Handle Technical Skills as table if it contains colons
                    if section['name'] == 'Technical Skills' and ':' in section['content']:
                        lines = section['content'].split('\n')
                        skills_data = []
                        for line in lines:
                            if line.strip() and ':' in line:
                                parts = line.split(':', 1)
                                skills_data.append([f"{parts[0].strip()}:", parts[1].strip()])
                            elif line.strip():
                                skills_data.append([line.strip(), ''])
                        
                        if skills_data:
                            skills_table = Table(skills_data, colWidths=[2*inch, 4*inch])
                            skills_table.setStyle(TableStyle([
                                ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                                ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
                                ('FONTSIZE', (0, 0), (-1, -1), 11),
                                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                                ('RIGHTPADDING', (0, 0), (0, -1), 15),
                                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                            ]))
                            content.append(skills_table)
                    else:
                        # Regular plain text
                        lines = section['content'].split('\n')
                        for line in lines:
                            if line.strip():
                                content.append(Paragraph(line.strip(), bullet_style))
                    content.append(Spacer(1, 0.1*inch))
                
                elif section['type'] == 'projects':
                    for project in section['content']:
                        header = project.get('header', [])
                        if isinstance(header, list) and header:
                            # Company
                            content.append(Paragraph(header[0], company_style))
                            # Duration (usually the last element)
                            if len(header) > 1:
                                content.append(Paragraph(header[-1], duration_style))
                            # Role (middle element if exists)
                            if len(header) > 2:
                                content.append(Paragraph(header[1], role_style))
                        
                        # Enhanced responsibilities
                        for resp in project.get('responsibilities', []):
                            content.append(Paragraph(f"• {resp}", bullet_style))
                        
                        # Environment section
                        if project.get('environment'):
                            env_text = f"<b>Environment:</b> {project['environment']}"
                            content.append(Paragraph(env_text, environment_style))
                        
                        content.append(Spacer(1, 0.15*inch))
            
            # Build PDF
            doc.build(content)
            
            logger.info(f"PDF generated successfully using ReportLab: {output_path}")
            return output_path
            
        except Exception as reportlab_error:
            logger.error(f"Both WeasyPrint and ReportLab failed. WeasyPrint: {weasyprint_error}, ReportLab: {reportlab_error}")
            raise Exception("Failed to generate PDF resume. Both WeasyPrint and ReportLab failed.")

def save_optimized_resume(filename: str, resume_text: str, suggestions: str) -> str:
    base_name = os.path.splitext(filename)[0]
    output_path = os.path.join(GENERATED_DIR, f"{base_name}_optimized.pdf")
    logger.info(f"Saving optimized resume to {output_path}")

    # Use WeasyPrint to generate a PDF with the resume text and suggestions
    try:
        from weasyprint import HTML
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <style>
                body {{ font-family: Calibri, Arial, sans-serif; margin: 0.7in; line-height: 1.2; font-size: 11pt; }}
                h1 {{ font-size: 14pt; }}
                h2 {{ font-size: 12pt; }}
            </style>
        </head>
        <body>
            <h1>Optimized Resume Content</h1>
            <pre>{{ resume_text }}</pre>
            <h2>Suggestions</h2>
            <p>{{ suggestions }}</p>
        </body>
        </html>
        """
        if 'darwin' in str(sys.platform):
            os.environ['DYLD_LIBRARY_PATH'] = '/opt/homebrew/lib:' + os.environ.get('DYLD_LIBRARY_PATH', '')
        HTML(string=html_content).write_pdf(output_path)
        logger.info("Optimized resume PDF saved successfully.")
        return output_path
    except Exception as e:
        logger.error(f"Failed to save optimized resume PDF: {e}", exc_info=True)
        raise Exception("PDF generation failed. Please ensure WeasyPrint is properly installed and configured.")
