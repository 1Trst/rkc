#!/usr/bin/env python3
"""
Streamlit Article Cleaner and Newsletter Generator: upload, clean, generate, and download Word documents with minimal API calls
"""
import streamlit as st
import zipfile
import io
import tempfile
import os
import re
import base64
from docx import Document as DocxDocument
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import AzureOpenAI
import asyncio
import httpx
import json
from concurrent.futures import ThreadPoolExecutor

# Azure settings - Get from environment variables
key = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT", "https://rkcazureai.cognitiveservices.azure.com/")

#Edit
# Language selection
LANGUAGES = {
    "Français": "French",
    "English": "English", 
    "Deutsch": "German"
}
def check_password():
    """Returns True if the user has entered the correct password."""
    
    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == "rkcnewsletterWS2025":
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Don't store the password
        else:
            st.session_state["password_correct"] = False

    # Return True if password is validated
    if st.session_state.get("password_correct", False):
        return True

    # Show input for password
    st.markdown("### 🔐 Authentication Required")
    st.text_input(
        "Please enter the password to access the application:",
        type="password",
        on_change=password_entered,
        key="password",
        placeholder="Enter password..."
    )
    
    if "password_correct" in st.session_state:
        st.error("😞 Password incorrect. Please try again.")
    
    st.markdown("---")
    st.info("💡 Contact your administrator if you need access.")
    return False

# Check password before showing the main app
if not check_password():
    st.stop()
# UI
st.title("📝 Article Cleaner & Newsletter Generator v1")

# Show API key status in sidebar
with st.sidebar:
    if key:
        st.success("✅ Azure OpenAI API Key loaded")
        st.info(f"🔗 Endpoint: {AZURE_OPENAI_ENDPOINT}")
    else:
        st.error("❌ Azure OpenAI API Key not found")
        st.warning("Please set AZURE_OPENAI_API_KEY environment variable")

# Stop execution if no API key
if not key:
    st.error("🔑 **Configuration Error**: Azure OpenAI API Key not found in environment variables.")
    st.markdown("""
    **For local development:**
    1. Create a `.env` file in your project directory
    2. Add: `AZURE_OPENAI_API_KEY=your_api_key_here`
    
    **For deployment:**
    - Set the `AZURE_OPENAI_API_KEY` environment variable in your deployment platform
    """)
    st.stop()

# Main tabs - séparés pour utilisation indépendante
tab1, tab2, tab3 = st.tabs(["🧹 Article Cleaning", "📰 Newsletter Generator", "⚙️ Settings"])

# Initialize Azure OpenAI client
def get_client(key):
    """Get Azure OpenAI client properly"""
    if not key:
        return None
    try:
        from openai import AzureOpenAI
        
        client = AzureOpenAI(
            api_key=key,
            api_version="2024-02-15-preview",
            azure_endpoint=AZURE_OPENAI_ENDPOINT,
        )
        return client
    except Exception as e:
        st.error(f"Error initializing Azure OpenAI client: {str(e)}")
        return None

# Ensure Templates directory exists
def ensure_templates_dir():
    """Ensure Templates directory exists with at least a Generic template"""
    os.makedirs("Templates", exist_ok=True)
    # Create a generic template if none exists
    generic_path = os.path.join("Templates", "Generic_Template.docx")
    if not os.path.exists(generic_path):
        doc = DocxDocument()
        doc.save(generic_path)
    return os.path.exists("Templates")

# File handling functions
def read_file_content(file_path):
    """Read file content based on extension"""
    ext = file_path.split('.')[-1].lower()
    
    try:
        if ext == 'pdf':
            import PyPDF2
            with open(file_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                return "\n".join([page.extract_text() for page in reader.pages])
        elif ext == 'txt':
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                return f.read()
        elif ext == 'docx':
            doc = DocxDocument(file_path)
            return "\n".join([paragraph.text for paragraph in doc.paragraphs])
        elif ext == 'md':
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                return f.read()
        else:
            return ""
    except Exception as e:
        st.error(f"Error reading {file_path}: {str(e)}")
        return ""

# Get language-specific prompts
def get_language_instructions(language, article_summary_sentences=3, exec_summary_sentences=7):
    """Get language-specific instructions for prompts"""
    return {
        "summary_lang": language.upper(),
        "summary_instruction": f"Generate around {article_summary_sentences} sentences summary in {language.upper()}",
        "exec_summary_intro": f"You are a consultant in a French consulting firm. Your task is to write an Executive summary of around {exec_summary_sentences} sentences based on the following articles:",
        "exec_summary_requirements": f"""This summary should:
1. Synthesize the key ideas and points from the articles
2. Identify trends and points of divergence
3. Highlight relevant implications for our consulting firm (no vague or general implications)
4. If relevant, formulate concrete recommendations

Respond only with the executive summary in {language.upper()} (around {exec_summary_sentences} sentences), without additional introduction or conclusion."""
    }

# Process articles with Azure OpenAI
async def process_article_async(client, text, language="French", article_summary_sentences=3):
    """Process a single article with Azure OpenAI"""
    lang_instructions = get_language_instructions(language, article_summary_sentences)
    
    # Get language-specific instructions for translation
    if language == "French":
        translation_instruction = "Keep the title and article text in their original language if they are already in French, otherwise translate them to French."
        output_language = "French"
    elif language == "English":
        translation_instruction = "Translate the title and article text to English."
        output_language = "English"
    else:  # German
        translation_instruction = "Translate the title and article text to German."
        output_language = "German"
    
    # Extract metadata and clean in one call
    combined_prompt = f"""You will be provided a poorly copy/pasted article from a single journal/website. Please extract the following information and clean this article in JSON format:

1. Extract and clean the title (remove problematic characters like &,/,<,>,#,»,«,é,è,ê,â,à), then translate it to {output_language}.
2. Extract the source of the article (journal or website name) it has to be in the following list : ["Consultor","Financial Times","Handelsblatt","La Lettre_du_Conseil","La Lettre","Les Echos Investir","Les Echos","Le Monde"]. Pay close attention to this part, and detect the difference between "Les Echos" et "Les Echos Investir"
3. Extract the date (format: d MMMM yyyy). If you can't find the article date, or aren't 100% certain, use today's date
4. Clean the article by removing any website boilerplate, ads, or irrelevant content. Please try to respect and understand the different paragraphs of the original article and reproduce those in your cleaned text. Then translate the entire cleaned article text to {output_language}.
5. {lang_instructions["summary_instruction"]}
6. Double check the source : is what you found correct ?

IMPORTANT: {translation_instruction}

Return ONLY a JSON object with these fields:
{{
    "title": "cleaned title translated to {output_language}",
    "source": "source (keep original name)", 
    "date": "date",
    "year": "yyyy",
    "cleaned_text": "full cleaned article text translated to {output_language}",
    "summary": "around {article_summary_sentences} sentences summary in {lang_instructions['summary_lang']}"
}}

Here's the text to process:

{text}"""

    try:
        response = await asyncio.to_thread(
            client.chat.completions.create,
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": combined_prompt}],
            response_format={"type": "json_object"},
            temperature=0
        )
        
        return response.choices[0].message.content
    except Exception as e:
        st.error(f"Error processing article: {str(e)}")
        return "{}"

# Generate executive summary
async def generate_executive_summary(client, articles, language="French", exec_summary_sentences=7):
    """Generate an executive summary from article summaries"""
    lang_instructions = get_language_instructions(language, 3, exec_summary_sentences)
    
    # Create a prompt with all article summaries
    article_summaries = []
    for article in articles:
        article_summaries.append(f"**{article['title']}**: {article['summary']}")
    
    prompt = f"""{lang_instructions["exec_summary_intro"]}

{chr(10).join(article_summaries)}

{lang_instructions["exec_summary_requirements"]}"""

    try:
        response = await asyncio.to_thread(
            client.chat.completions.create,
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )
        
        return response.choices[0].message.content
    except Exception as e:
        st.error(f"Error generating executive summary: {str(e)}")
        return "Error generating executive summary."

# Process multiple articles in parallel
async def process_articles(client, articles, language="French", article_summary_sentences=3):
    """Process multiple articles concurrently"""
    tasks = [process_article_async(client, text, language, article_summary_sentences) for text in articles]
    return await asyncio.gather(*tasks)

# Create Word document from cleaned article
def create_word_doc(article_data, output_path):
    """Create a Word document from cleaned article data using source-based template"""
    source = article_data.get("source", "Generic")
    source_template_name = source.replace(' ', '_') + '_Template.docx'
    
    # Look for source-specific template
    template_path = os.path.join('Templates', source_template_name)
    if not os.path.exists(template_path):
        # Fall back to Generic template
        template_path = os.path.join('Templates', 'Generic_Template.docx')
        if not os.path.exists(template_path):
            # If generic template doesn't exist, create a new document
            doc = DocxDocument()
        else:
            doc = DocxDocument(template_path)
    else:
        # Found source-specific template
        doc = DocxDocument(template_path)
    
    # Set document properties
    core_props = doc.core_properties
    core_props.title = "{}_{}_{}".format(
    article_data.get("source", "Unknown"),
    article_data.get("title", "Untitled"),
    article_data.get("date", "No date"))
    core_props.author = article_data.get("source", "Unknown")
    
    # Add title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run(f"{article_data.get('title', 'Untitled')} — {article_data.get('source', 'Unknown')}, {article_data.get('date', 'No date')}")
    title_run.bold = True
    title_run.font.size = Pt(16)
    title_run.font.name = 'Aptos'
    
    # Add article body
    body = doc.add_paragraph(article_data.get("cleaned_text", "No content available"))
    body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    # Save document
    doc.save(output_path)
    return output_path

# Create newsletter document
def create_newsletter_doc(exec_summary, articles, output_path):
    """Create a Word document with newsletter content"""    
    doc = DocxDocument()
    # Set document properties
    core_props = doc.core_properties
    core_props.title = "AI Generated Newsletter"
    # Title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    title_run = title.add_run("AI Generated Newsletter")
    title_run.bold = True
    title_run.font.size = Pt(18)
    title_run.font.name = 'Aptos'
    
    # Date
    date = doc.add_paragraph()
    date.alignment = WD_ALIGN_PARAGRAPH.LEFT
    from datetime import datetime
    date_run = date.add_run(datetime.now().strftime("%d %B %Y"))
    date_run.italic = True
    date_run.font.size = Pt(12)
    
    # Executive summary 
    # heading
    doc.add_paragraph()
    exec_heading = doc.add_paragraph()
    exec_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
    exec_heading_run = exec_heading.add_run("Executive Summary")
    exec_heading_run.bold = True
    exec_heading_run.font.size = Pt(16)    
    # content
    exec_content = doc.add_paragraph(exec_summary)
    exec_content.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    # Separator
    doc.add_paragraph()
    separator = doc.add_paragraph("* * *")
    separator.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    
    # Add articles
    for article in articles:
        # Article title
        article_title = doc.add_paragraph()
        article_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
        title_text = f"{article['title']} — {article['source']}, {article['date']}"
        title_run = article_title.add_run(title_text)
        title_run.bold = True
        title_run.font.size = Pt(14)
        
        # Article summary
        summary = doc.add_paragraph(article['summary'])
        summary.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        # Separator between articles
        doc.add_paragraph()
        separator = doc.add_paragraph("* * *")
        separator.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()
    
    # Save document
    doc.save(output_path)
    return output_path

# Create download link for file
def get_download_link(file_path, link_text):
    """Generate a download link for a file"""
    with open(file_path, "rb") as f:
        data = f.read()
    b64 = base64.b64encode(data).decode()
    file_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    href = f'<a href="data:{file_type};base64,{b64}" download="{os.path.basename(file_path)}">{link_text}</a>'
    return href

# Create zip file from multiple files
def create_zip_from_files(file_paths):
    """Create a zip file from multiple files"""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for file_path in file_paths:
            zf.write(file_path, os.path.basename(file_path))
    return zip_buffer.getvalue()

# Extract file paths
def get_paths(files):
    """Get file paths from uploaded files"""
    workdir = tempfile.mkdtemp()
    paths = []
    for f in files or []:
        dest = os.path.join(workdir, f.name)
        with open(dest, 'wb') as out:
            out.write(f.read())
        paths.append(dest)
    return paths, workdir

# Settings tab
with tab3:
    st.header("⚙️ Settings")
    
    # Language selection
    selected_language = st.selectbox(
        "Choose output language / Choisir la langue de sortie:",
        options=list(LANGUAGES.keys()),
        index=0,  # Default to French
        help="This will affect the language of article summaries and newsletter content",
        key="language_selector"  # Add unique key
)
    
    # Store in session state
    st.session_state.output_language = LANGUAGES[selected_language]
    
    st.info(f"Selected language: {selected_language} / Langue sélectionnée: {selected_language}")
    
    # Summary length settings
    st.subheader("📏 Summary Length Settings")
    
    col1, col2 = st.columns(2)
    
    with col1:
        article_summary_sentences = st.number_input(
            "Article summary length (sentences):",
            min_value=1,
            max_value=10,
            value=3,
            step=1,
            help="Number of sentences for each article summary in the newsletter",
             key="article_summary_length"
        )
        st.session_state.article_summary_sentences = article_summary_sentences
    
    with col2:
        exec_summary_sentences = st.number_input(
            "Executive summary length (sentences):",
            min_value=3,
            max_value=15,
            value=7,
            step=1,
            help="Number of sentences for the executive summary",
            key="exec_summary_length" 
        )
        st.session_state.exec_summary_sentences = exec_summary_sentences
    
    # Preview of settings
    st.subheader("📋 Current Settings Preview")
    st.info(f"""
    **Language:** {selected_language}
    **Article summaries:** Around {article_summary_sentences} sentences each
    **Executive summary:** Around {exec_summary_sentences} sentences
    """)
    
    st.markdown("---")
    st.markdown("💡 **Tip:** These settings will be applied to all new articles and newsletters you generate.")

# Article Cleaning Tab
with tab1:
    st.header("🧹 Article Cleaning")
    
    # Check if Templates directory exists
    ensure_templates_dir()
    
    # Get language setting
    output_language = st.session_state.get('output_language', 'French')
    article_summary_sentences = st.session_state.get('article_summary_sentences', 3)
    
    st.subheader("Input Methods")
    
    # Input method selection
    input_method = st.radio(
        "Choose input method:",
        ["Upload Files", "Copy-Paste Text"],
        horizontal=True
    )
    
    articles_to_process = []
    article_names = []
    
    if input_method == "Upload Files":
        # File upload
        uploaded_files = st.file_uploader(
            "Upload one or multiple files", 
            type=["pdf", "txt", "docx", "md"], 
            accept_multiple_files=True
        )
        
        if uploaded_files:
            paths, workdir = get_paths(uploaded_files)
            for path in paths:
                content = read_file_content(path)
                if content:
                    articles_to_process.append(content)
                    article_names.append(os.path.basename(path))
    
    else:  # Copy-Paste Text
        # Text input
        pasted_text = st.text_area(
            "Paste your article text here:",
            height=300,
            placeholder="Copy and paste your article content here..."
        )
        
        if pasted_text.strip():
            articles_to_process.append(pasted_text.strip())
            article_names.append("Pasted_Article")
    
    # Show current settings
    if articles_to_process:
        st.info(f"Ready to process {len(articles_to_process)} articles with {article_summary_sentences} sentences summaries in {output_language}")
    
    # Clean articles button
    if st.button("🧹 Clean Articles", use_container_width=True):
        client = get_client(key)
        
        if not client:
            st.error("Failed to initialize Azure OpenAI client")
        elif not articles_to_process:
            st.error("No articles to process")
        else:
            with st.spinner("Cleaning articles..."):
                # Process articles
                results = asyncio.run(process_articles(client, articles_to_process, output_language, article_summary_sentences))
                
                # Parse results
                cleaned_articles = []
                for i, result in enumerate(results):
                    try:
                        article_data = json.loads(result)
                        cleaned_articles.append(article_data)
                    except json.JSONDecodeError:
                        st.error(f"Error parsing response for {article_names[i]}")
                
                if cleaned_articles:
                    # Create temporary directory for this session
                    workdir = tempfile.mkdtemp()
                    
                    # Create Word documents
                    doc_paths = []
                    for i, article in enumerate(cleaned_articles):
                        art_source = re.sub(r'[^0-9A-Za-z_]+', '_', article.get("source", "Unknown"))
                        art_title = re.sub(r'[^0-9A-Za-z_]+', '_', article.get("title", f"doc_{i}"))
                        art_date = re.sub(r'[^0-9A-Za-z_]+', '_', article.get("date", "NoDate"))
                        file_name = f"{art_source}_{art_title}_{art_date}.docx"
                        doc_path = os.path.join(workdir, file_name)
                        create_word_doc(article, doc_path)
                        doc_paths.append(doc_path)
                    
                    # Save results to session state
                    st.session_state.cleaned_articles_tab1 = cleaned_articles
                    st.session_state.doc_paths_tab1 = doc_paths
                    st.session_state.workdir_tab1 = workdir
                    
                    st.success(f"Successfully cleaned {len(cleaned_articles)} articles!")

    # Display results if available
    if 'cleaned_articles_tab1' in st.session_state:
        st.subheader("Cleaned Articles")
        
        # Display articles
        for i, article in enumerate(st.session_state.cleaned_articles_tab1):
            with st.expander(f"{article.get('title', f'Article {i+1}')} - {article.get('source', 'Unknown')}"):
                col1, col2 = st.columns([1, 1])
                with col1:
                    st.markdown(f"**Source:** {article.get('source', 'Unknown')}")
                    st.markdown(f"**Date:** {article.get('date', 'No date')}")
                with col2:
                    st.markdown(f"**Output Language:** {output_language}")
                    st.markdown(f"**Summary Length:** {article_summary_sentences} sentences")
                
                st.markdown("**Summary:**")
                st.write(article.get('summary', 'No summary'))
                
                st.markdown("**Cleaned & Translated Text:**")
                st.text_area("", article.get("cleaned_text", "No content"), height=200, key=f"clean_text_{i}")
        
        # Download options
        st.subheader("Download Options")
        
        col1, col2 = st.columns([1, 1])
        
        with col1:
            st.markdown("**Individual Downloads:**")
            for i, path in enumerate(st.session_state.doc_paths_tab1):
                article = st.session_state.cleaned_articles_tab1[i]
                st.markdown(
                    get_download_link(path, f"📄 {article.get('title', f'Article {i+1}')}"), 
                    unsafe_allow_html=True
                )
        
        with col2:
            if len(st.session_state.doc_paths_tab1) > 1:
                st.markdown("**Batch Download:**")
                zip_data = create_zip_from_files(st.session_state.doc_paths_tab1)
                b64_zip = base64.b64encode(zip_data).decode()
                href = f'<a href="data:application/zip;base64,{b64_zip}" download="cleaned_articles.zip">📦 Download All Articles (ZIP)</a>'
                st.markdown(href, unsafe_allow_html=True)

# Newsletter Generator Tab
with tab2:
    st.header("📰 Newsletter Generator")
    
    # Get language setting
    output_language = st.session_state.get('output_language', 'French')
    article_summary_sentences = st.session_state.get('article_summary_sentences', 3)
    exec_summary_sentences = st.session_state.get('exec_summary_sentences', 7)
    
    st.info(f"Newsletter will be generated in: {output_language} | Article summaries: ~{article_summary_sentences} sentences | Executive summary: ~{exec_summary_sentences} sentences")
    
    st.subheader("Article Sources")
    
    # Option to use cleaned articles from Tab 1 or upload new ones
    newsletter_source = st.radio(
        "Choose article source for newsletter:",
        ["Use cleaned articles from Article Cleaning tab", "Upload/Import new articles"],
        help="You can either use articles you've already cleaned, or process new articles specifically for the newsletter"
    )
    
    articles_for_newsletter = []
    newsletter_articles_to_process = []
    newsletter_article_names = []
        
    if newsletter_source == "Use cleaned articles from Article Cleaning tab":
        if 'cleaned_articles_tab1' in st.session_state:
            articles_for_newsletter = st.session_state.cleaned_articles_tab1
            st.success(f"Found {len(articles_for_newsletter)} cleaned articles ready for newsletter generation")
            
            # Show preview of articles
            with st.expander("Preview articles for newsletter"):
                for i, article in enumerate(articles_for_newsletter):
                    st.markdown(f"**{i+1}. {article.get('title', 'Untitled')}** - {article.get('source', 'Unknown')}")
        else:
            st.warning("No cleaned articles found. Please clean some articles in the 'Article Cleaning' tab first, or choose to upload new articles.")
    
    else:  # Upload/Import new articles
        st.subheader("Upload Articles for Newsletter")
        
        # Input method for newsletter
        newsletter_input_method = st.radio(
            "Choose input method:",
            ["Upload Files", "Copy-Paste Text"],
            horizontal=True,
            key="newsletter_input"
        )
        
        
        if newsletter_input_method == "Upload Files":
            newsletter_uploaded_files = st.file_uploader(
                "Upload files for newsletter", 
                type=["pdf", "txt", "docx", "md"], 
                accept_multiple_files=True,
                key="newsletter_upload"
            )
            
            if newsletter_uploaded_files:
                paths, workdir = get_paths(newsletter_uploaded_files)
                for path in paths:
                    content = read_file_content(path)
                    if content:
                        newsletter_articles_to_process.append(content)
                        newsletter_article_names.append(os.path.basename(path))
        
        else:  # Copy-Paste for newsletter
            newsletter_pasted_text = st.text_area(
                "Paste articles for newsletter (separate multiple articles with '---'):",
                height=300,
                placeholder="Copy and paste your article content here...\n\n---\n\nUse --- to separate multiple articles",
                key="newsletter_paste"
            )
            
            if newsletter_pasted_text.strip():
                # Split by --- if multiple articles
                split_articles = [art.strip() for art in newsletter_pasted_text.split('---') if art.strip()]
                newsletter_articles_to_process.extend(split_articles)
                newsletter_article_names.extend([f"Pasted_Article_{i+1}" for i in range(len(split_articles))])
        
        # Show what will be processed
        if newsletter_articles_to_process:
            st.info(f"Ready to process {len(newsletter_articles_to_process)} articles for newsletter")
            with st.expander("Preview articles to process"):
                for i, name in enumerate(newsletter_article_names):
                    st.markdown(f"**{i+1}. {name}**")
    
    # Single button to generate newsletter
    generate_newsletter_button = False
    
    if newsletter_source == "Use cleaned articles from Article Cleaning tab":
        if articles_for_newsletter:
            generate_newsletter_button = st.button("📰 Generate Newsletter", use_container_width=True, key="gen_newsletter_existing")
    else:
        if newsletter_articles_to_process:
            generate_newsletter_button = st.button("🔄 Process Articles & Generate Newsletter", use_container_width=True, key="gen_newsletter_new")
    
    # Generate Newsletter (unified process)
    if generate_newsletter_button:
        client = get_client(key)
        if not client:
            st.error("Failed to initialize Azure OpenAI client")
        else:
            # If we need to process new articles first
            if newsletter_source != "Use cleaned articles from Article Cleaning tab":
                with st.spinner("Processing articles for newsletter..."):
                    # Process articles
                    results = asyncio.run(process_articles(client, newsletter_articles_to_process, output_language, article_summary_sentences))
                    
                    # Parse results
                    cleaned_articles = []
                    for i, result in enumerate(results):
                        try:
                            article_data = json.loads(result)
                            cleaned_articles.append(article_data)
                        except json.JSONDecodeError:
                            st.error(f"Error parsing response for {newsletter_article_names[i]}")
                    
                    if cleaned_articles:
                        articles_for_newsletter = cleaned_articles
                        st.success(f"Successfully processed {len(cleaned_articles)} articles!")
                    else:
                        st.error("Failed to process articles")
                        articles_for_newsletter = []
            
            # Generate newsletter if we have articles
            if articles_for_newsletter:
                with st.spinner("Generating newsletter..."):
                    # Extract summaries for newsletter
                    article_summaries = []
                    for article in articles_for_newsletter:
                        article_summaries.append({
                            'title': article.get('title', 'Untitled'),
                            'source': article.get('source', 'Unknown'),
                            'date': article.get('date', 'No date'),
                            'summary': article.get('summary', 'No summary available')
                        })
                    
                    # Generate executive summary
                    exec_summary = asyncio.run(generate_executive_summary(client, articles_for_newsletter, output_language, exec_summary_sentences))
                    
                    # Save to session state
                    st.session_state.article_summaries_newsletter = article_summaries
                    st.session_state.exec_summary_newsletter = exec_summary
                    st.session_state.newsletter_language = output_language
                    st.session_state.newsletter_article_sentences = article_summary_sentences
                    st.session_state.newsletter_exec_sentences = exec_summary_sentences
                    
                    st.success("Newsletter generated successfully!")
    
    # Display Newsletter Preview
    if 'exec_summary_newsletter' in st.session_state:
        st.subheader("Newsletter Preview")
        
        # Language info
        newsletter_lang = st.session_state.get('newsletter_language', 'French')
        st.info(f"Newsletter language: {newsletter_lang}")
        
        # Executive Summary
        st.markdown("### Executive Summary")
        st.write(st.session_state.exec_summary_newsletter)
        st.markdown("---")
        
        # Article summaries
        st.markdown("### Articles")
        for article in st.session_state.article_summaries_newsletter:
            st.markdown(f"**{article['title']}** — *{article['source']}, {article['date']}*")
            st.write(article['summary'])
            st.markdown("---")
        
        # Download Newsletter
        if st.button("📄 Download Newsletter as Word Document", use_container_width=True):
            with st.spinner("Creating newsletter document..."):
                # Create temporary directory if it doesn't exist
                workdir = tempfile.mkdtemp()
                
                # Create newsletter document
                newsletter_filename = "newsletter.docx"
                doc_path = create_newsletter_doc(
                    st.session_state.exec_summary_newsletter, 
                    st.session_state.article_summaries_newsletter,
                    os.path.join(workdir, newsletter_filename)
                )
                
                # Generate download link
                st.markdown(
                    get_download_link(doc_path, "📄 Download Newsletter Document"),
                    unsafe_allow_html=True
                )
                
                st.success("Newsletter document ready for download!")
    
    elif not articles_for_newsletter and newsletter_source == "Use cleaned articles from Article Cleaning tab":
        st.info("Please clean some articles in the 'Article Cleaning' tab first, or choose to upload new articles.")
    elif not newsletter_articles_to_process and newsletter_source != "Use cleaned articles from Article Cleaning tab":
        st.info("Please upload files or paste text to generate a newsletter.")
