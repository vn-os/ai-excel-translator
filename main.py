import os
import time
import re
import glob
import xlwings as xw
from gooey import Gooey, GooeyParser
from openai import OpenAI
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Set proxy environment variables if they exist
if http_proxy := os.environ.get("HTTP_PROXY"):
    os.environ["http_proxy"] = http_proxy
if https_proxy := os.environ.get("HTTPS_PROXY"):
    os.environ["https_proxy"] = https_proxy
os.environ["no_proxy"] = "127.0.0.1,localhost,.local"

llm_api = {
    "url": os.getenv("LLM_API_URL"),
    "api_key": os.getenv("LLM_API_KEY"),
    "model": os.getenv("LLM_MODEL_NAME"),
}

llm_model_no_think = False
if temp := os.getenv("LLM_MODEL_NO_THINK"):
    if temp.lower() in ["1", "true"]:
        llm_model_no_think = True

lang_map = {
    'en': 'English',      # Most widely spoken language
    'zh': 'Chinese',      # Mandarin Chinese
    'hi': 'Hindi',        # Hindi
    'es': 'Spanish',      # Spanish
    'ar': 'Arabic',       # Arabic
    'bn': 'Bengali',      # Bengali
    'pt': 'Portuguese',   # Portuguese
    'ru': 'Russian',      # Russian
    'ja': 'Japanese',     # Japanese
    'vi': 'Vietnamese',   # Vietnamese
    'ko': 'Korean',       # Korean
    'id': 'Indonesian',   # Indonesian
    'fr': 'French',       # French
    'de': 'German',       # German
    'tr': 'Turkish',      # Turkish
    'it': 'Italian',      # Italian
    'th': 'Thai',         # Thai
    'pl': 'Polish',       # Polish
    'uk': 'Ukrainian',    # Ukrainian
    'nl': 'Dutch',        # Dutch
}

separator = "¦¦¦"

system_prompt = f"""
You are a professional IT translator specializing in software development, programming, and technical documentation. Follow these rules strictly:

1. Output ONLY the translation, nothing else
2. DO NOT include the original text in your response
3. DO NOT add any explanations or notes
4. Keep IDs, model numbers, and special characters unchanged
5. Use standard terminology for technical terms in IT and software development
6. Preserve the original formatting (spaces, line breaks)
7. Use proper grammar and punctuation
8. Only keep unchanged: proper names, IDs, and technical codes
9. Translate all segments separated by "{separator}" and keep them separated with the same delimiter

For IT-specific terminology:
- Maintain consistency in technical terms
- Keep programming language keywords, function names, and variable names unchanged
- Use industry-standard translations for common IT concepts
- Preserve acronyms like API, UI, UX, SQL, HTML, CSS, etc.
- Keep file extensions and paths unchanged
"""

# LLM API Client
llm_client = None

# Set API delay and batch size
API_DELAY = 2  # Delay 2 seconds between API calls
BATCH_SIZE = 100  # Maximum number of cells in a batch

def clean_text(text):
    """Clean and normalize text before translation"""
    if not text or not isinstance(text, str):
        return ""
    text = ' '.join(text.split())  # Normalize whitespace
    return text.strip()

def should_translate(text):
    """Check if a cell needs translation"""
    text = clean_text(text)
    if not text or len(text) < 2:
        return False
    if re.match(r'^[\d\s,.-]+$', text):  # Contains only numbers and number formatting characters
        return False
    if text.startswith('='):  # Excel formula
        return False
    return True

def translate_batch(texts, source_lang, target_lang):
    """Translate a batch of texts to the target language"""
    if not texts:
        return []

    # Combine texts with separator
    combined_text = separator.join(texts)

    # Get translation direction
    source = lang_map.get(source_lang)
    target = lang_map.get(target_lang)
    if not source or not target:
        print(f"❌ Invalid language combination: from {source_lang} to {target_lang}")
        return texts

    user_prompt = f"Translate the following text from {source} to {target}, keeping segments separated by '{separator}':\n\n{combined_text}"

    if llm_model_suffix := os.getenv("LLM_MODEL_SUFFIX"):
        user_prompt += "\n"
        user_prompt += llm_model_suffix

    try:
        global llm_client
        if not llm_client:
            llm_client = OpenAI(
                base_url=llm_api["url"],
                api_key=llm_api["api_key"],
            )
            print(f"🤖 Using LLM model: '{llm_api['model']}' at '{llm_api['url']}'")

        # Call translation API
        response = llm_client.chat.completions.create(
            model=llm_api["model"], # Or "gemini-pro" or other suitable model
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ]
        )

        # Split translation result into separate parts
        translated_text = response.choices[0].message.content

        if llm_model_no_think:
            translated_text = re.sub(r'<think>.*?</think>\n*', '', translated_text, flags=re.DOTALL)

        translated_parts = translated_text.split(separator)

        # Handle case when number of translated parts doesn't match
        if len(translated_parts) != len(texts):
            print(f"   ⚠️ Number of translated parts ({len(translated_parts)}) does not match number of original texts ({len(texts)})")
            # Ensure number of translated parts equals number of original texts
            if len(translated_parts) < len(texts):
                translated_parts.extend(texts[len(translated_parts):])
            else:
                translated_parts = translated_parts[:len(texts)]

        # Delay to avoid exceeding API limits
        time.sleep(API_DELAY)
        return translated_parts

    except Exception as e:
        print(f"❌ Error translating batch: {str(e)}")
        # Return original texts if translation fails
        return texts

def process_excel(input_path, output_dir, source_lang, target_lang):
    """Process Excel file: read, translate and save with original format"""
    try:
        # Create output file path
        filename = os.path.basename(input_path)
        base_name, ext = os.path.splitext(filename)

        # Use provided output_dir
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, f"{base_name}-{target_lang}{ext}")

        print(f"\n🔄 Processing file: {filename}")

        # Open workbook with xlwings to preserve formatting
        app = xw.App(visible=False)
        wb = None # Initialize wb
        try:
            wb = app.books.open(input_path)

            # Loop through each sheet
            for sheet in wb.sheets:
                print(f"\n📋 Processing sheet: {sheet.name}")

                # Collect data from cells that need translation
                texts_to_translate = []
                cell_references = []

                # Scan through used data range
                used_rng = sheet.used_range
                if used_rng.count > 1 or used_rng.value is not None: # Only scan if sheet has data
                    for cell in used_rng:
                        # Check if cell.value is None before calling str()
                        cell_value_str = str(cell.value) if cell.value is not None else ""
                        if cell_value_str and should_translate(cell_value_str):
                            texts_to_translate.append(clean_text(cell_value_str))
                            cell_references.append(cell)
                else:
                     print(f"   ⚠️ Sheet '{sheet.name}' is empty or has no data.")

                # --- START SHAPES PROCESSING FIX ---
                # Process shapes with text
                try:
                    shapes_collection = sheet.api.Shapes
                    shapes_count = shapes_collection.Count

                    if shapes_count > 0:
                        print(f"   📊 Sheet '{sheet.name}' has {shapes_count} shapes to check")

                        # Process each shape by index (Excel COM API indexes from 1)
                        for i in range(1, shapes_count + 1):
                            shape = None # Initialize to avoid errors if .Item(i) fails
                            try:
                                shape = shapes_collection.Item(i)
                                shape_text = None

                                # --- Try multiple methods to get text from shape ---
                                
                                # Method 1: TextFrame
                                try:
                                    if hasattr(shape, 'TextFrame'):
                                        if shape.TextFrame.HasText:
                                            shape_text = shape.TextFrame.Characters().Text
                                except:
                                    pass
                                
                                # Method 2: TextFrame2
                                if not shape_text:
                                    try:
                                        if hasattr(shape, 'TextFrame2'):
                                            shape_text = shape.TextFrame2.TextRange.Text
                                    except:
                                        pass
                                
                                # Method 3: AlternativeText
                                if not shape_text:
                                    try:
                                        if hasattr(shape, 'AlternativeText') and shape.AlternativeText:
                                            shape_text = shape.AlternativeText
                                    except:
                                        pass
                                
                                # Method 4: OLEFormat (for OLE objects)
                                if not shape_text:
                                    try:
                                        if hasattr(shape, 'OLEFormat') and hasattr(shape.OLEFormat, 'Object'):
                                            if hasattr(shape.OLEFormat.Object, 'Text'):
                                                shape_text = shape.OLEFormat.Object.Text
                                    except:
                                        pass
                                
                                # Method 5: TextEffect (for WordArt)
                                if not shape_text:
                                    try:
                                        if hasattr(shape, 'TextEffect') and hasattr(shape.TextEffect, 'Text'):
                                            shape_text = shape.TextEffect.Text
                                    except:
                                        pass
                                
                                # If text is found, add to translation list
                                if shape_text and should_translate(shape_text):
                                    clean_shape_text = clean_text(shape_text)
                                    print(f"   💬 Shape {i}: Found text: {clean_shape_text[:30]}...")
                                    texts_to_translate.append(clean_shape_text)
                                    
                                    # Save tuple with information for later updates:
                                    # ('shape', sheet object, shape index, list of methods tried)
                                    cell_references.append(('shape', sheet, i))

                            except Exception as outer_e:
                                # General error when processing shape
                                print(f"   ⚠️ Error processing shape {i}: {str(outer_e)}")
                                continue

                except Exception as e:
                    print(f"   ⚠️ Error processing shapes on sheet '{sheet.name}': {str(e)}")
                # --- END SHAPES PROCESSING FIX ---


                # Split into batches for processing
                if not texts_to_translate:
                     print(f"   ✅ No text to translate on sheet '{sheet.name}'.")
                     continue # Move to next sheet

                total_batches = (len(texts_to_translate) - 1) // BATCH_SIZE + 1
                print(f"   📦 Preparing to translate {len(texts_to_translate)} text segments in {total_batches} batches.")

                for i in range(0, len(texts_to_translate), BATCH_SIZE):
                    batch_texts = texts_to_translate[i:i+BATCH_SIZE]
                    batch_refs = cell_references[i:i+BATCH_SIZE]
                    current_batch_num = i // BATCH_SIZE + 1

                    print(f"   🔄 Translating batch {current_batch_num}/{total_batches} ({len(batch_texts)} texts)")

                    # Translate batch
                    translated_batch = translate_batch(batch_texts, source_lang, target_lang)

                    # Update translated content
                    print(f"   ✍️ Updating content for batch {current_batch_num}...")
                    for j, ref in enumerate(batch_refs):
                        # Check if index j is within translated_batch
                        if j < len(translated_batch) and translated_batch[j] is not None:
                            try:
                                # Update content for shape and cell
                                if isinstance(ref, tuple) and ref[0] == 'shape':
                                    # Process shape: ref is ('shape', sheet_obj, shape_index)
                                    _, sheet_obj, shape_index = ref # Unpack tuple
                                    try:
                                        # Get shape object again
                                        shape_to_update = sheet_obj.api.Shapes.Item(shape_index)
                                        updated = False
                                        
                                        # --- Try multiple methods to update text for shape ---
                                        
                                        # Method 1: TextFrame
                                        try:
                                            if hasattr(shape_to_update, 'TextFrame') and shape_to_update.TextFrame.HasText:
                                                shape_to_update.TextFrame.Characters().Text = translated_batch[j]
                                                updated = True
                                        except:
                                            pass
                                            
                                        # Method 2: TextFrame2
                                        if not updated:
                                            try:
                                                if hasattr(shape_to_update, 'TextFrame2'):
                                                    shape_to_update.TextFrame2.TextRange.Text = translated_batch[j]
                                                    updated = True
                                            except:
                                                pass
                                                
                                        # Method 3: AlternativeText
                                        if not updated:
                                            try:
                                                if hasattr(shape_to_update, 'AlternativeText'):
                                                    shape_to_update.AlternativeText = translated_batch[j]
                                                    updated = True
                                            except:
                                                pass
                                                
                                        # Method 4: TextEffect (for WordArt)
                                        if not updated:
                                            try:
                                                if hasattr(shape_to_update, 'TextEffect') and hasattr(shape_to_update.TextEffect, 'Text'):
                                                    shape_to_update.TextEffect.Text = translated_batch[j]
                                                    updated = True
                                            except:
                                                pass
                                                
                                        # Method 5: OLEFormat
                                        if not updated:
                                            try:
                                                if hasattr(shape_to_update, 'OLEFormat') and hasattr(shape_to_update.OLEFormat, 'Object'):
                                                    if hasattr(shape_to_update.OLEFormat.Object, 'Text'):
                                                        shape_to_update.OLEFormat.Object.Text = translated_batch[j]
                                                        updated = True
                                            except:
                                                pass
                                                
                                        if updated:
                                            print(f"   ✅ Updated text for shape {shape_index} on sheet '{sheet_obj.name}'")
                                        else:
                                            print(f"   ⚠️ Could not update text for shape {shape_index} on sheet '{sheet_obj.name}' after trying all methods")
                                        
                                    except Exception as update_err:
                                        print(f"   ⚠️ Error updating shape {shape_index} on sheet '{sheet_obj.name}': {str(update_err)}")
                                elif isinstance(ref, xw.main.Range):
                                    # Is a cell
                                    ref.value = translated_batch[j]
                                else:
                                    print(f"   ⚠️ Unknown reference type: {type(ref)}")

                            except Exception as update_single_err:
                                # Catch general errors when updating a specific cell/shape
                                ref_info = f"Shape index {ref[2]} on sheet {ref[1].name}" if isinstance(ref, tuple) else f"Cell {ref.address}"
                                print(f"   ⚠️ Could not update content for {ref_info}: {str(update_single_err)}")
                        else:
                            # Notify if a translation is missing for a reference
                            ref_info = f"Shape index {ref[2]} on sheet {ref[1].name}" if isinstance(ref, tuple) else f"Cell {ref.address}"
                            print(f"   ⚠️ Missing translation for {ref_info} (index {j} in batch). Keeping original value.")


            # Save file with original format
            print(f"\n💾 Saving translated file to: {output_path}")
            wb.save(output_path)
            print(f"✅ File saved successfully: {output_path}")

        except Exception as wb_process_err:
             print(f"❌ Error processing workbook '{filename}': {str(wb_process_err)}")
             # Ensure workbook is closed if error occurs before saving
             if wb is not None:
                 try:
                     wb.close()
                 except Exception as close_err:
                     print(f"   ⚠️ Error trying to close workbook after processing error: {close_err}")
        finally:
            # Close workbook (if not already closed) and Excel app
            # wb.close() has been called in the except block if needed
            # Just need to ensure app is closed
            if 'app' in locals() and app.pid: # Check if app exists and is still running
                 app.quit()
                 print("🧊 Excel application closed.")

        return output_path

    except Exception as e:
        print(f"❌ Critical error when starting Excel file processing '{input_path}': {str(e)}")
        # Ensure Excel app is closed if error occurs right at the beginning
        if 'app' in locals() and app.pid:
            app.quit()
        return None

def process_directory(input_dir, output_dir, source_lang, target_lang):
    """Process all Excel files in the input directory"""
    # Ensure directory path exists
    if not os.path.isdir(input_dir):
        print(f"❌ Directory does not exist or is not a directory: {input_dir}")
        return

    # Find all Excel files in the directory (including .xls if needed)
    # Note: xlwings processing .xls files may require additional libraries or have limitations
    excel_files = glob.glob(os.path.join(input_dir, "*.xlsx")) + glob.glob(os.path.join(input_dir, "*.xls"))

    if not excel_files:
        print(f"⚠️ No Excel files (.xlsx, .xls) found in directory: {input_dir}")
        return

    print(f"🔍 Found {len(excel_files)} Excel files in input directory: {input_dir}")

    # Process each file
    successful_files = []
    failed_files = []
    for file_path in excel_files:
        # Skip Excel temporary files (usually starting with ~$)
        if os.path.basename(file_path).startswith('~$'):
            print(f"   ⏩ Skipping temporary file: {os.path.basename(file_path)}")
            continue

        output_file = process_excel(file_path, output_dir, source_lang, target_lang)
        if output_file:
            successful_files.append(os.path.basename(file_path))
        else:
            failed_files.append(os.path.basename(file_path))

    print(f"\n✅ Successful: {len(successful_files)} files")
    if failed_files:
        print(f"❌ Failed: {len(failed_files)} files: {', '.join(failed_files)}")

@Gooey(program_name="AI Excel Translator")
def main():
    # Create language help text dynamically
    lang_help = ', '.join(f"{code}: {name}" for code, name in lang_map.items())
    
    parser = GooeyParser(description='Translate Excel files from input directory to output directory')
    parser.add_argument('--input_dir', dest='input_dir', required=True,
                        help='Input directory containing Excel files to translate (required)',
                        widget='DirChooser')
    parser.add_argument('--output_dir', dest='output_dir', required=False, default='output',
                        help='Output directory for translated files (default: output folder in current directory)',
                        widget='DirChooser')
    parser.add_argument('--from', dest='source_lang', choices=lang_map.keys(), required=True, default="ja",
                        help=f'Source language ({lang_help})')
    parser.add_argument('--to', dest='target_lang', choices=lang_map.keys(), required=True, default="en",
                        help=f'Target language ({lang_help})')
    args = parser.parse_args()

    # Determine input directory
    input_dir = args.input_dir
    if not os.path.isabs(input_dir):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        input_dir = os.path.join(script_dir, input_dir)

    # Check if input directory exists
    if not os.path.exists(input_dir):
        os.makedirs(input_dir)
        print(f"📁 Created input directory at: {input_dir}")
        print("   Please place Excel files to translate in this directory.")
        return # Stop to let user add files
    print(f"📁 Input directory: '{input_dir}'")

    # Determine output directory
    output_dir = args.output_dir
    if output_dir is None:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(script_dir, "output")
    elif not os.path.isabs(output_dir):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(script_dir, output_dir)

    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    print(f"📂 Output directory: '{output_dir}'")

    print("")

    # Determine source and target language
    source_lang = lang_map.get(args.source_lang)
    target_lang = lang_map.get(args.target_lang)
    if not source_lang or not target_lang:
        print(f"❌ Invalid language combination: from {args.source_lang} to {args.target_lang}")
        return

    print(f"🎯 Translation direction: {source_lang} to {target_lang}")

    # Process all files in the input directory
    process_directory(input_dir, output_dir, args.source_lang, args.target_lang)

if __name__ == "__main__":
    # Note: Running this script may take time depending on the number of files and text to translate
    start_time = time.time()
    main()
    end_time = time.time()
    print(f"\n⏱️ Total execution time: {end_time - start_time:.2f} seconds\n")