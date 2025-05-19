import pandas as pd
import os
from datetime import datetime
import fitz  # PyMuPDF
import yaml
import sys


def debug_pdf_info(pdf_path):
    """
    Print information about the PDF to help with debugging
    """
    print(f"Analyzing PDF: {pdf_path}")

    # Using PyMuPDF (fitz) to get more detailed info
    doc = fitz.open(pdf_path)
    print(f"Number of pages: {doc.page_count}")

    # Get page dimensions for the first page
    page = doc[0]
    print(f"Page dimensions: {page.rect}")
    doc.close()


def add_text_overlay(input_pdf_path, output_pdf_path, field_values, field_configs):
    """
    Add text as an overlay on the PDF using PyMuPDF with configs from YAML
    Using a multi-method approach to try to get custom fonts working

    Parameters:
    - input_pdf_path: Path to the template PDF
    - output_pdf_path: Where to save the filled PDF
    - field_values: Dictionary containing values for each field
    - field_configs: Dictionary containing configuration for each field
    """
    # Open the input PDF
    doc = fitz.open(input_pdf_path)
    page = doc[0]  # First page

    # Get page dimensions for calculating coordinates
    page_width = page.rect.width
    page_height = page.rect.height

    print(f"Adding text to PDF with dimensions: {page_width} x {page_height}")

    # Dictionary to cache font objects so we don't reload them for each field
    font_cache = {}

    # Process each field according to its configuration
    for field_name, config in field_configs.items():
        # Skip if this field doesn't have a value
        if field_name not in field_values:
            print(f"Warning: No value provided for field '{field_name}', skipping")
            continue

        field_value = field_values[field_name]
        if not field_value:  # Skip empty values
            continue

        # Get position values, calculating actual coordinates from percentages
        x_percent = config.get('x_percent', 50)
        y_percent = config.get('y_percent', 50)
        x = page_width * (x_percent / 100)
        y = page_height * (y_percent / 100)

        # Get font size
        font_size = config.get('font_size', 12)

        # Check for text alignment
        alignment = config.get('alignment', 'left')  # Default is left alignment

        # Check if a custom font was specified
        font_path = config.get('font')

        success = False

        # If a custom font is specified and exists, try various methods
        if font_path and os.path.exists(font_path):

            # Method 1: Try using reportlab to create a PDF with the custom font
            try:
                print(f"Trying to create a new PDF with the custom font for '{field_name}'")

                # Create a temporary PDF with the text in the custom font
                from reportlab.pdfgen import canvas as rl_canvas
                from reportlab.lib.pagesizes import letter
                from reportlab.pdfbase import pdfmetrics
                from reportlab.pdfbase.ttfonts import TTFont
                import tempfile

                # Create a temporary file
                temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
                temp_pdf.close()
                temp_pdf_path = temp_pdf.name

                # Register the font with reportlab
                font_name = os.path.splitext(os.path.basename(font_path))[0]
                try:
                    pdfmetrics.registerFont(TTFont(font_name, font_path))
                except:
                    # Fall back to a simpler font name if needed
                    font_name = "CustomFont"
                    pdfmetrics.registerFont(TTFont(font_name, font_path))

                # Create a canvas with the same dimensions as our target PDF
                c = rl_canvas.Canvas(temp_pdf_path, pagesize=(page_width, page_height))

                # Set the font and add the text
                c.setFont(font_name, font_size)

                # Handle text alignment
                if alignment == 'center':
                    # ReportLab can handle centering directly
                    # We need to convert from PyMuPDF coordinates to ReportLab coordinates
                    c.drawCentredString(
                        page_width / 2,  # x position (center of page)
                        page_height - y,  # ReportLab uses bottom-left origin
                        field_value
                    )
                else:
                    # Default left alignment
                    c.drawString(x, page_height - y, field_value)

                # Save the PDF
                c.save()

                # Now open this temporary PDF
                temp_doc = fitz.open(temp_pdf_path)
                if temp_doc.page_count > 0:
                    temp_page = temp_doc[0]

                    # Extract the text image from the temporary PDF
                    # We'll use insertPage instead of directly copying text
                    page.show_pdf_page(
                        page.rect,  # target rectangle
                        temp_doc,  # source document
                        0,  # source page number
                        keep_proportion=True,
                        overlay=True
                    )

                    print(f"Added text for field '{field_name}' with custom font using ReportLab")
                    success = True

                # Close and delete the temporary PDF
                temp_doc.close()
                try:
                    os.unlink(temp_pdf_path)
                except:
                    pass

            except Exception as e:
                print(f"ReportLab method failed: {e}")

            # If ReportLab method failed, try direct method
            if not success:
                try:
                    print(f"Trying direct insert_text with fontfile for '{field_name}'")

                    # For center alignment, calculate the position
                    if alignment == 'center':
                        # Simple estimation for text width
                        text_width = len(field_value) * font_size * 0.6
                        x = (page_width - text_width) / 2

                    # Make sure we're using a fully qualified path
                    abs_font_path = os.path.abspath(font_path)
                    page.insert_text(
                        (x, y),
                        field_value,
                        fontsize=font_size,
                        fontfile=abs_font_path
                    )
                    print(f"Added text for field '{field_name}' with direct fontfile method")
                    success = True
                except Exception as e:
                    print(f"Direct fontfile method failed: {e}")

        # If custom font methods failed or no custom font was specified, use default font
        if not success:
            # For center alignment, calculate the position
            if alignment == 'center':
                # Simple estimation for text width
                text_width = len(field_value) * font_size * 0.6
                x = (page_width - text_width) / 2

            page.insert_text(
                (x, y),
                field_value,
                fontsize=font_size
            )
            print(f"Added text for field '{field_name}' with default font")

    # Save the modified PDF
    doc.save(output_pdf_path)
    doc.close()

    print(f"Saved modified PDF to {output_pdf_path}")

def process_certificates(config):
    """
    Process all certificates from an Excel spreadsheet using YAML config

    Parameters:
    - config: Dictionary containing all configuration from YAML file
    """
    # Get basic paths
    excel_path = config.get('excel_path', 'refuge_names.xlsx')
    template_pdf_path = config.get('template_pdf_path', 'certificate_template.pdf')
    output_folder = config.get('output_folder', 'completed_certificates')

    # Get field configurations
    field_configs = config.get('fields', {})

    # First, debug the PDF to understand its structure
    debug_pdf_info(template_pdf_path)

    # Read data from Excel
    try:
        df = pd.read_excel(excel_path)
        print(f"Successfully read Excel file: {excel_path}")
        print(f"Excel columns found: {', '.join(df.columns)}")
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)

    # Create output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # Format today's date (used as fallback if date not in Excel)
    today = datetime.now()
    formatted_today = today.strftime("%B %d %Y")

    # Process each row in the Excel file
    for index, row in list(df.iterrows()):
        # Prepare field values for this certificate
        field_values = {}

        # Map Excel columns to field values
        field_mappings = config.get('field_mappings', {})
        for field_name, excel_column in field_mappings.items():
            if excel_column and excel_column in row:
                # Handle date column specially
                if field_name == 'date' and pd.notna(row[excel_column]):
                    # Try to parse the date if it's not already a datetime object
                    if isinstance(row[excel_column], (datetime, pd.Timestamp)):
                        field_values[field_name] = row[excel_column].strftime("%B %d, %Y.")
                    else:
                        # Try to parse string as date
                        try:
                            date_value = pd.to_datetime(row[excel_column])
                            field_values[field_name] = date_value.strftime("%B %-d, %Y.")
                        except:
                            # If parsing fails, use the value as-is
                            field_values[field_name] = str(row[excel_column])
                else:
                    # For non-date fields, convert to string if not already
                    field_values[field_name] = str(row[excel_column]) if pd.notna(row[excel_column]) else ""
            else:
                # If date field is missing or empty in Excel, use today's date
                if field_name == 'date' and field_name in field_configs:
                    field_values[field_name] = formatted_today

        # Get person's name for filename
        name_field = config.get('name_field', 'full_name')
        if name_field in field_values and field_values[name_field]:
            person_name = field_values[name_field]
        else:
            person_name = f"person_{index}"

        # Create a unique output filename
        safe_name = ''.join(c if c.isalnum() else '_' for c in person_name)
        email = field_values['email']
        output_filename = f"{person_name}_{email}.pdf"
        output_path = os.path.join(output_folder, output_filename)

        print(f"\nProcessing certificate for {person_name}")

        # Use text overlay with configurations
        add_text_overlay(
            template_pdf_path,
            output_path,
            field_values,
            field_configs
        )

        print(f"Certificate created for {person_name}")

    print(f"Completed processing {len(df)} certificates.")


def process_single_test(config):
    """
    Process a single test certificate with placeholder data

    Parameters:
    - config: Dictionary containing all configuration from YAML file
    """
    # Get basic paths
    template_pdf_path = config.get('template_pdf_path', 'certificate_template.pdf')
    output_folder = config.get('output_folder', 'completed_certificates')

    # Get field configurations
    field_configs = config.get('fields', {})

    # Create output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # Debug the PDF
    debug_pdf_info(template_pdf_path)

    # Create test field values
    field_values = {}
    for field_name in field_configs.keys():
        field_values[field_name] = f"Test {field_name}"

    # Add date if configured
    if 'date' in field_configs:
        field_values['date'] = datetime.now().strftime("%B %d, %Y")

    # Output path
    output_path = os.path.join(output_folder, "test_certificate.pdf")

    # Use text overlay with configurations
    add_text_overlay(
        template_pdf_path,
        output_path,
        field_values,
        field_configs
    )

    print(f"Test certificate created at {output_path}")
    print("Review this file and adjust your YAML configuration as needed")


def load_config(config_path):
    """
    Load configuration from YAML file
    """
    try:
        with open(config_path, 'r', encoding='utf-8') as file:
            config = yaml.safe_load(file)
            print(f"Loaded configuration from {config_path}")
            return config
    except Exception as e:
        print(f"Error loading configuration from {config_path}: {e}")
        sys.exit(1)


def main():
    """
    Main function to run the script
    """
    # Default config path
    config_path = "certificate_config.yaml"

    # Check if a config path was provided
    if len(sys.argv) > 1:
        config_path = sys.argv[1]

    # Load configuration
    config = load_config(config_path)

    # Check if this is a test run
    test_mode = config.get('test_mode', False)

    if test_mode:
        print("Running in test mode with placeholder data...")
        process_single_test(config)
    else:
        print("Processing certificates from Excel data...")
        process_certificates(config)

    print(
        "\nIf the text positioning needs adjustment, modify the x_percent and y_percent values in your YAML configuration.")


if __name__ == "__main__":
    main()