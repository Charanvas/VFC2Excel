import os
import re
import pandas as pd
from flask import Flask, request, render_template, send_file, jsonify, session
from werkzeug.utils import secure_filename
import json
import uuid
from datetime import datetime
import logging

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this-in-production'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Create directories if they don't exist
UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
TEMP_FOLDER = 'temp'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
os.makedirs(TEMP_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
app.config['TEMP_FOLDER'] = TEMP_FOLDER

class VCFParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.contacts = []
        self.all_fields = set()
        self.field_mapping = self._get_field_mapping()
    
    def _get_field_mapping(self):
        """Define automatic field mapping for common VCF properties"""
        return {
            'FN': 'Full Name',
            'N': ['Last Name', 'First Name', 'Middle Name', 'Name Prefix', 'Name Suffix'],
            'ORG': 'Organization',
            'TITLE': 'Job Title',
            'TEL': 'Phone',
            'EMAIL': 'Email',
            'URL': 'Website',
            'BDAY': 'Birthday',
            'NOTE': 'Notes',
            'ADR': ['PO Box', 'Extended Address', 'Street Address', 'City', 'State/Province', 'Postal Code', 'Country'],
            'NICKNAME': 'Nickname',
            'CATEGORIES': 'Categories',
            'X-ANNIVERSARY': 'Anniversary',
            'X-MANAGER': 'Manager',
            'X-ASSISTANT': 'Assistant',
            'X-SPOUSE': 'Spouse',
            'ROLE': 'Role',
            'REV': 'Last Modified'
        }
    
    def parse(self):
        """Parse VCF file and extract all contacts"""
        try:
            # Try different encodings
            content = None
            encodings = ['utf-8', 'utf-8-sig', 'latin1', 'cp1252', 'iso-8859-1']
            
            for encoding in encodings:
                try:
                    with open(self.file_path, 'r', encoding=encoding) as file:
                        content = file.read()
                    logger.info(f"Successfully read file with {encoding} encoding")
                    break
                except UnicodeDecodeError:
                    continue
            
            if content is None:
                raise ValueError("Could not decode the VCF file with any supported encoding")
            
            # Normalize line endings
            content = content.replace('\r\n', '\n').replace('\r', '\n')
            
            # Split content into individual vCards
            vcard_pattern = r'BEGIN:VCARD.*?END:VCARD'
            vcards = re.findall(vcard_pattern, content, re.DOTALL | re.IGNORECASE)
            
            logger.info(f"Found {len(vcards)} vCards in file")
            
            for i, vcard in enumerate(vcards):
                try:
                    contact = self.parse_vcard(vcard)
                    if contact:
                        self.contacts.append(contact)
                        logger.debug(f"Parsed contact {i+1}: {contact.get('Full Name', 'Unknown')}")
                except Exception as e:
                    logger.error(f"Error parsing vCard {i+1}: {str(e)}")
                    continue
            
            logger.info(f"Successfully parsed {len(self.contacts)} contacts")
            return self.contacts
            
        except Exception as e:
            logger.error(f"Error parsing VCF file: {str(e)}")
            raise
    
    def parse_vcard(self, vcard_text):
        """Parse individual vCard and extract fields"""
        contact = {}
        lines = vcard_text.strip().split('\n')
        
        # Handle line continuations
        processed_lines = []
        current_line = ""
        
        for line in lines:
            line = line.strip()
            if line.startswith(' ') or line.startswith('\t'):
                # This is a continuation of the previous line
                current_line += line[1:]  # Remove the leading space/tab
            else:
                if current_line:
                    processed_lines.append(current_line)
                current_line = line
        
        if current_line:
            processed_lines.append(current_line)
        
        for line in processed_lines:
            if not line or line.upper().startswith('BEGIN:VCARD') or line.upper().startswith('END:VCARD'):
                continue
            
            if ':' not in line:
                continue
            
            try:
                # Split property and value
                property_part, value = line.split(':', 1)
                value = value.strip()
                
                if not value:
                    continue
                
                # Parse property parameters
                if ';' in property_part:
                    property_name = property_part.split(';')[0].upper()
                    parameters = [p.upper() for p in property_part.split(';')[1:]]
                else:
                    property_name = property_part.upper()
                    parameters = []
                
                # Apply field mapping and extract data
                self._extract_field_data(contact, property_name, value, parameters)
                
            except Exception as e:
                logger.error(f"Error processing line '{line}': {str(e)}")
                continue
        
        return contact if contact else None
    
    def _extract_field_data(self, contact, property_name, value, parameters):
        """Extract and map field data based on property type"""
        
        if property_name == 'N':
            # Name: Last;First;Middle;Prefix;Suffix
            name_parts = [part.strip() for part in value.split(';')]
            name_fields = ['Last Name', 'First Name', 'Middle Name', 'Name Prefix', 'Name Suffix']
            
            for i, field in enumerate(name_fields):
                if i < len(name_parts) and name_parts[i]:
                    contact[field] = name_parts[i]
                    self.all_fields.add(field)
        
        elif property_name == 'FN':
            contact['Full Name'] = value
            self.all_fields.add('Full Name')
        
        elif property_name == 'TEL':
            # Handle phone numbers with types
            phone_type = self._get_phone_type(parameters)
            field_name = f'Phone ({phone_type})' if phone_type != 'Phone' else 'Phone'
            
            # Clean phone number
            clean_phone = re.sub(r'[^\d\+\-\(\)\s]', '', value)
            contact[field_name] = clean_phone
            self.all_fields.add(field_name)
        
        elif property_name == 'EMAIL':
            # Handle emails with types
            email_type = self._get_email_type(parameters)
            field_name = f'Email ({email_type})' if email_type != 'Email' else 'Email'
            
            contact[field_name] = value.lower()
            self.all_fields.add(field_name)
        
        elif property_name == 'ADR':
            # Address: PO Box;Extended;Street;City;State;Postal Code;Country
            addr_parts = [part.strip() for part in value.split(';')]
            addr_fields = ['PO Box', 'Extended Address', 'Street Address', 'City', 'State/Province', 'Postal Code', 'Country']
            
            addr_type = self._get_address_type(parameters)
            
            for i, field in enumerate(addr_fields):
                if i < len(addr_parts) and addr_parts[i]:
                    field_name = f'{field} ({addr_type})' if addr_type != 'Address' else field
                    contact[field_name] = addr_parts[i]
                    self.all_fields.add(field_name)
        
        elif property_name == 'ORG':
            # Organization can have department separated by ;
            org_parts = [part.strip() for part in value.split(';')]
            contact['Organization'] = org_parts[0]
            self.all_fields.add('Organization')
            
            if len(org_parts) > 1 and org_parts[1]:
                contact['Department'] = org_parts[1]
                self.all_fields.add('Department')
        
        elif property_name == 'TITLE':
            contact['Job Title'] = value
            self.all_fields.add('Job Title')
        
        elif property_name == 'BDAY':
            # Handle different date formats
            birthday = self._format_date(value)
            contact['Birthday'] = birthday
            self.all_fields.add('Birthday')
        
        elif property_name == 'NOTE':
            contact['Notes'] = value
            self.all_fields.add('Notes')
        
        elif property_name == 'URL':
            contact['Website'] = value
            self.all_fields.add('Website')
        
        elif property_name == 'NICKNAME':
            contact['Nickname'] = value
            self.all_fields.add('Nickname')
        
        elif property_name == 'CATEGORIES':
            contact['Categories'] = value
            self.all_fields.add('Categories')
        
        elif property_name.startswith('X-'):
            # Handle extended properties
            field_name = property_name.replace('X-', '').replace('-', ' ').title()
            contact[field_name] = value
            self.all_fields.add(field_name)
        
        else:
            # Handle other properties
            field_name = property_name.replace('-', ' ').title()
            contact[field_name] = value
            self.all_fields.add(field_name)
    
    def _get_phone_type(self, parameters):
        """Determine phone type from parameters"""
        type_mapping = {
            'HOME': 'Home',
            'WORK': 'Work',
            'CELL': 'Mobile',
            'MOBILE': 'Mobile',
            'FAX': 'Fax',
            'PAGER': 'Pager',
            'VOICE': 'Voice',
            'MAIN': 'Main'
        }
        
        for param in parameters:
            if param in type_mapping:
                return type_mapping[param]
        
        return 'Phone'
    
    def _get_email_type(self, parameters):
        """Determine email type from parameters"""
        type_mapping = {
            'HOME': 'Home',
            'WORK': 'Work',
            'INTERNET': 'Internet'
        }
        
        for param in parameters:
            if param in type_mapping:
                return type_mapping[param]
        
        return 'Email'
    
    def _get_address_type(self, parameters):
        """Determine address type from parameters"""
        type_mapping = {
            'HOME': 'Home',
            'WORK': 'Work'
        }
        
        for param in parameters:
            if param in type_mapping:
                return type_mapping[param]
        
        return 'Address'
    
    def _format_date(self, date_str):
        """Format date string to readable format"""
        try:
            # Common VCF date formats
            formats = ['%Y%m%d', '%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y']
            
            for fmt in formats:
                try:
                    date_obj = datetime.strptime(date_str, fmt)
                    return date_obj.strftime('%Y-%m-%d')
                except ValueError:
                    continue
            
            # If no format matches, return as is
            return date_str
        except:
            return date_str
    
    def get_all_fields(self):
        """Get all unique fields found in the VCF file"""
        return sorted(list(self.all_fields))
    
    def get_field_suggestions(self):
        """Get suggested field groupings for better organization"""
        suggestions = {
            'Name Fields': [f for f in self.all_fields if any(keyword in f.lower() for keyword in ['name', 'nickname'])],
            'Contact Fields': [f for f in self.all_fields if any(keyword in f.lower() for keyword in ['phone', 'email', 'website'])],
            'Address Fields': [f for f in self.all_fields if any(keyword in f.lower() for keyword in ['address', 'street', 'city', 'state', 'postal', 'country'])],
            'Work Fields': [f for f in self.all_fields if any(keyword in f.lower() for keyword in ['organization', 'title', 'job', 'department', 'work'])],
            'Personal Fields': [f for f in self.all_fields if any(keyword in f.lower() for keyword in ['birthday', 'anniversary', 'note', 'categories'])],
            'Other Fields': []
        }
        
        # Add remaining fields to 'Other Fields'
        used_fields = set()
        for fields in suggestions.values():
            used_fields.update(fields)
        
        suggestions['Other Fields'] = [f for f in self.all_fields if f not in used_fields]
        
        # Remove empty categories
        return {k: v for k, v in suggestions.items() if v}

def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'vcf'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file selected'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not file or not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file format. Please upload a .vcf file.'}), 400
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        logger.info(f"File saved: {filepath}")
        
        # Parse VCF file
        parser = VCFParser(filepath)
        contacts = parser.parse()
        all_fields = parser.get_all_fields()
        field_suggestions = parser.get_field_suggestions()
        
        if not contacts:
            os.remove(filepath)
            return jsonify({'error': 'No valid contacts found in the VCF file'}), 400
        
        # Store data temporarily
        session_id = str(uuid.uuid4())
        temp_data = {
            'filename': filename,
            'contacts': contacts,
            'all_fields': all_fields,
            'field_suggestions': field_suggestions
        }
        
        temp_file = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}.json')
        with open(temp_file, 'w', encoding='utf-8') as f:
            json.dump(temp_data, f, ensure_ascii=False, indent=2)
        
        # Clean up uploaded file
        os.remove(filepath)
        
        logger.info(f"Successfully processed {len(contacts)} contacts")
        
        return jsonify({
            'success': True,
            'session_id': session_id,
            'filename': filename,
            'contacts_count': len(contacts),
            'available_fields': all_fields,
            'field_suggestions': field_suggestions,
            'preview_contacts': contacts[:3]  # Send first 3 contacts for preview
        })
        
    except Exception as e:
        logger.error(f"Error in upload_file: {str(e)}")
        # Clean up files in case of error
        if 'filepath' in locals() and os.path.exists(filepath):
            os.remove(filepath)
        return jsonify({'error': f'Error processing file: {str(e)}'}), 500

@app.route('/convert', methods=['POST'])
def convert_to_excel():
    try:
        data = request.json
        session_id = data.get('session_id')
        selected_fields = data.get('selected_fields', [])
        
        if not session_id:
            return jsonify({'error': 'Missing session data'}), 400
        
        # Load temporary data
        temp_file = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}.json')
        if not os.path.exists(temp_file):
            return jsonify({'error': 'Session expired. Please upload the file again.'}), 400
        
        with open(temp_file, 'r', encoding='utf-8') as f:
            temp_data = json.load(f)
        
        contacts = temp_data['contacts']
        filename = temp_data['filename']
        
        if not contacts:
            return jsonify({'error': 'No contacts data found'}), 400
        
        # Create DataFrame
        if selected_fields:
            # Filter contacts to include only selected fields
            filtered_contacts = []
            for contact in contacts:
                filtered_contact = {}
                for field in selected_fields:
                    filtered_contact[field] = contact.get(field, '')
                filtered_contacts.append(filtered_contact)
            df = pd.DataFrame(filtered_contacts)
        else:
            # Use all available fields
            df = pd.DataFrame(contacts)
        
        # Fill NaN values with empty strings
        df = df.fillna('')
        
        # Generate Excel filename
        excel_filename = os.path.splitext(filename)[0] + '.xlsx'
        excel_filepath = os.path.join(app.config['DOWNLOAD_FOLDER'], excel_filename)
        
        # Save to Excel with formatting
        with pd.ExcelWriter(excel_filepath, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Contacts', index=False)
            
            # Get the workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Contacts']
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                
                adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # Style header row
            from openpyxl.styles import Font, PatternFill, Alignment
            
            header_font = Font(bold=True, color='FFFFFF')
            header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
            header_alignment = Alignment(horizontal='center', vertical='center')
            
            for cell in worksheet[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
        
        # Clean up temporary file
        os.remove(temp_file)
        
        logger.info(f"Excel file created: {excel_filepath}")
        
        return jsonify({
            'success': True,
            'excel_filename': excel_filename,
            'download_url': f'/download/{excel_filename}',
            'records_count': len(df)
        })
        
    except Exception as e:
        logger.error(f"Error in convert_to_excel: {str(e)}")
        return jsonify({'error': f'Error converting to Excel: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        filepath = os.path.join(app.config['DOWNLOAD_FOLDER'], filename)
        if os.path.exists(filepath):
            return send_file(filepath, as_attachment=True, download_name=filename)
        else:
            return jsonify({'error': 'File not found'}), 404
    except Exception as e:
        logger.error(f"Error downloading file: {str(e)}")
        return jsonify({'error': f'Error downloading file: {str(e)}'}), 500

@app.route('/preview/<session_id>')
def preview_data(session_id):
    """Preview the parsed data before conversion"""
    try:
        temp_file = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}.json')
        if not os.path.exists(temp_file):
            return jsonify({'error': 'Session expired'}), 400
        
        with open(temp_file, 'r', encoding='utf-8') as f:
            temp_data = json.load(f)
        
        return jsonify({
            'success': True,
            'contacts': temp_data['contacts'][:10],  # First 10 contacts
            'total_count': len(temp_data['contacts'])
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, threaded=True)