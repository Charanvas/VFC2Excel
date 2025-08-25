let uploadedData = null;
let sessionId = null;

// File input handling
document.getElementById('file-input').addEventListener('change', handleFileSelect);

// Drag and drop handling
const uploadArea = document.getElementById('upload-area');
uploadArea.addEventListener('dragover', handleDragOver);
uploadArea.addEventListener('drop', handleFileDrop);
uploadArea.addEventListener('dragleave', handleDragLeave);
uploadArea.addEventListener('click', () => document.getElementById('file-input').click());

// Button event listeners
document.getElementById('select-all-btn').addEventListener('click', selectAllFields);
document.getElementById('clear-all-btn').addEventListener('click', clearAllFields);
document.getElementById('convert-btn').addEventListener('click', convertToExcel);
document.getElementById('preview-btn').addEventListener('click', showFullPreview);

function handleDragOver(e) {
    e.preventDefault();
    uploadArea.classList.add('dragover');
}

function handleDragLeave(e) {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
}

function handleFileDrop(e) {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
}

function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        handleFile(file);
    }
}

function handleFile(file) {
    if (!file.name.toLowerCase().endsWith('.vcf')) {
        showError('Please select a valid .vcf file.');
        return;
    }

    const formData = new FormData();
    formData.append('file', file);

    showProgress('Uploading and parsing VCF file...');

    fetch('/upload', {
        method: 'POST',
        body: formData
    })
    .then(response => {
        if (!response.ok) {
            return response.json().then(err => Promise.reject(err));
        }
        return response.json();
    })
    .then(data => {
        if (data.success) {
            uploadedData = data;
            sessionId = data.session_id;
            showFieldSelection(data);
        } else {
            showError(data.error || 'Unknown error occurred');
        }
    })
    .catch(error => {
        console.error('Upload error:', error);
        showError(error.error || error.message || 'Error uploading file');
    });
}

function showProgress(message = 'Processing...') {
    hideAllSections();
    document.getElementById('progress-section').style.display = 'block';
    document.getElementById('progress-text').textContent = message;
}

function showFieldSelection(data) {
    hideAllSections();
    document.getElementById('field-selection-section').style.display = 'block';
    document.getElementById('file-info').style.display = 'block';
    
    // Update file info
    document.getElementById('info-filename').textContent = data.filename;
    document.getElementById('info-contacts').textContent = data.contacts_count;
    document.getElementById('info-fields').textContent = data.available_fields.length;
    document.getElementById('info-categories').textContent = Object.keys(data.field_suggestions).length;
    
    // Populate field categories
    populateFieldCategories(data.field_suggestions);
    
    // Show preview
    updatePreview(data.preview_contacts);
}

function populateFieldCategories(fieldSuggestions) {
    const container = document.getElementById('field-categories');
    container.innerHTML = '';
    
    Object.entries(fieldSuggestions).forEach(([category, fields]) => {
        if (fields.length === 0) return;
        
        const categoryDiv = document.createElement('div');
        categoryDiv.className = 'mb-4';
        
        // Category header with select all for category
        const headerDiv = document.createElement('div');
        headerDiv.className = 'd-flex justify-content-between align-items-center mb-2';
        
        const categoryTitle = document.createElement('h6');
        categoryTitle.className = 'mb-0';
        categoryTitle.innerHTML = `<i class="fas fa-folder"></i> ${category} (${fields.length})`;
        
        const categorySelectBtn = document.createElement('button');
        categorySelectBtn.className = 'btn btn-sm btn-outline-primary';
        categorySelectBtn.innerHTML = '<i class="fas fa-check"></i> Select Category';
        categorySelectBtn.onclick = () => selectCategoryFields(category);
        
        headerDiv.appendChild(categoryTitle);
        headerDiv.appendChild(categorySelectBtn);
        
        // Fields container
        const fieldsContainer = document.createElement('div');
        fieldsContainer.className = 'row';
        fieldsContainer.id = `category-${category.replace(/\s+/g, '-').toLowerCase()}`;
        
        fields.forEach(field => {
            const colDiv = document.createElement('div');
            colDiv.className = 'col-md-4 col-sm-6 mb-2';
            
            const checkDiv = document.createElement('div');
            checkDiv.className = 'form-check';
            
            const checkbox = document.createElement('input');
            checkbox.className = 'form-check-input field-checkbox';
            checkbox.type = 'checkbox';
            checkbox.id = 'field-' + field.replace(/[^a-zA-Z0-9]/g, '');
            checkbox.value = field;
            checkbox.dataset.category = category;
            checkbox.onchange = updatePreviewOnFieldChange;
            
            const label = document.createElement('label');
            label.className = 'form-check-label';
            label.htmlFor = checkbox.id;
            label.textContent = field;
            label.title = field; // Tooltip for long field names
            
            checkDiv.appendChild(checkbox);
            checkDiv.appendChild(label);
            colDiv.appendChild(checkDiv);
            fieldsContainer.appendChild(colDiv);
        });
        
        categoryDiv.appendChild(headerDiv);
        categoryDiv.appendChild(fieldsContainer);
        container.appendChild(categoryDiv);
    });
    
    // Auto-select common fields
    autoSelectCommonFields();
}

function autoSelectCommonFields() {
    const commonFields = [
        'Full Name', 'First Name', 'Last Name', 'Phone', 'Email', 
        'Organization', 'Job Title', 'Phone (Mobile)', 'Email (Work)'
    ];
    
    commonFields.forEach(fieldName => {
        const checkbox = document.querySelector(`input[value="${fieldName}"]`);
        if (checkbox) {
            checkbox.checked = true;
        }
    });
    
    updatePreviewOnFieldChange();
}

function selectCategoryFields(category) {
    const categoryCheckboxes = document.querySelectorAll(`input[data-category="${category}"]`);
    const allChecked = Array.from(categoryCheckboxes).every(cb => cb.checked);
    
    categoryCheckboxes.forEach(checkbox => {
        checkbox.checked = !allChecked;
    });
    
    updatePreviewOnFieldChange();
}

function selectAllFields() {
    const checkboxes = document.querySelectorAll('.field-checkbox');
    const allChecked = Array.from(checkboxes).every(cb => cb.checked);
    
    checkboxes.forEach(checkbox => {
        checkbox.checked = !allChecked;
    });
    
    // Update button text
    const btn = document.getElementById('select-all-btn');
    btn.innerHTML = allChecked ? 
        '<i class="fas fa-check-double"></i> Select All' : 
        '<i class="fas fa-times"></i> Deselect All';
    
    updatePreviewOnFieldChange();
}

function clearAllFields() {
    const checkboxes = document.querySelectorAll('.field-checkbox');
    checkboxes.forEach(checkbox => {
        checkbox.checked = false;
    });
    
    updatePreviewOnFieldChange();
}

function updatePreviewOnFieldChange() {
    const selectedFields = getSelectedFields();
    if (uploadedData && uploadedData.preview_contacts) {
        updatePreview(uploadedData.preview_contacts, selectedFields);
    }
}

function getSelectedFields() {
    const selectedFields = [];
    const checkedBoxes = document.querySelectorAll('.field-checkbox:checked');
    checkedBoxes.forEach(checkbox => {
        selectedFields.push(checkbox.value);
    });
    return selectedFields;
}

function updatePreview(contacts, selectedFields = null) {
    const headerRow = document.getElementById('preview-header');
    const bodyRows = document.getElementById('preview-body');
    
    if (!contacts || contacts.length === 0) {
        headerRow.innerHTML = '<th>No data available</th>';
        bodyRows.innerHTML = '';
        return;
    }
    
    // Determine which fields to show
    let fieldsToShow;
    if (selectedFields && selectedFields.length > 0) {
        fieldsToShow = selectedFields;
    } else {
        // Show all available fields from first contact
        fieldsToShow = Object.keys(contacts[0]).slice(0, 8); // Limit to 8 columns for preview
    }
    
    // Create header
    headerRow.innerHTML = '';
    fieldsToShow.forEach(field => {
        const th = document.createElement('th');
        th.textContent = field;
        th.style.fontSize = '0.8rem';
        headerRow.appendChild(th);
    });
    
    // Create body rows
    bodyRows.innerHTML = '';
    contacts.slice(0, 5).forEach((contact, index) => { // Show max 5 rows in preview
        const tr = document.createElement('tr');
        fieldsToShow.forEach(field => {
            const td = document.createElement('td');
            const value = contact[field] || '';
            td.textContent = value.length > 30 ? value.substring(0, 30) + '...' : value;
            td.title = value; // Full value in tooltip
            td.style.fontSize = '0.8rem';
            tr.appendChild(td);
        });
        bodyRows.appendChild(tr);
    });
}

function showFullPreview() {
    if (!sessionId) {
        showError('No session data available');
        return;
    }
    
    showProgress('Loading full preview...');
    
    fetch(`/preview/${sessionId}`)
    .then(response => response.json())
    .then(data => {
        hideAllSections();
        document.getElementById('field-selection-section').style.display = 'block';
        document.getElementById('file-info').style.display = 'block';
        
        if (data.success) {
            populateFullPreviewModal(data.contacts);
            const modal = new bootstrap.Modal(document.getElementById('previewModal'));
            modal.show();
        } else {
            showError(data.error);
        }
    })
    .catch(error => {
        showError('Error loading preview: ' + error.message);
    });
}

function populateFullPreviewModal(contacts) {
    const selectedFields = getSelectedFields();
    const fieldsToShow = selectedFields.length > 0 ? selectedFields : Object.keys(contacts[0] || {});
    
    const headerRow = document.getElementById('full-preview-header');
    const bodyRows = document.getElementById('full-preview-body');
    
    // Create header
    headerRow.innerHTML = '';
    fieldsToShow.forEach(field => {
        const th = document.createElement('th');
        th.textContent = field;
        headerRow.appendChild(th);
    });
    
    // Create body rows
    bodyRows.innerHTML = '';
    contacts.forEach(contact => {
        const tr = document.createElement('tr');
        fieldsToShow.forEach(field => {
            const td = document.createElement('td');
            td.textContent = contact[field] || '';
            tr.appendChild(td);
        });
        bodyRows.appendChild(tr);
    });
}

function convertToExcel() {
    if (!sessionId) {
        showError('No session data available. Please upload a file first.');
        return;
    }
    
    const selectedFields = getSelectedFields();
    
    showProgress('Converting to Excel...');
    
    const convertData = {
        session_id: sessionId,
        selected_fields: selectedFields
    };

    fetch('/convert', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(convertData)
    })
    .then(response => {
        if (!response.ok) {
            return response.json().then(err => Promise.reject(err));
        }
        return response.json();
    })
    .then(data => {
        if (data.success) {
            showDownload(data);
        } else {
            showError(data.error || 'Conversion failed');
        }
    })
    .catch(error => {
        console.error('Conversion error:', error);
        showError(error.error || error.message || 'Error converting file');
    });
}

function showDownload(data) {
    hideAllSections();
    document.getElementById('download-section').style.display = 'block';
    
    const downloadLink = document.getElementById('download-link');
    downloadLink.href = data.download_url;
    downloadLink.download = data.excel_filename;
    
    document.getElementById('records-count').textContent = data.records_count || 'Unknown';
}

function showError(message) {
    hideAllSections();
    document.getElementById('error-section').style.display = 'block';
    document.getElementById('error-message').textContent = message;
}

function hideAllSections() {
    const sections = [
        'upload-section', 'progress-section', 'field-selection-section',
        'download-section', 'error-section', 'file-info'
    ];
    
    sections.forEach(sectionId => {
        document.getElementById(sectionId).style.display = 'none';
    });
}

// Auto-download functionality
document.addEventListener('DOMContentLoaded', function() {
    // Check if there's a download link and auto-trigger download
    const downloadLink = document.getElementById('download-link');
    if (downloadLink && downloadLink.href && downloadLink.href !== window.location.href) {
        setTimeout(() => {
            downloadLink.click();
        }, 1000);
    }
});