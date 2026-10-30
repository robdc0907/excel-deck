document.addEventListener('DOMContentLoaded', () => {
  const uploadForm = document.getElementById('uploadForm');
  const fileInput = document.getElementById('fileInput');
  const sheetSelect = document.getElementById('sheetSelect');
  const columnList = document.getElementById('columnList');
  const statusBox = document.getElementById('statusBox');
  const generateButton = document.getElementById('generateButton');
  const titleInput = document.getElementById('titleInput');

  let uploadedFileName = '';
  let currentSheet = '';
  let currentHeaders = [];

  const setStatus = (message, type = 'info') => {
    statusBox.textContent = message;
    statusBox.className = `status ${type}`;
  };

  uploadForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    const file = fileInput.files[0];
    if (!file) {
      setStatus('Please choose an Excel file first.', 'error');
      return;
    }

    const formData = new FormData();
    formData.append('file', file);

    setStatus('Uploading file...', 'info');

    try {
      const response = await fetch('/api/upload', {
        method: 'POST',
        body: formData,
      });

      const result = await response.json();
      if (!response.ok) {
        throw new Error(result.detail || 'Upload failed');
      }

      uploadedFileName = result.filename;
      sheetSelect.innerHTML = '';

      result.sheets.forEach((sheet) => {
        const option = document.createElement('option');
        option.value = sheet.name;
        option.textContent = `${sheet.name} (${sheet.rowCount} rows)`;
        sheetSelect.appendChild(option);
      });

      currentSheet = result.sheets[0].name;
      currentHeaders = result.sheets[0].headers || [];
      renderColumnControls(currentHeaders);
      setStatus('File uploaded successfully. Choose columns and formulas.', 'success');
    } catch (error) {
      setStatus(error.message, 'error');
    }
  });

  sheetSelect.addEventListener('change', (event) => {
    currentSheet = event.target.value;
    const selectedSheet = JSON.parse(localStorage.getItem('sheetData') || '[]').find((sheet) => sheet.name === currentSheet);
    if (selectedSheet) {
      currentHeaders = selectedSheet.headers || [];
      renderColumnControls(currentHeaders);
    }
  });

  const renderColumnControls = (headers) => {
    columnList.innerHTML = '';
    headers.forEach((header) => {
      const row = document.createElement('div');
      row.className = 'column-row';

      const label = document.createElement('label');
      label.textContent = header;

      const select = document.createElement('select');
      select.innerHTML = `
        <option value="direct">Direct</option>
        <option value="subtraction">Subtraction</option>
        <option value="multiplication">Multiplication</option>
        <option value="growth_percent">Growth %</option>
        <option value="percentage_of_total">Percentage of total</option>
      `;

      row.appendChild(label);
      row.appendChild(select);
      columnList.appendChild(row);
    });
  };

  generateButton.addEventListener('click', async () => {
    const selectedColumns = [];
    [...columnList.children].forEach((row) => {
      const label = row.querySelector('label');
      const formula = row.querySelector('select').value;
      if (label && formula) {
        selectedColumns.push({ name: label.textContent, formula });
      }
    });

    if (!uploadedFileName || selectedColumns.length === 0) {
      setStatus('Please upload a file and select at least one column.', 'error');
      return;
    }

    try {
      setStatus('Generating PowerPoint deck...', 'info');
      const response = await fetch('/api/generate-ppt', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileName: uploadedFileName,
          sheetName: currentSheet,
          columns: selectedColumns,
          title: titleInput.value || 'Executive Summary',
        }),
      });

      if (!response.ok) {
        const error = await response.json();
        throw new Error(error.detail || 'Presentation generation failed');
      }

      const blob = await response.blob();
      const downloadUrl = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = 'excel-deck-presentation.pptx';
      link.click();

      setStatus('PowerPoint created successfully and downloaded.', 'success');
    } catch (error) {
      setStatus(error.message, 'error');
    }
  });
});
