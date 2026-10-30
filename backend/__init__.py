<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Excel Deck</title>
  <link rel="stylesheet" href="/static/styles.css" />
</head>
<body>
  <div class="container">
    <h1>Excel Deck</h1>
    <p>Upload an Excel file, choose your columns, apply formulas, and generate a PowerPoint presentation.</p>

    <form id="uploadForm">
      <div class="form-group">
        <label for="fileInput">Upload Excel file</label>
        <input id="fileInput" type="file" accept=".xlsx,.csv" />
      </div>
      <button type="submit">Upload file</button>
    </form>

    <div class="form-group">
      <label for="sheetSelect">Select sheet</label>
      <select id="sheetSelect"></select>
    </div>

    <div class="form-group">
      <label for="titleInput">Presentation title</label>
      <input id="titleInput" type="text" value="Executive Summary" />
    </div>

    <div>
      <h2>Column configuration</h2>
      <div id="columnList" class="column-list"></div>
    </div>

    <button id="generateButton" type="button">Generate PPT</button>
    <div id="statusBox" class="status info">Ready to upload your file.</div>
  </div>

  <script src="/static/app.js"></script>
</body>
</html>
