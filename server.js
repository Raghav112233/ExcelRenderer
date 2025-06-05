const express = require('express');
const fs = require('fs');
const path = require('path');
const app = express();

const PORT = process.env.PORT || 3000;

// Serve index.html and static files
app.use(express.static(__dirname));

// Route to list .xlsx files
app.get('/excel-files', (req, res) => {
  fs.readdir(__dirname, (err, files) => {
    if (err) return res.status(500).json({ error: 'Failed to read directory' });
    const excelFiles = files.filter(file => file.endsWith('.xlsx'));
    res.json(excelFiles);
  });
});

// Route to serve .xlsx files directly
app.get('/:filename', (req, res) => {
  const filePath = path.join(__dirname, req.params.filename);
  if (fs.existsSync(filePath)) {
    res.sendFile(filePath);
  } else {
    res.status(404).send('File not found');
  }
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
