const express = require('express');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const cors = require('cors');
const app = express();
const port = 12739;

app.use(cors());

// 获取所有sheet名称的API端点
app.get('/api/sheets', (req, res) => {
  const filePath = path.join(__dirname, 'data.xlsx');
  if (!fs.existsSync(filePath)) {
    return res.status(404).send('File not found');
  }

  const fileBuffer = fs.readFileSync(filePath);
  const workbook = XLSX.read(fileBuffer, { type: 'buffer' });

  const sheetNames = workbook.SheetNames;

  res.json(sheetNames);
});

// 获取指定sheet数据的API端点
app.get('/api/data/:sheetName', (req, res) => {
  const filePath = path.join(__dirname, 'data.xlsx');
  if (!fs.existsSync(filePath)) {
    return res.status(404).send('File not found');
  }

  const sheetName = req.params.sheetName;
  const fileBuffer = fs.readFileSync(filePath);
  const workbook = XLSX.read(fileBuffer, { type: 'buffer' });

  if (!workbook.SheetNames.includes(sheetName)) {
    return res.status(404).send('Sheet not found');
  }

  const worksheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(worksheet);

  console.log(`Excel data from sheet ${sheetName}:`, jsonData);

  res.json(jsonData);
});

// 获取日志的API端点
app.get('/api/log', (req, res) => {
  const logPath = path.join(__dirname, 'scraping_log.txt');
  if (!fs.existsSync(logPath)) {
    return res.status(404).send('Log file not found');
  }

  const logContent = fs.readFileSync(logPath, 'utf-8');
  res.send(logContent);
});

app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
