import React, { useEffect, useState, useRef } from 'react';
import axios from 'axios';
import { HotTable } from '@handsontable/react';
import 'handsontable/dist/handsontable.full.min.css';
import './App.css';  // 引入新的样式文件
import * as XLSX from 'xlsx';
import { mean, std } from 'mathjs';

function App() {
  const [sheets, setSheets] = useState([]);
  const [data, setData] = useState([]);
  const [log, setLog] = useState('');
  const [selectedSheet, setSelectedSheet] = useState('');
  const [highlightChanges, setHighlightChanges] = useState(false);
  const [onlyPositiveAnomalies, setOnlyPositiveAnomalies] = useState(false);
  const [impactFactor, setImpactFactor] = useState(1.5);
  const [activeTab, setActiveTab] = useState('normal'); // 添加选项卡状态
  const [showRegressionAnomalies, setShowRegressionAnomalies] = useState(false);
  const tableRef = useRef(null);

  useEffect(() => {
    // 获取所有sheet名称
    axios.get('http://www.xfkenzify.com:12739/api/sheets')
      .then(response => {
        const sheetsData = response.data;
        setSheets(sheetsData);
        const lastSheet = sheetsData[sheetsData.length - 1];
        setSelectedSheet(lastSheet);
      })
      .catch(error => {
        console.error('Error fetching sheets:', error);
      });

    // 获取日志内容
    axios.get('http://www.xfkenzify.com:12739/api/log')
      .then(response => {
        console.log('Received log:', response.data);
        setLog(response.data);
      })
      .catch(error => {
        console.error('Error fetching log:', error);
      });
  }, []);

  useEffect(() => {
    // 获取选定sheet的数据
    if (selectedSheet) {
      axios.get(`http://www.xfkenzify.com:12739/api/data/${selectedSheet}`)
        .then(response => {
          setData(response.data);
        })
        .catch(error => {
          console.error('Error fetching sheet data:', error);
        });
    }
  }, [selectedSheet]);

  const handleSheetChange = (e) => {
    setSelectedSheet(e.target.value);
  };

  const handleHighlightChangesChange = (e) => {
    setHighlightChanges(e.target.checked);
  };

  const handleOnlyPositiveAnomaliesChange = (e) => {
    setOnlyPositiveAnomalies(e.target.checked);
  };

  const handleImpactFactorChange = (e) => {
    setImpactFactor(e.target.value);
  };

  const handleShowRegressionAnomaliesChange = (e) => {
    setShowRegressionAnomalies(e.target.checked);
  };

  const calculateColumnWidths = (data, columns) => {
    return columns.map(col => {
      const maxLength = data.reduce((max, row) => {
        const cellValue = row[col] ? row[col].toString() : '';
        return Math.max(max, cellValue.length);
      }, col.length);
      return Math.max(maxLength * 7, 50); // 基于字符数计算宽度，最小宽度为50
    });
  };

  const processGrowthStatistics = (data) => {
    return data.map(row => {
      const newRow = { ...row };
      Object.keys(row).forEach((col, index) => {
        if (index > 2) { // 跳过前三列
          newRow[col] = index === 3 ? 0 : row[Object.keys(row)[index]] - row[Object.keys(row)[index - 1]];
        }
      });
      return newRow;
    });
  };

  const processDifferenceComparison = (data) => {
    const growthData = processGrowthStatistics(data);
    return growthData.map(row => {
      const newRow = { ...row };
      Object.keys(row).forEach((col, index) => {
        if (index > 3) { // 跳过前四列
          newRow[col] = row[Object.keys(row)[index]] - row[Object.keys(row)[index - 1]];
        }
      });
      return newRow;
    });
  };

  const findRegressionAnomalies = (data) => {
    const anomalies = [];
    for (let i = 0; i < data.length; i += 3) {
      const group = data.slice(i, i + 3);
      const prevDistributions = group.map(row => {
        return Object.keys(row).slice(3).map(key => row[key]);
      });
      
      const avgDistribution = prevDistributions.reduce((acc, curr) => {
        return acc.map((val, idx) => val + curr[idx]);
      }, new Array(prevDistributions[0].length).fill(0)).map(val => val / prevDistributions.length);

      group.forEach((row, rowIndex) => {
        Object.keys(row).slice(3).forEach((key, colIndex) => {
          const currValue = row[key];
          const avgValue = avgDistribution[colIndex];
          const deviation = Math.abs(currValue - avgValue);
          if (deviation / avgValue > impactFactor) {
            anomalies.push({ row: i + rowIndex, col: colIndex + 3 });
          }
        });
      });
    }
    return anomalies;
  };

  const downloadExcel = () => {
    const normalData = data;
    const growthData = processGrowthStatistics(data);
    const differenceData = processDifferenceComparison(data);

    const combinedData = [
      ...normalData,
      {},
      ...growthData,
      {},
      ...differenceData
    ];

    const worksheet = XLSX.utils.json_to_sheet(combinedData, { skipHeader: false });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');
    XLSX.writeFile(workbook, `${selectedSheet}_data.xlsx`);
  };

  let processedData = data;
  if (activeTab === 'growth') {
    processedData = processGrowthStatistics(data);
  } else if (activeTab === 'difference') {
    processedData = processDifferenceComparison(data);
  }

  const columns = data.length ? Object.keys(data[0]) : [];
  const colWidths = calculateColumnWidths(data, columns);
  const hotData = processedData.map(row => columns.map(column => row[column]));

  const regressionAnomalies = showRegressionAnomalies && activeTab === 'growth' ? findRegressionAnomalies(processedData) : [];

  const customRenderer = (hotInstance, td, row, col, prop, value, cellProperties) => {
    td.innerText = value !== null ? value : '';

    // 行背景颜色
    if (Math.floor(row / 3) % 2 === 0) {
      td.style.backgroundColor = '#444444';
    } else {
      td.style.backgroundColor = '#555555';
    }

    // 前三列文本颜色
    if (col < 3) {
      td.style.color = '#ffffff';
    }

    // 高亮异常变化
    if (highlightChanges && col > 2) { // 从第四列开始计算
      const prevValue = parseFloat(hotData[row][col - 1]);
      const currValue = parseFloat(hotData[row][col]);
      const prevPrevValue = parseFloat(hotData[row][col - 2]);
      const prevChange = prevValue - prevPrevValue;
      const currChange = currValue - prevValue;
      const changeRate = Math.abs(currChange) / Math.abs(prevChange);

      if (!isNaN(prevChange) && !isNaN(currChange) && prevChange !== 0) {
        if (changeRate >= impactFactor) {
          if (!onlyPositiveAnomalies || (onlyPositiveAnomalies && currChange > 0)) {
            const normalizedRate = Math.min(changeRate / impactFactor, 1);
            const opacity = normalizedRate * 0.9 + 0.1;
            td.style.backgroundColor = `rgba(255, 0, 0, ${opacity})`;
          }
        }
      }
    }

    // 显示回归异常
    if (showRegressionAnomalies && regressionAnomalies.some(anomaly => anomaly.row === row && anomaly.col === col)) {
      td.style.border = '2px solid yellow';
    }

    // 禁止换行并使用省略号
    td.style.whiteSpace = 'nowrap';
    td.style.overflow = 'hidden';
    td.style.textOverflow = 'ellipsis';

    return td;
  };

  return (
    <div className="app-container">
      <div className="table-container">
        <h1>HaiYangYiDeng™ Excel File Reader And Analyzer</h1>
        <div className="controls-row">
          <div className="tab-selector">
            <button onClick={() => setActiveTab('normal')} className={`tab-button ${activeTab === 'normal' ? 'active' : ''}`}>原始表格模式</button>
            <button onClick={() => setActiveTab('growth')} className={`tab-button ${activeTab === 'growth' ? 'active' : ''}`}>增长数量模式</button>
            <button onClick={() => setActiveTab('difference')} className={`tab-button ${activeTab === 'difference' ? 'active' : ''}`}>增长差值模式</button>
          </div>
          <div className="download-button-container">
            <button onClick={downloadExcel} className="download-button">下载Excel文件</button>
          </div>
        </div>
        <div className="controls-row">
          <div className="sheet-selector">
            <label htmlFor="sheet-select">期数选择: </label>
            <select
              id="sheet-select"
              value={selectedSheet}
              onChange={handleSheetChange}
            >
              {sheets.map(sheet => (
                <option key={sheet} value={sheet}>
                  {sheet}
                </option>
              ))}
            </select>
          </div>
          <div className="slider-container">
            <label>
              异常变化显示模式:
              <input
                type="checkbox"
                checked={highlightChanges}
                onChange={handleHighlightChangesChange}
              />
            </label>
            {activeTab === 'growth' && (
              <label style={{ marginLeft: '10px' }}>
                回归异常:
                <input
                  type="checkbox"
                  checked={showRegressionAnomalies}
                  onChange={handleShowRegressionAnomaliesChange}
                />
              </label>
            )}
            {activeTab === 'difference' && (
              <label style={{ marginLeft: '10px' }}>
                只显示正异常:
                <input
                  type="checkbox"
                  checked={onlyPositiveAnomalies}
                  onChange={handleOnlyPositiveAnomaliesChange}
                />
              </label>
            )}
            <label>
              显示逻辑权重:
              <input
                type="range"
                min="1"
                max="10"
                step="0.01"
                value={impactFactor}
                onChange={handleImpactFactorChange}
              />
              <input
                type="number"
                min="1"
                max="10"
                step="0.01"
                value={impactFactor}
                onChange={handleImpactFactorChange}
              />
            </label>
          </div>
        </div>
        <div className="table-wrapper">
          <HotTable
            ref={tableRef}
            data={hotData}
            colHeaders={columns}
            rowHeaders={true}
            width="100%"
            height="100%"
            licenseKey="non-commercial-and-evaluation"
            stretchH="all"
            contextMenu={true}
            manualColumnResize={true}
            manualRowResize={true}
            fillHandle={true}
            copyPaste={true}
            colWidths={colWidths}
            className="handsontable-root"
            autoRowSize={false}
            autoColumnSize={false}
            fixedColumnsLeft={3}
            viewportRowRenderingOffset={20} // 提前渲染20行数据
            viewportColumnRenderingOffset={10} // 提前渲染10列数据
            cells={(row, col) => {
              const cellProperties = {};
              cellProperties.renderer = customRenderer;
              return cellProperties;
            }}
          />
        </div>
      </div>
      <div className="log-container">
        <h1>Scraping Log</h1>
        <pre>{log}</pre>
      </div>
    </div>
  );
}

export default App;
