# test
<html lang="zh">
<head>
  <meta charset="UTF-8" />
  <title>Decathlon Excel</title>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <style>
    body {
      font-family: Arial, sans-serif;
      margin: 40px;
      background-color: #f5f5f5;
    }
    h2 {
      color: #2c3e50;
    }
    #fileUpload {
      margin: 10px 0;
    }
    input[type="text"] {
      padding: 6px 10px;
      font-size: 16px;
      border-radius: 4px;
      border: 1px solid #ccc;
      width: 300px;
      margin-bottom: 10px;
    }
    button {
      background-color: #3498db;
      color: white;
      border: none;
      padding: 10px 20px;
      border-radius: 5px;
      cursor: pointer;
      font-size: 16px;
    }
    button:hover {
      background-color: #2980b9;
    }
    #status {
      margin-top: 15px;
      font-weight: bold;
      color: #27ae60;
    }
    .error {
      color: #e74c3c;
    }
    p.tip {
      font-size: 14px;
      color: #555;
    }
  </style>
</head>
<body>

  <h2>合并多个 Excel 文件</h2>
  <p class="tip">提示：可以按住 Ctrl 选择多个 Excel 文件一起上传</p>
  <input type="file" id="fileUpload" multiple accept=".xlsx, .xls" />
  <br />
  <input type="text" id="fileNameInput" placeholder="请输入合并后文件名（不带扩展名）" />
  <br />
  <button onclick="mergeExcels()">合并并导出 Excel</button>
  <p id="status"></p>

  <script>
    let allData = [];

    function mergeExcels() {
      const fileInput = document.getElementById('fileUpload');
      const files = fileInput.files;
      const status = document.getElementById('status');

      if (!files.length) {
        alert("请选择至少一个 Excel 文件！");
        return;
      }

      allData = [];
      status.innerHTML = "正在读取文件...";

      let filesProcessed = 0;
      let errorFiles = [];

      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const reader = new FileReader();

        reader.onload = function (e) {
          try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            workbook.SheetNames.forEach(sheetName => {
              // 跳过包含“样衣汇总”的工作表
              if (sheetName.includes("样衣汇总")) return;

              const worksheet = workbook.Sheets[sheetName];
              const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

              if (json.length > 0) {
                const year = parseInt(sheetName.substring(0, 4), 10);
                const monthMatch = sheetName.match(/(\d{1,2})月/);
                const month = monthMatch ? parseInt(monthMatch[1], 10) : 0;
                const type = sheetName;

                const annotatedJson = json.map(row => ({
                  Type: type,
                  Year: isNaN(year) ? "" : year,
                  Month: isNaN(month) ? "" : month,
                  ...row
                }));

                allData = allData.concat(annotatedJson);
              }
            });

          } catch (err) {
            errorFiles.push(file.name);
          }

          filesProcessed++;
          status.innerHTML = `已处理 ${filesProcessed} / ${files.length} 个文件...`;

          if (filesProcessed === files.length) {
            if (allData.length === 0) {
              status.innerHTML = `<span class="error">没有成功读取任何数据。</span>`;
              return;
            }

            if (errorFiles.length > 0) {
              status.innerHTML += `<br><span class="error">以下文件出错，已跳过：<br>${errorFiles.join('<br>')}</span>`;
            }

            exportMerged();
          }
        };

        reader.readAsArrayBuffer(file);
      }
    }

    function removeEmptyColumns(data) {
      if (data.length === 0) return data;
      const columns = Object.keys(data[0]);
      const emptyCols = columns.filter(col => data.every(row => {
        const val = row[col];
        return val === null || val === undefined || val === "";
      }));
      return data.map(row => {
        const newRow = {};
        for (const col of columns) {
          if (!emptyCols.includes(col)) newRow[col] = row[col];
        }
        return newRow;
      });
    }

    function exportMerged() {
      const cleanedData = removeEmptyColumns(allData);

      const wantedColumns = ["Type", "Year", "季节", "客户", "款号", "数量", "业务", "期段", "品牌"];

      // 保留 Month 字段用于排序
      const filteredData = cleanedData.map(row => {
        const newRow = {};
        wantedColumns.forEach(col => {
          newRow[col] = row[col] !== undefined ? row[col] : "";
        });
        newRow.Year = row.Year;
        newRow.Type = row.Type;
        newRow.Month = row.Month;
        return newRow;
      });

      // 先按 Year 再按 Month 排序
      filteredData.sort((a, b) => {
        if (a.Year !== b.Year) return a.Year - b.Year;
        if (a.Month !== b.Month) return a.Month - b.Month;
        return a.Type.localeCompare(b.Type, 'zh');
      });

      // 导出前去除 Month 字段
      const finalData = filteredData.map(({ Month, ...row }) => row);

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(finalData);
      XLSX.utils.book_append_sheet(wb, ws, "合并结果");

      const colWidths = wantedColumns.map(col => {
        let maxLen = col.length;
        finalData.forEach(row => {
          const cellVal = row[col];
          const len = cellVal ? String(cellVal).length : 0;
          if (len > maxLen) maxLen = len;
        });
        return { wch: maxLen + 4 };
      });
      ws['!cols'] = colWidths;

      ws['!ref'] = XLSX.utils.encode_range({
        s: { c: 0, r: 0 },
        e: { c: wantedColumns.length - 1, r: finalData.length }
      });

      let fileName = document.getElementById("fileNameInput").value.trim();
      if (!fileName) {
        fileName = "合并总表";
      }
      if (!fileName.toLowerCase().endsWith(".xlsx")) {
        fileName += ".xlsx";
      }

      XLSX.writeFile(wb, fileName);

      document.getElementById("status").innerHTML += `
        <br><strong>合并完成！文件已下载，文件名为：${fileName}</strong>
        <br><span class="tip">提示：在 Excel 中选中表格 → 点击 “格式为表格” → 套用样式即可显示颜色</span>
      `;
    }
  </script>

</body>
</html>
