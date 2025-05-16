import express from 'express';
import multer from 'multer';
import cors from 'cors';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';
import { uploadPDF, generateExcel } from './pcrver.js';

// 模擬 __dirname（因為是 ES 模組）
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const app = express();
const port = 3000;
const upload = multer({ dest: 'uploads/' });

app.use(cors());
app.use(express.json());

// 提供 Excel 下載的靜態資料夾
app.use('/generated', express.static(path.join(__dirname, 'generated')));

app.post('/api/analyze', upload.fields([
  { name: 'pcrFile', maxCount: 1 },
  { name: 'productFile', maxCount: 1 }
]), async (req, res) => {
  try {
    const pcrPath = req.files?.pcrFile?.[0]?.path;
    const productPath = req.files?.productFile?.[0]?.path;

    if (!pcrPath || !productPath) {
      throw new Error('缺少 pcr 或產品 PDF 檔案');
    }

    const result = await uploadPDF(pcrPath, productPath);
    res.json({ success: true, data: result });
  } catch (error) {
    console.error('分析錯誤：', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post('/api/export-excel', async (req, res) => {
  try {
    const data = req.body;
    const filePath = path.join(__dirname, 'generated', 'Carbon_Footprint_Report.xlsx');

    if (!Array.isArray(data) || data.length === 0) {
      throw new Error('傳入的資料無效');
    }

    await generateExcel(data, filePath);

    res.json({ success: true, filePath: '/generated/Carbon_Footprint_Report.xlsx' });
  } catch (error) {
    console.error('匯出錯誤：', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
