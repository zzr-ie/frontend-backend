const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// 中间件
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// 配置 multer 用于文件上传
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    if (!fs.existsSync('uploads')) {
      fs.mkdirSync('uploads');
    }
    cb(null, 'uploads/');
  },
  filename: function (req, file, cb) {
    cb(null, Date.now() + '-' + file.originalname);
  }
});

const upload = multer({ 
  storage: storage,
  fileFilter: function (req, file, cb) {
    const allowedMimeTypes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
      'application/vnd.ms-excel', // .xls
      'text/csv', // .csv
      'application/csv', // .csv
      'text/plain' // .csv
    ];
    const allowedExtensions = ['.xlsx', '.xls', '.csv'];
    const fileExt = path.extname(file.originalname).toLowerCase();
    
    if (allowedMimeTypes.includes(file.mimetype) || allowedExtensions.includes(fileExt)) {
      cb(null, true);
    } else {
      cb(new Error('只支持 Excel 文件 和 CSV 文件 (.csv)'));
    }
  },
  limits: {
    fileSize: 50 * 1024 * 1024 // 50MB 限制
  }
});

// 确保上传目录存在
if (!fs.existsSync('uploads')) {
  fs.mkdirSync('uploads');
}

// 内存中存储数据（生产环境应使用数据库）
let uploadedData = {
  sales: [],
  products: [],
  regions: [],
  balance: [],
  disc: [],
  pads: [],
  moto: [],
  fluid: []
};

// 解析 Excel 和 CSV 文件
function parseExcel(filePath) {
  try {
    const workbook = xlsx.readFile(filePath, {
      type: 'binary',
      cellDates: true,
      cellNF: false,
      cellText: false,
      dateNF: 'yyyy-mm-dd',
      raw: true,
      codepage: 65001
    });

    const result = {};

    // 数值解析辅助: 处理逗号、括号负数、尾部负号格式
    function parseNumber(val) {
      if (val === null || val === undefined) return null;
      if (typeof val === 'number') return val;
      if (typeof val === 'string') {
        let cleaned = val.trim();
        if (cleaned === '') return null;
        // 括号表示负数 (1,234.56) -> -1234.56
        cleaned = cleaned.replace(/\(([^)]+)\)/, '-$1');
        // 去掉千位逗号
        cleaned = cleaned.replace(/,/g, '');
        // 处理形如 1234- 结尾负号
        if (/^-?\d+(\.\d+)?-$/.test(cleaned)) {
          cleaned = '-' + cleaned.replace(/-$/, '');
        }
        if (!isNaN(cleaned)) {
          const num = Number(cleaned);
          return isNaN(num) ? null : num;
        }
      }
      return null;
    }

    workbook.SheetNames.forEach(sheetName => {
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = xlsx.utils.sheet_to_json(worksheet, {
        raw: false,
        dateNF: 'yyyy-mm-dd',
        defval: null,
        blankrows: false
      });

      if (jsonData.length > 0) {
        const processedData = jsonData.map(row => {
          const newRow = {};
          Object.entries(row).forEach(([key, value]) => {
            const maybeNumber = parseNumber(value);
            if (maybeNumber !== null) {
              newRow[key] = maybeNumber;
            } else {
              newRow[key] = value === undefined ? null : value;
            }
          });
          return newRow;
        });
        result[sheetName] = processedData;
      }
    });

    if (Object.keys(result).length === 0) {
      throw new Error('文件中没有找到有效数据');
    }
    return result;
  } catch (error) {
    throw new Error('解析文件失败: ' + error.message);
  }
}
// API 路由
// 根据 sheet 名推断年份月份 (YYMM -> 20YY-MM)
function deriveYearMonth(sheetName) {
  if (!sheetName) return null;
  const trimmed = sheetName.trim();
  // 只取前4位数字模式
  const match = trimmed.match(/^(\d{4})/);
  if (match) {
    const raw = match[1];
    const yy = raw.slice(0,2);
    const mm = raw.slice(2,4);
    if (Number(mm) >= 1 && Number(mm) <= 12) {
      return `20${yy}-${mm}`; // 例如 2509 -> 2025-09
    }
  }
  return null;
}
// 上传 Disc 文件
app.post('/api/upload/disc', upload.array('files', 5), (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: '没有上传文件' });
    }
    console.log(`\n========== Disc 文件上传开始 ==========`);
    console.log(`收到 ${req.files.length} 个文件`);
    let newRecords = 0;
    req.files.forEach(file => {
      console.log(`处理文件: ${file.originalname}`);
      const parsed = parseExcel(file.path);
      console.log(`解析到 ${Object.keys(parsed).length} 个工作表:`, Object.keys(parsed));
      Object.entries(parsed).forEach(([sheetName, sheetRows]) => {
        console.log(`  工作表 "${sheetName}": ${sheetRows.length} 条记录`);
        if (sheetRows.length > 0) {
          console.log(`  第一条数据字段:`, Object.keys(sheetRows[0]));
        }
        const ym = deriveYearMonth(sheetName);
        console.log(`  推导的年月: ${ym}`);
        const processed = sheetRows.map(r => {
          if (!r.year_month && ym) {
            return { ...r, year_month: ym };
          }
          return r;
        });
        uploadedData.disc = uploadedData.disc.concat(processed);
        newRecords += processed.length;
      });
    });
    console.log(`Disc 新增记录数: ${newRecords}`);
    console.log(`Disc 累计总记录数: ${uploadedData.disc.length}`);
    console.log(`========== Disc 文件上传完成 ==========\n`);
    res.json({ newRecords, totalRecords: uploadedData.disc.length });
  } catch (error) {
    console.error('Disc 上传错误:', error);
    res.status(500).json({ error: error.message });
  }
});

// 上传 Pads 文件
app.post('/api/upload/pads', upload.array('files', 5), (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: '没有上传文件' });
    }
    console.log(`\n========== Pads 文件上传开始 ==========`);
    console.log(`收到 ${req.files.length} 个文件`);
    let newRecords = 0;
    req.files.forEach(file => {
      console.log(`处理文件: ${file.originalname}`);
      const parsed = parseExcel(file.path);
      console.log(`解析到 ${Object.keys(parsed).length} 个工作表:`, Object.keys(parsed));
      Object.entries(parsed).forEach(([sheetName, sheetRows]) => {
        console.log(`  工作表 "${sheetName}": ${sheetRows.length} 条记录`);
        if (sheetRows.length > 0) {
          console.log(`  第一条数据字段:`, Object.keys(sheetRows[0]));
        }
        const ym = deriveYearMonth(sheetName);
        console.log(`  推导的年月: ${ym}`);
        const processed = sheetRows.map(r => {
          if (!r.year_month && ym) {
            return { ...r, year_month: ym };
          }
          return r;
        });
        uploadedData.pads = uploadedData.pads.concat(processed);
        newRecords += processed.length;
      });
    });
    console.log(`Pads 新增记录数: ${newRecords}`);
    console.log(`Pads 累计总记录数: ${uploadedData.pads.length}`);
    console.log(`========== Pads 文件上传完成 ==========\n`);
    res.json({ newRecords, totalRecords: uploadedData.pads.length });
  } catch (error) {
    console.error('Pads 上传错误:', error);
    res.status(500).json({ error: error.message });
  }
});

// 获取 Disc 数据(分页)
app.get('/api/data/disc', (req, res) => {
  try {
    const total = uploadedData.disc?.length || 0;
    console.log(`获取 Disc 数据，请求分页: total=${total}`);
    if (total === 0) return res.json({ totalRecords: 0, data: [] });
    const limit = Math.min(Number(req.query.limit) || 5000, 20000);
    const offset = Math.max(Number(req.query.offset) || 0, 0);
    if (offset === 0) {
      // 输出样例
      console.log('Disc 数据示例（前3条）:');
      uploadedData.disc.slice(0, 3).forEach((item, idx) => {
        console.log(`  [${idx}] year_month: "${item.year_month}", On-hand: ${item['On-hand']}, Inventory value: ${item['Inventory value']}`);
      });
    }
    const slice = uploadedData.disc.slice(offset, offset + limit);
    res.json({ totalRecords: total, offset, limit: slice.length, data: slice });
  } catch (e) {
    console.error('Disc 数据获取错误:', e);
    res.status(500).json({ error: e.message });
  }
});

// 获取 Pads 数据
app.get('/api/data/pads', (req, res) => {
  console.log(`获取 Pads 数据，当前记录数: ${uploadedData.pads?.length || 0}`);
  res.json(uploadedData.pads || []);
});

// 获取 Moto 数据
app.get('/api/data/moto', (req, res) => {
  console.log(`获取 Moto 数据，当前记录数: ${uploadedData.moto?.length || 0}`);
  res.json(uploadedData.moto || []);
});

// 获取 Fluid 数据
app.get('/api/data/fluid', (req, res) => {
  console.log(`获取 Fluid 数据，当前记录数: ${uploadedData.fluid?.length || 0}`);
  res.json(uploadedData.fluid || []);
});

// 获取 Disc 所有月份(不返回数据,只返回月份列表)
app.get('/api/data/disc/months', (req, res) => {
  try {
    const months = new Set();
    if (uploadedData.disc && uploadedData.disc.length > 0) {
      uploadedData.disc.forEach(item => {
        if (item.year_month) months.add(item.year_month);
      });
    }
    const sortedMonths = Array.from(months).sort();
    console.log(`Disc 月份数: ${sortedMonths.length}, 列表:`, sortedMonths);
    res.json(sortedMonths);
  } catch (e) {
    console.error('获取 Disc 月份错误:', e);
    res.status(500).json({ error: e.message });
  }
});

// 获取 Pads 所有月份
app.get('/api/data/pads/months', (req, res) => {
  try {
    const months = new Set();
    if (uploadedData.pads && uploadedData.pads.length > 0) {
      uploadedData.pads.forEach(item => {
        if (item.year_month) months.add(item.year_month);
      });
    }
    const sortedMonths = Array.from(months).sort();
    console.log(`Pads 月份数: ${sortedMonths.length}`);
    res.json(sortedMonths);
  } catch (e) {
    console.error('获取 Pads 月份错误:', e);
    res.status(500).json({ error: e.message });
  }
});

// 获取 Moto 所有月份
app.get('/api/data/moto/months', (req, res) => {
  try {
    const months = new Set();
    if (uploadedData.moto && uploadedData.moto.length > 0) {
      uploadedData.moto.forEach(item => {
        if (item.year_month) months.add(item.year_month);
      });
    }
    const sortedMonths = Array.from(months).sort();
    console.log(`Moto 月份数: ${sortedMonths.length}`);
    res.json(sortedMonths);
  } catch (e) {
    console.error('获取 Moto 月份错误:', e);
    res.status(500).json({ error: e.message });
  }
});

// 获取 Fluid 所有月份
app.get('/api/data/fluid/months', (req, res) => {
  try {
    const months = new Set();
    if (uploadedData.fluid && uploadedData.fluid.length > 0) {
      uploadedData.fluid.forEach(item => {
        if (item.year_month) months.add(item.year_month);
      });
    }
    const sortedMonths = Array.from(months).sort();
    console.log(`Fluid 月份数: ${sortedMonths.length}`);
    res.json(sortedMonths);
  } catch (e) {
    console.error('获取 Fluid 月份错误:', e);
    res.status(500).json({ error: e.message });
  }
});

// 计算滚动12个月 DIO
function computeRollingDIO(inventoryByMonth, salesByMonth) {
  const results = [];
  const months = Object.keys(inventoryByMonth).sort();
  months.forEach(ym => {
    try {
      const [yearStr, monthStr] = ym.split('-');
      const year = Number(yearStr);
      const month = Number(monthStr);
      const baseDate = new Date(year, month - 1, 1);
      const prev12 = [];
      for (let i = 0; i < 12; i++) {
        const d = new Date(baseDate);
        d.setMonth(d.getMonth() - i);
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        prev12.push(`${y}-${m}`);
      }
      const monthlySales = prev12.map(m => salesByMonth[m] || 0);
      const totalSales = monthlySales.reduce((a,b)=>a+b,0);
      const inventoryValue = typeof inventoryByMonth[ym] === 'object' && inventoryByMonth[ym] !== null
        ? Number(inventoryByMonth[ym].value) || 0
        : Number(inventoryByMonth[ym]) || 0;
      if (totalSales > 0) {
        const dio = Math.round((inventoryValue / totalSales) * 360 * 100) / 100;
        results.push({ year_month: ym, inventory_value: inventoryValue, sales_12m: totalSales, dio, net_sales_month: (salesByMonth[ym] || 0), sales_filter_months: monthlySales });
      } else {
        results.push({ year_month: ym, inventory_value: inventoryValue, sales_12m: null, dio: null, net_sales_month: (salesByMonth[ym] || 0), sales_filter_months: monthlySales });
      }
    } catch (e) {
      const inventoryValue = typeof inventoryByMonth[ym] === 'object' && inventoryByMonth[ym] !== null
        ? Number(inventoryByMonth[ym].value) || 0
        : Number(inventoryByMonth[ym]) || 0;
      results.push({ year_month: ym, inventory_value: inventoryValue, sales_12m: null, dio: null, net_sales_month: (salesByMonth[ym] || 0), sales_filter_months: [] });
    }
  });
  return results;
}

// 计算 Disc DIO（Days Inventory Outstanding）
app.get('/api/data/disc/dio', (req, res) => {
  try {
    console.log('\n========== DIO 计算开始 ==========');
    console.log('Disc 数据条数:', uploadedData.disc?.length || 0);
    console.log('Sales 数据条数:', uploadedData.sales?.length || 0);
    
    if (!uploadedData.disc || uploadedData.disc.length === 0) {
      console.log('没有 Disc 数据，返回空数组');
      return res.json([]);
    }
    
    // 按月份汇总 Disc 库存
    const discByMonth = {};
    // 使用健壮的数值解析
    function toNumber(v){
      if (v === null || v === undefined) return 0;
      if (typeof v === 'number') return v;
      if (typeof v === 'string') {
        let cleaned = v.trim();
        if (cleaned === '') return 0;
        cleaned = cleaned.replace(/\(([^)]+)\)/, '-$1').replace(/,/g,'');
        if (/^-?\d+(\.\d+)?-$/.test(cleaned)) cleaned = '-' + cleaned.replace(/-$/,'');
        const n = Number(cleaned);
        return isNaN(n) ? 0 : n;
      }
      return 0;
    }

    uploadedData.disc.forEach(item => {
      const ym = item.year_month;
      if (!ym) return;
      if (!discByMonth[ym]) {
        discByMonth[ym] = { value: 0, qty: 0, rows: 0 };
      }
      discByMonth[ym].value += toNumber(item['Inventory value']);
      discByMonth[ym].qty += toNumber(item['On-hand']);
      discByMonth[ym].rows += 1;
    });
    console.log('Disc 月份汇总(过滤后):', Object.keys(discByMonth).sort());
    Object.keys(discByMonth).sort().slice(0,12).forEach(m => {
      const info = discByMonth[m];
      console.log(`  月份 ${m}: 参与行数=${info.rows}, 累计金额=${info.value.toFixed(2)}`);
    });
    
      // 过滤 Disc Cost Center=34N00001 且 Item Group 包含 2300 或 2330
      const discSalesByMonth = {};
      const allSalesByMonth = {};
      if (uploadedData.sales && uploadedData.sales.length > 0) {
        uploadedData.sales.forEach(row => {
          const costCenter = (row['Cost Center'] || row['cost center'] || row['Cost center'] || '').toString().trim();
          const itemGroup = (row['Item - Item Group Full Name'] || row['Item group'] || row['Item Group'] || '').toString();
          const year = row.Year || row.year;
          const month = row.Month || row.month;
          if (!year || !month) return;
          const ym = `${year}-${String(month).padStart(2, '0')}`;
          if (!allSalesByMonth[ym]) allSalesByMonth[ym] = 0;
          allSalesByMonth[ym] += Number(row['Net Sales']) || 0;
          const isDiscGroup = itemGroup.includes('2300') || itemGroup.includes('2330');
          if (costCenter === '34N00001' && isDiscGroup) {
            if (!discSalesByMonth[ym]) discSalesByMonth[ym] = 0;
            discSalesByMonth[ym] += Number(row['Net Sales']) || 0;
          }
        });
        console.log('筛选后的 Disc 销售月份:', Object.keys(discSalesByMonth).length);
        if (Object.keys(discSalesByMonth).length === 0) {
          console.warn('Disc 销售筛选结果为空，使用回退逻辑(全部销售)计算 DIO');
          uploadedData.sales.forEach(row => {
            const year = row.Year || row.year;
            const month = row.Month || row.month;
            if (!year || !month) return;
            const ym = `${year}-${String(month).padStart(2, '0')}`;
            if (!discSalesByMonth[ym]) discSalesByMonth[ym] = 0;
            const netSales = Number(row['Net Sales']) || 0;
            discSalesByMonth[ym] += netSales;
          });
          console.log('回退后 Disc 销售月份:', Object.keys(discSalesByMonth).length);
        }
       
        if (Object.keys(discSalesByMonth).length === 0 && uploadedData.disc) {
          console.warn('Disc 回退全部销售仍为空，尝试使用库存行 sales_12m 字段');
          uploadedData.disc.forEach(item => {
            const ym = item.year_month;
            if (!ym) return;
            const sale12 = Number(item['sales_12m']) || 0;
            if (!discSalesByMonth[ym]) discSalesByMonth[ym] = 0;
            discSalesByMonth[ym] += sale12;
          });
          console.log('使用库存 sales_12m 后 Disc 销售月份:', Object.keys(discSalesByMonth).length);
        }
      } else {
        console.log('没有 Sales 数据 (用于 Disc DIO 新逻辑)');
      }
  
      // 计算每月 DIO：过去12个月净销售额求和 
      const dioResults = [];
      Object.keys(discByMonth).sort().forEach(ym => {
        try {
          const [year, month] = ym.split('-').map(Number);
          const baseDate = new Date(year, month - 1, 1);
          const prev12 = [];
          // 当前月份 + 前11个月
          for (let i = 0; i < 12; i++) {
            const d = new Date(baseDate);
            d.setMonth(d.getMonth() - i);
            const y = d.getFullYear();
            const m = String(d.getMonth() + 1).padStart(2, '0');
            prev12.push(`${y}-${m}`);
          }
          const monthlySales = prev12.map(m => (discSalesByMonth[m] !== undefined ? discSalesByMonth[m] : (allSalesByMonth[m] || 0)));
          const totalSales12m = monthlySales.reduce((a,b)=>a+b,0);
          const inventoryValue = discByMonth[ym].value;
          let dio = null;
          if (totalSales12m > 0) {
            dio = Math.round((inventoryValue / totalSales12m) * 360 * 100) / 100;
          }
          dioResults.push({
            year_month: ym,
            inventory_value: inventoryValue,
            sales_12m: totalSales12m,
            dio: dio,
            sales_filter_months: monthlySales,
            net_sales_month: (discSalesByMonth[ym] !== undefined ? discSalesByMonth[ym] : (allSalesByMonth[ym] || 0))
          });
        } catch (e) {
          console.error(`Disc DIO 计算失败 (${ym}):`, e);
          dioResults.push({
            year_month: ym,
            inventory_value: discByMonth[ym].value,
            sales_12m: null,
            dio: null
          });
        }
      });
    
    console.log(`DIO 计算完成，返回 ${dioResults.length} 条记录`);
    console.log('DIO 示例（前3条）:', dioResults.slice(0, 3));
    console.log('========== DIO 计算结束 ==========\n');
    
    res.json(dioResults);
  } catch (error) {
    console.error('DIO 计算错误:', error);
    res.status(500).json({ error: error.message });
  }
});

// 计算 Pads DIO
app.get('/api/data/pads/dio', (req, res) => {
  try {
    console.log('\n========== Pads DIO 计算开始 ==========');
    console.log('Pads 数据条数:', uploadedData.pads?.length || 0);
    console.log('Sales 数据条数:', uploadedData.sales?.length || 0);
    
    if (!uploadedData.pads || uploadedData.pads.length === 0) {
      console.log('没有 Pads 数据，返回空数组');
      return res.json([]);
    }
    
    const padsByMonth = {};
    uploadedData.pads.forEach(item => {
      const ym = item.year_month;
      if (!ym) return;
      if (!padsByMonth[ym]) {
        padsByMonth[ym] = { value: 0, qty: 0, rows: 0 };
      }
      padsByMonth[ym].value += Number(item['Inventory value']) || 0;
      padsByMonth[ym].qty += Number(item['On-hand']) || 0;
      padsByMonth[ym].rows += 1;
    });
    console.log('Pads 月份汇总(含行数):', Object.keys(padsByMonth).sort());
    Object.keys(padsByMonth).sort().slice(0,12).forEach(m => {
      const info = padsByMonth[m];
      console.log(`  月份 ${m}: 行数=${info.rows}, 金额=${info.value.toFixed(2)}`);
    });
    
    // 筛选 Pads Cost Center=34N00001 且 Item Group为2000
    const padsSalesByMonth = {};
    const allSalesByMonthPads = {}; 
    if (uploadedData.sales && uploadedData.sales.length > 0) {
      uploadedData.sales.forEach(row => {
        const costCenter = (row['Cost Center'] || row['cost center'] || row['Cost center'] || '').toString().trim();
        const itemGroup = (row['Item - Item Group Full Name'] || row['Item group'] || row['Item Group'] || '').toString();
        const year = row.Year || row.year;
        const month = row.Month || row.month;
        if (!year || !month) return;
        const ym = `${year}-${String(month).padStart(2, '0')}`;
        if (!allSalesByMonthPads[ym]) allSalesByMonthPads[ym] = 0;
        allSalesByMonthPads[ym] += Number(row['Net Sales']) || 0;
        const isPadsGroup = itemGroup.includes('2000');
        if (costCenter === '34N00001' && isPadsGroup) {
          if (!padsSalesByMonth[ym]) padsSalesByMonth[ym] = 0;
          padsSalesByMonth[ym] += Number(row['Net Sales']) || 0;
        }
      });
      console.log('筛选后的 Pads 销售月份数:', Object.keys(padsSalesByMonth).length);
      if (Object.keys(padsSalesByMonth).length === 0) {
        console.warn('Pads 销售筛选结果为空，使用回退逻辑(全部销售)');
        uploadedData.sales.forEach(row => {
          const year = row.Year || row.year;
          const month = row.Month || row.month;
          if (!year || !month) return;
          const ym = `${year}-${String(month).padStart(2, '0')}`;
          if (!padsSalesByMonth[ym]) padsSalesByMonth[ym] = 0;
          const netSales = Number(row['Net Sales']) || 0;
          padsSalesByMonth[ym] += netSales;
        });
        console.log('回退后 Pads 销售月份数:', Object.keys(padsSalesByMonth).length);
      }
      if (Object.keys(padsSalesByMonth).length === 0 && uploadedData.pads) {
        console.warn('Pads 回退全部销售仍为空，尝试使用库存行 Sales Volume 1Y / sales_12m 字段');
        uploadedData.pads.forEach(item => {
          const ym = item.year_month;
          if (!ym) return;
          const vol1Y = Number(item['Sales Volume 1Y']) || Number(item['sales_12m']) || 0;
          if (!padsSalesByMonth[ym]) padsSalesByMonth[ym] = 0;
          padsSalesByMonth[ym] += vol1Y;
        });
        console.log('使用库存替代后 Pads 销售月份数:', Object.keys(padsSalesByMonth).length);
      }
    } else {
      console.log('没有 Sales 数据 (用于 Pads DIO 新逻辑)');
    }
    
    // 计算 Pads 每月 DIO
    const dioResults = [];
    Object.keys(padsByMonth).sort().forEach(ym => {
      try {
        const [year, month] = ym.split('-').map(Number);
        const baseDate = new Date(year, month - 1, 1);
        const prev12 = [];
        for (let i = 0; i < 12; i++) {
          const d = new Date(baseDate);
          d.setMonth(d.getMonth() - i);
          const y = d.getFullYear();
          const m = String(d.getMonth() + 1).padStart(2, '0');
          prev12.push(`${y}-${m}`);
        }
        const monthlySales = prev12.map(m => (padsSalesByMonth[m] !== undefined ? padsSalesByMonth[m] : (allSalesByMonthPads[m] || 0)));
        const totalSales12m = monthlySales.reduce((a,b)=>a+b,0);
        const inventoryValue = padsByMonth[ym].value;
        let dio = null;
        if (totalSales12m > 0) {
          dio = Math.round((inventoryValue / totalSales12m) * 360* 100) / 100;
        }
        dioResults.push({
          year_month: ym,
          inventory_value: inventoryValue,
          sales_12m: totalSales12m,
          dio: dio,
          sales_filter_months: monthlySales,
          net_sales_month: (padsSalesByMonth[ym] !== undefined ? padsSalesByMonth[ym] : (allSalesByMonthPads[ym] || 0))
        });
      } catch (e) {
        console.error(`Pads DIO 计算失败 (${ym}):`, e);
        dioResults.push({
          year_month: ym,
          inventory_value: padsByMonth[ym].value,
          sales_12m: null,
          dio: null
        });
      }
    });
    
    console.log(`Pads DIO 计算完成，返回 ${dioResults.length} 条记录`);
    res.json(dioResults);
  } catch (error) {
    console.error('Pads DIO 计算错误:', error);
    res.status(500).json({ error: error.message });
  }
});

// 计算 Moto DIO
app.get('/api/data/moto/dio', (req, res) => {
  try {
    console.log('\n========== Moto DIO 计算开始 ==========');
    console.log('Moto 数据条数:', uploadedData.moto?.length || 0);
    console.log('Sales 数据条数:', uploadedData.sales?.length || 0);
    
    if (!uploadedData.moto || uploadedData.moto.length === 0) {
      console.log('没有 Moto 数据，返回空数组');
      return res.json([]);
    }
    
    const motoByMonth = {};
    uploadedData.moto.forEach(item => {
      const ym = item.year_month;
      if (!ym) return;
      if (!motoByMonth[ym]) {
        motoByMonth[ym] = { value: 0, qty: 0 };
      }
      motoByMonth[ym].value += Number(item['Inventory value']) || 0;
      motoByMonth[ym].qty += Number(item['On-hand']) || 0;
    });
    console.log('Moto 月份汇总:', Object.keys(motoByMonth).sort());
    
    // Cost Center=34N00037 或 34N00039
    const filteredSalesByMonth = {};
    const allSalesByMonthMoto = {};
    if (uploadedData.sales && uploadedData.sales.length > 0) {
      uploadedData.sales.forEach(row => {
        const year = row.Year || row.year;
        const month = row.Month || row.month;
        if (!year || !month) return;
        const ym = `${year}-${String(month).padStart(2, '0')}`;
        if (!allSalesByMonthMoto[ym]) allSalesByMonthMoto[ym] = 0;
        const net = Number(row['Net Sales']) || 0;
        allSalesByMonthMoto[ym] += net;
        const costCenter = (row['Cost Center'] || row['cost center'] || row['Cost center'] || '').toString().trim();
        if (costCenter === '34N00037' || costCenter === '34N00039') {
          if (!filteredSalesByMonth[ym]) filteredSalesByMonth[ym] = 0;
          filteredSalesByMonth[ym] += net;
        }
      });
    }
    console.log('Moto 筛选销售月份数:', Object.keys(filteredSalesByMonth).length);
    const combinedSalesByMonth = {};
    Object.keys(allSalesByMonthMoto).forEach(m => {
      combinedSalesByMonth[m] = (filteredSalesByMonth[m] !== undefined ? filteredSalesByMonth[m] : allSalesByMonthMoto[m]);
    });
    const dioResults = computeRollingDIO(motoByMonth, combinedSalesByMonth);
    
    console.log(`Moto DIO 计算完成，返回 ${dioResults.length} 条记录`);
    res.json(dioResults);
  } catch (error) {
    console.error('Moto DIO 计算错误:', error);
    res.status(500).json({ error: error.message });
  }
});

// 计算 Fluid DIO
app.get('/api/data/fluid/dio', (req, res) => {
  try {
    console.log('\n========== Fluid DIO 计算开始 ==========');
    console.log('Fluid 数据条数:', uploadedData.fluid?.length || 0);
    console.log('Sales 数据条数:', uploadedData.sales?.length || 0);
    
    if (!uploadedData.fluid || uploadedData.fluid.length === 0) {
      console.log('没有 Fluid 数据，返回空数组');
      return res.json([]);
    }
    
    const fluidByMonth = {};
    uploadedData.fluid.forEach(item => {
      const ym = item.year_month;
      if (!ym) return;
      if (!fluidByMonth[ym]) {
        fluidByMonth[ym] = { value: 0, qty: 0 };
      }
      fluidByMonth[ym].value += Number(item['Inventory value']) || 0;
      fluidByMonth[ym].qty += Number(item['On-hand']) || 0;
    });
    console.log('Fluid 月份汇总:', Object.keys(fluidByMonth).sort());
    
    // Fluid：Cost Center=34N00001 且 Item number 以 L 开头
    const filteredSalesByMonth = {};
    const allSalesByMonthFluid = {};
    if (uploadedData.sales && uploadedData.sales.length > 0) {
      uploadedData.sales.forEach(row => {
        const year = row.Year || row.year;
        const month = row.Month || row.month;
        if (!year || !month) return;
        const ym = `${year}-${String(month).padStart(2, '0')}`;
        if (!allSalesByMonthFluid[ym]) allSalesByMonthFluid[ym] = 0;
        const net = Number(row['Net Sales']) || 0;
        allSalesByMonthFluid[ym] += net;
        const costCenter = (row['Cost Center'] || row['cost center'] || row['Cost center'] || '').toString().trim();
        const itemNumber = (row['Item number'] || row['Item Number'] || '').toString().trim();
        if (costCenter === '34N00001' && itemNumber.startsWith('L')) {
          if (!filteredSalesByMonth[ym]) filteredSalesByMonth[ym] = 0;
          filteredSalesByMonth[ym] += net;
        }
      });
    }
    console.log('Fluid 筛选销售月份数:', Object.keys(filteredSalesByMonth).length);
    const combinedSalesByMonth = {};
    Object.keys(allSalesByMonthFluid).forEach(m => {
      combinedSalesByMonth[m] = (filteredSalesByMonth[m] !== undefined ? filteredSalesByMonth[m] : allSalesByMonthFluid[m]);
    });
    const dioResults = computeRollingDIO(fluidByMonth, combinedSalesByMonth);
    
    console.log(`Fluid DIO 计算完成，返回 ${dioResults.length} 条记录`);
    res.json(dioResults);
  } catch (error) {
    console.error('Fluid DIO 计算错误:', error);
    res.status(500).json({ error: error.message });
  }
});

// 健康检查端点,快速确认服务是否运行
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', uptime: process.uptime(), records: uploadedData.sales.length });
});

// 上传 Excel 文件 - 支持多文件上传
app.post('/api/upload', upload.array('excelFiles', 10), (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: '没有上传文件' });
    }

    let totalNewRecords = 0;
    const uploadedSheets = [];

    req.files.forEach(file => {
      try {
        const filePath = file.path;
        const parsedData = parseExcel(filePath);

        // 获取第一个工作表的数据
        const firstSheetName = Object.keys(parsedData)[0];
        if (firstSheetName) {
          const newData = parsedData[firstSheetName];
          
          if (!uploadedData.sales) {
            uploadedData.sales = [];
          }
          uploadedData.sales = uploadedData.sales.concat(newData);
          
          totalNewRecords += newData.length;
          uploadedSheets.push({
            fileName: file.originalname,
            records: newData.length
          });
        }

        fs.unlinkSync(filePath);
      } catch (error) {
        console.error(`处理文件 ${file.originalname} 失败:`, error.message);
      }
    });

    console.log('批量上传完成');
    console.log('上传文件数:', req.files.length);
    console.log('新增总记录数:', totalNewRecords);
    console.log('累计总记录数:', uploadedData.sales?.length || 0);

    res.json({
      message: '文件上传成功',
      filesUploaded: req.files.length,
      uploadedSheets: uploadedSheets,
      newRecords: totalNewRecords,
      totalRecords: uploadedData.sales?.length || 0
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Total Inventory DIO calculation endpoint
app.get('/api/data/total/dio', (req, res) => {
  try {
    console.log('计算 Total Inventory DIO...');
 
    const allInventoryData = [
      ...(uploadedData.disc || []),
      ...(uploadedData.pads || []),
      ...(uploadedData.moto || []),
      ...(uploadedData.fluid || [])
    ];
    
    console.log('Total inventory records:', allInventoryData.length);
    
    if (allInventoryData.length === 0) {
      return res.json([]);
    }
    
    const inventoryByMonth = {};
    allInventoryData.forEach(item => {
      const month = item.year_month;
      const inventoryValue = Number(item['Inventory value']) || 0;
      if (!month) return;
      if (!inventoryByMonth[month]) {
        inventoryByMonth[month] = { value: 0 };
      }
      inventoryByMonth[month].value += inventoryValue;
    });
    
    console.log('Total inventory months:', Object.keys(inventoryByMonth).length);
    
    const salesByMonth = {};
    console.log('Sales data exists:', !!uploadedData.sales);
    console.log('Sales data length:', uploadedData.sales?.length || 0);
    
    if (uploadedData.sales && uploadedData.sales.length > 0) {
      uploadedData.sales.forEach(item => {
        const year = item.Year || item.year;
        const month = item.Month || item.month;
        if (!year || !month) return;
        const yearMonth = `${year}-${String(month).padStart(2, '0')}`;
        if (!salesByMonth[yearMonth]) {
          salesByMonth[yearMonth] = 0;
        }
        salesByMonth[yearMonth] += Number(item['Net Sales']) || 0;
      });
    }
    
    console.log('Sales months:', Object.keys(salesByMonth).length);
    
    const dioResults = computeRollingDIO(inventoryByMonth, salesByMonth);
    
    console.log('Total DIO results:', dioResults.length);
    res.json(dioResults);
    
  } catch (error) {
    console.error('Error calculating Total DIO:', error);
    res.status(500).json({ error: error.message });
  }
});

// 清空所有数据
app.post('/api/clear', (req, res) => {
  try {
    uploadedData = {
      sales: [],
      products: [],
      regions: [],
      balance: [],
      disc: [],
      pads: [],
      moto: [],
      fluid: []
    };
    console.log('已清空所有数据');
    res.json({ message: '数据已清空' });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post('/api/clear/inventory', (req, res) => {
  try {
    uploadedData.balance = [];
    uploadedData.disc = [];
    uploadedData.pads = [];
    uploadedData.moto = [];
    uploadedData.fluid = [];
    console.log('已清空 Inventory 数据');
    res.json({ message: 'Inventory 数据已清空' });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 获取销售数据
app.get('/api/data/sales', (req, res) => {
  try {
    const total = uploadedData.sales?.length || 0;
    if (total === 0) {
      return res.json({ totalRecords: 0, data: [] });
    }
    // 支持分页,避免一次性返回超大数据导致 Invalid string length
    const limit = Math.min(Number(req.query.limit) || 5000, 20000); // 单次最大2万
    const offset = Math.max(Number(req.query.offset) || 0, 0);
    const slice = uploadedData.sales.slice(offset, offset + limit);
    res.json({
      totalRecords: total,
      offset,
      limit: slice.length,
      data: slice
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 获取产品类型数据 - 从sales表中提取
app.get('/api/data/products', (req, res) => {
  try {
    if (uploadedData.sales && uploadedData.sales.length > 0) {
      // 按产品类型分组汇总销售额
      const productSummary = {};
      uploadedData.sales.forEach(item => {
        const productType = item['Product Type'];
        if (productType) {
          if (!productSummary[productType]) {
            productSummary[productType] = 0;
          }
          productSummary[productType] += Number(item['Total Sales Amt']) || 0;
        }
      });
      
      // 转换为数组格式
      const result = Object.entries(productSummary).map(([product, sales]) => ({
        category: product,
        sales: sales
      }));
      
      res.json(result);
    } else {
      // 返回空数组而不是错误
      res.json([]);
    }
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 获取地区数据 - 从sales表中提取
app.get('/api/data/regions', (req, res) => {
  try {
    if (uploadedData.sales && uploadedData.sales.length > 0) {
      // 按地区分组汇总销售额
      const regionSummary = {};
      uploadedData.sales.forEach(item => {
        const region = item.Region || item['Sub Region'] || item.region;
        if (region) {
          if (!regionSummary[region]) {
            regionSummary[region] = 0;
          }
          regionSummary[region] += Number(item['Total Sales Amt']) || 0;
        }
      });
      
      // 转换为数组格式
      const result = Object.entries(regionSummary).map(([region, sales]) => ({
        region: region,
        sales: sales
      }));
      
      res.json(result);
    } else {
      // 返回空数组而不是错误
      res.json([]);
    }
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 获取年份数据 - 从sales表中提取唯一年份
app.get('/api/data/year', (req, res) => {
  try {
    if (uploadedData.sales && uploadedData.sales.length > 0) {
      // 提取唯一年份
      const yearsSet = new Set();
      uploadedData.sales.forEach(item => {
        const year = item.Year || item.year;
        if (year) {
          yearsSet.add(year);
        }
      });
      
      const years = Array.from(yearsSet).sort().map(year => ({ year: year }));
      
      res.json(years);
    } else {
      // 返回空数组而不是错误
      res.json([]);
    }
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 获取统计数据
app.get('/api/stats', (req, res) => {
  try {
    if (uploadedData.sales && uploadedData.sales.length > 0) {
      // 统计Total Sales Amt、Sales Qua和Net Sales字段
      const totalSales = uploadedData.sales.reduce((sum, item) => sum + (Number(item['Total Sales Amt']) || 0), 0);
      const salesQuantity = uploadedData.sales.reduce((sum, item) => sum + (Number(item['Sales Qua']) || 0), 0);
      const netSales = uploadedData.sales.reduce((sum, item) => sum + (Number(item['Net Sales']) || 0), 0);
      
      // 计算USD和RMB金额 (Total Sales Amt * rate)
      const usdAmount = uploadedData.sales.reduce((sum, item) => {
        const totalSalesAmt = Number(item['Total Sales Amt']) || 0;
        const usdRate = Number(item['USD_rate']) || 0;
        return sum + (totalSalesAmt * usdRate);
      }, 0);
      
      const rmbAmount = uploadedData.sales.reduce((sum, item) => {
        const totalSalesAmt = Number(item['Total Sales Amt']) || 0;
        const rmbRate = Number(item['RMB_rate']) || 0;
        return sum + (totalSalesAmt * rmbRate);
      }, 0);
      
      console.log('统计结果:', { totalSales, salesQuantity, netSales, usdAmount, rmbAmount });
      
      res.json({
        totalSales,
        salesQuantity,
        netSales,
        usdAmount,
        rmbAmount
      });
    } else {
      // 返回零值而不是错误
      res.json({
        totalSales: 0,
        salesQuantity: 0,
        netSales: 0,
        usdAmount: 0,
        rmbAmount: 0
      });
    }
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});
// 按年份分组汇总接口
app.get('/api/data/yearly-summary', (req, res) => {
  try {
    if (uploadedData.sales && uploadedData.sales.length > 0) {
      console.log('开始处理汇总,数据条数:', uploadedData.sales.length);
      console.log('第一条数据示例:', uploadedData.sales[0]);
      
      const summary = {};
      uploadedData.sales.forEach(item => {
        const year = item.Year || item.year || item.YEAR;
        if (!year) return;
        
        // 确保每个年份都有独立的对象
        if (!summary[year]) {
          summary[year] = {
            year: Number(year),
            unitPriceTotal: 0,
            unitPriceCount: 0,
            salesQuantity: 0,
            grossSalesAmount: 0,
            ebit: 0
          };
        }
        
        // 累加各项指标
        const salesAmtPerUnit = Number(item['Sales Amt per Unit']) || 0;
        if (salesAmtPerUnit !== 0) {
          summary[year].unitPriceTotal += salesAmtPerUnit;
          summary[year].unitPriceCount += 1;
        }
        
        summary[year].salesQuantity += Number(item['Sales Qua']) || 0;
        summary[year].grossSalesAmount += Number(item['Total Sales Amt']) || 0;
        summary[year].ebit += Number(item['EBIT']) || 0;
      });
      
      const result = Object.values(summary).map(item => ({
        year: item.year,
        unitPrice: item.unitPriceCount > 0 ? (item.unitPriceTotal / item.unitPriceCount) : 0,
        salesQuantity: item.salesQuantity,
        grossSalesAmount: item.grossSalesAmount,
        ebit: item.ebit
      })).sort((a, b) => a.year - b.year);
      
      console.log('汇总结果:', result);
      res.json(result);
    } else {
      // 返回空数组而不是错误
      res.json([]);
    }
  } catch (error) {
    console.error('汇总错误:', error);
    res.status(500).json({ error: error.message });
  }
});

// 按月份分组汇总接口 (需要指定年份)
app.get('/api/data/monthly-summary', (req, res) => {
  try {
    const { year } = req.query;
    
    if (uploadedData.sales && uploadedData.sales.length > 0) {
      console.log('开始处理月度汇总, 年份:', year);
      
      // 过滤指定年份的数据
      let filteredData = uploadedData.sales;
      if (year && year !== 'all') {
        filteredData = uploadedData.sales.filter(item => {
          const itemYear = item.Year || item.year || item.YEAR;
          return itemYear == year;
        });
      }
      
      const summary = {};
      filteredData.forEach(item => {
        // 尝试从Date或Month字段获取月份
        let month = item.Month || item.month;
        
        // 如果没有Month字段,尝试从Date字段提取
        if (!month && item.Date) {
          const date = new Date(item.Date);
          if (!isNaN(date.getTime())) {
            month = date.getMonth() + 1; // JavaScript月份从0开始
          }
        }
        
        if (!month) return;
        
        if (!summary[month]) {
          summary[month] = {
            month: Number(month),
            unitPriceTotal: 0,
            unitPriceCount: 0,
            salesQuantity: 0,
            grossSalesAmount: 0,
            ebit: 0
          };
        }
        
        if (item['Sales Amt per Unit']) {
          summary[month].unitPriceTotal += Number(item['Sales Amt per Unit']) || 0;
          summary[month].unitPriceCount += 1;
        }
        if (item['Sales Qua']) summary[month].salesQuantity += Number(item['Sales Qua']) || 0;
        if (item['Total Sales Amt']) summary[month].grossSalesAmount += Number(item['Total Sales Amt']) || 0;
        if (item['EBIT']) summary[month].ebit += Number(item['EBIT']) || 0;
      });
      
      const result = Object.values(summary).map(item => ({
        month: item.month,
        unitPrice: item.unitPriceCount ? (item.unitPriceTotal / item.unitPriceCount) : 0,
        salesQuantity: item.salesQuantity,
        grossSalesAmount: item.grossSalesAmount,
        ebit: item.ebit
      })).sort((a, b) => a.month - b.month);
      
      console.log('月度汇总结果:', result);
      res.json(result);
    } else {
      // 返回空数组而不是错误
      res.json([]);
    }
  } catch (error) {
    console.error('月度汇总错误:', error);
    res.status(500).json({ error: error.message });
  }
});

// 获取成本分解数据 - Cost Breakdown
app.get('/api/data/cost-breakdown', (req, res) => {
  try {
    if (uploadedData.sales && uploadedData.sales.length > 0) {
      // 计算Net Sales Amt总和作为分母
      const netSalesAmt = uploadedData.sales.reduce((sum, item) => sum + (Number(item['Net Sales']) || 0), 0);
      
      // 定义成本类别及其对应的列名
      const costMapping = {
        'Rebate': 'Rebate Amount',
        'Freight Out': 'Freight Out',
        'Purchasing Material': 'Act Material Cost Total',
        'Warehouse Operation': 'Contracted Work Total',
        'Commission': 'Commission',
        'Other SG&A': 'SG&A Total Exclud. Commission',
        'EBIT': 'EBIT'
      };
      
      const costData = [];
      
      Object.entries(costMapping).forEach(([category, columnName]) => {
        const total = uploadedData.sales.reduce((sum, item) => {
          return sum + (Number(item[columnName]) || 0);
        }, 0);
        
        const percentage = netSalesAmt > 0 ? (total / netSalesAmt) * 100 : 0;
        
        costData.push({
          category: category,
          value: total,
          percentage: percentage
        });
      });
      
      console.log('Cost Breakdown - Net Sales Amt:', netSalesAmt);
      console.log('Cost Breakdown Data:', costData);
      
      res.json(costData);
    } else {
      res.json([]);
    }
  } catch (error) {
    console.error('成本分解错误:', error);
    res.status(500).json({ error: error.message });
  }
});

// 上传 Inventory 文件（智能识别 Balance、Disc、Pads）
app.post('/api/upload/balance', upload.array('files', 10), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: '没有上传文件' });
    }

    let totalNewRecords = 0;
    const errors = [];
    const uploadStats = {
      balance: 0,
      disc: 0,
      pads: 0,
      moto: 0,
      fluid: 0
    };

    for (const file of req.files) {
      try {
        console.log('\n========== 处理库存文件 ==========');
        console.log('文件名:', file.originalname);
        const parsedData = parseExcel(file.path);
        
        // 处理所有 sheet 的数据
        Object.entries(parsedData).forEach(([sheetName, newData]) => {
          if (newData && newData.length > 0) {
            console.log(`\n工作表 "${sheetName}": ${newData.length} 条记录`);
            
            // 检查第一行数据的字段，智能识别文件类型
            const firstRow = newData[0];
            const fields = Object.keys(firstRow);
            console.log('字段列表:', fields.slice(0, 10).join(', '), fields.length > 10 ? '...' : '');
            
            let fileType = 'balance'; // 默认为 balance
            
            // 识别 Disc/Pads/Moto/Fluid 文件：包含 "Item number" 和 "On-hand" 字段，且包含 "Inventory value"
            if (fields.includes('Item number') && fields.includes('On-hand') && fields.includes('Inventory value')) {
              // 通过文件名区分不同类型
              const fileName = file.originalname.toLowerCase();
              const categoryFromRow = (firstRow.category || firstRow.Category || '').toString().toLowerCase();

              // 优先文件名判断
              if (fileName.includes('disc') || fileName.includes('盘')) {
                fileType = 'disc';
              } else if (fileName.includes('pad') || fileName.includes('片')) {
                fileType = 'pads';
              } else if (fileName.includes('moto') || fileName.includes('摩托')) {
                fileType = 'moto';
              } else if (fileName.includes('fluid') || fileName.includes('液')) {
                fileType = 'fluid';
              // 次级依据：category 字段
              } else if (categoryFromRow.includes('moto')) {
                fileType = 'moto';
              } else if (categoryFromRow.includes('fluid') || categoryFromRow.includes('oil')) {
                fileType = 'fluid';
              } else if (categoryFromRow.includes('pad')) {
                fileType = 'pads';
              } else if (categoryFromRow.includes('disc')) {
                fileType = 'disc';
              } else {
                fileType = 'disc'; // 默认归入 disc（兼容旧逻辑）
              }
            }
            // 识别 Balance 文件：包含 "1400-INVENTORY" 等特定字段
            else if (fields.some(f => f.includes('1400-INVENTORY') || f.includes('102011101'))) {
              fileType = 'balance';
            }
            
            console.log(`识别为: ${fileType.toUpperCase()} 文件`);
            
            const ym = deriveYearMonth(sheetName);
            console.log(`推导年月: ${ym}`);
            
            const processed = newData.map(r => {
              if ((fileType === 'disc' || fileType === 'pads' || fileType === 'moto' || fileType === 'fluid') && ym) {
                return { ...r, year_month: ym };
              } else if (!r.year_month && ym) {
                return { ...r, year_month: ym };
              }
              return r;
            });

            if (fileType === 'disc') {
              uploadedData.disc = uploadedData.disc.concat(processed);
              uploadStats.disc += processed.length;
            } else if (fileType === 'pads') {
              uploadedData.pads = uploadedData.pads.concat(processed);
              uploadStats.pads += processed.length;
            } else if (fileType === 'moto') {
              uploadedData.moto = uploadedData.moto.concat(processed);
              uploadStats.moto += processed.length;
            } else if (fileType === 'fluid') {
              uploadedData.fluid = uploadedData.fluid.concat(processed);
              uploadStats.fluid += processed.length;
            } else {
              uploadedData.balance = uploadedData.balance.concat(processed);
              uploadStats.balance += processed.length;
            }
            
            totalNewRecords += processed.length;
          }
        });

        fs.unlinkSync(file.path);
      } catch (fileError) {
        console.error(`文件 ${file.originalname} 处理失败:`, fileError);
        errors.push(`${file.originalname}: ${fileError.message}`);
      }
    }

    console.log('\n========== 上传完成 ==========');
    console.log('总新增记录数:', totalNewRecords);
    console.log(`- Balance: ${uploadStats.balance} 条 (累计: ${uploadedData.balance?.length || 0})`);
    console.log(`- Disc: ${uploadStats.disc} 条 (累计: ${uploadedData.disc?.length || 0})`);
    console.log(`- Pads: ${uploadStats.pads} 条 (累计: ${uploadedData.pads?.length || 0})`);
    console.log(`- Moto: ${uploadStats.moto} 条 (累计: ${uploadedData.moto?.length || 0})`);
    console.log(`- Fluid: ${uploadStats.fluid} 条 (累计: ${uploadedData.fluid?.length || 0})`);

    res.json({
      message: `成功上传 ${req.files.length} 个文件`,
      newRecords: totalNewRecords,
      totalRecords: uploadedData.balance?.length || 0,
      uploadStats: uploadStats,
      errors: errors.length > 0 ? errors : undefined
    });
  } catch (error) {
    console.error('库存文件上传错误:', error);
    res.status(500).json({ error: error.message });
  }
});

// 获取 Balance 数据
app.get('/api/data/balance', (req, res) => {
  try {
    if (uploadedData.balance && uploadedData.balance.length > 0) {
      res.json(uploadedData.balance);
    } else {
      res.json([]);
    }
  } catch (error) {
    console.error('获取 Balance 数据错误:', error);
    res.status(500).json({ error: error.message });
  }
});

// 获取 Balance Overview 
app.get('/api/data/balance-overview', (req, res) => {
  try {
    if (!uploadedData.balance || uploadedData.balance.length === 0) {
      return res.json({ overall: [], goodsInTransit: [] });
    }

    if (uploadedData.balance.length > 0) {
      console.log('Balance 数据示例字段:', Object.keys(uploadedData.balance[0]));
    }

    const monthlyData = {};

    uploadedData.balance.forEach(item => {
      let month = item['year_month'] || item['月份'] || item['Month'] || item['期间'] || item['Date'] || item['日期'];
      if (month && /^\d{4}$/.test(month)) {
        const yy = month.slice(0,2);
        const mm = month.slice(2,4);
        month = `20${yy}-${mm}`;
      }
      if (!month) {
        console.log('找不到月份字段，可用字段:', Object.keys(item));
        return;
      }

      if (!monthlyData[month]) {
        monthlyData[month] = {
          month: month,
          overall: 0,
          goodsInTransit: 0
        };
      }

      const inventory = item['1400-INVENTORY'];
      if (inventory !== null && inventory !== undefined && !isNaN(inventory)) {
        monthlyData[month].overall += Number(inventory);
      }

      const transit1 = item['102011101 - 在途物资—产成品 刹车盘'];
      const transit2 = item['102011102 - 在途物资—产成品 刹车片'];
      
      if (transit1 !== null && transit1 !== undefined && !isNaN(transit1)) {
        monthlyData[month].goodsInTransit += Number(transit1);
      }
      if (transit2 !== null && transit2 !== undefined && !isNaN(transit2)) {
        monthlyData[month].goodsInTransit += Number(transit2);
      }
    });

    const result = Object.values(monthlyData).sort((a, b) => {
      return a.month > b.month ? 1 : -1;
    });

    res.json({
      overall: result.map(r => ({ month: r.month, value: r.overall })),
      goodsInTransit: result.map(r => ({ month: r.month, value: r.goodsInTransit }))
    });
  } catch (error) {
    console.error('Balance Overview 错误:', error);
    res.status(500).json({ error: error.message });
  }
});

const server = app.listen(PORT, '0.0.0.0', () => {
  console.log(`服务器运行在 http://localhost:${PORT}`);
  console.log(`API 接口可用:`);
  console.log(`  GET  /api/health         - 健康检查`);
  console.log(`  GET  /api/data/sales     - 获取销售数据`);
  console.log(`  GET  /api/data/products  - 获取产品数据`);
  console.log(`  GET  /api/data/regions   - 获取地区数据`);
  console.log(`  GET  /api/data/year      - 获取年份数据`);
  console.log(`  GET  /api/stats          - 获取统计数据`);
  console.log(`  POST /api/upload         - 上传 Excel 文件`);
  console.log(`  POST /api/upload/balance - 上传 Balance 文件`);
  console.log(`  GET  /api/data/balance   - 获取 Balance 数据`);
  console.log(`  GET  /api/data/balance-overview - 获取 Balance Overview`);
  console.log(`  GET  /api/data/disc            - 获取 Disc 数据`);
  console.log(`  GET  /api/data/pads            - 获取 Pads 数据`);
  console.log(`  GET  /api/data/moto            - 获取 Moto 数据`);
  console.log(`  GET  /api/data/fluid           - 获取 Fluid 数据`);
});

server.on('error', (err) => {
  console.error('服务器启动错误:', err);
});