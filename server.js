import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import multer from 'multer';
import XLSX from 'xlsx';

// Load environment variables
dotenv.config();

const app = express();
const port = process.env.PORT || 3000;

// Resolve static folder paths
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const publicPath = path.join(__dirname, 'public');

app.use(cors());
app.use(express.json());
app.use(express.static(publicPath));

// Configure multer for temp file uploads
const upload = multer({ dest: path.join(__dirname, 'uploads/') });

// Constants for file paths
const COUPANG_FILE_PATH = path.join(__dirname, 'coupang.xlsx');
const UPDATES_FILE_PATH = path.join(__dirname, 'updated_records.json');

// In-memory cache of parsed sheets and joined records
let rawSheets = {
  q1: [],
  q2: [],
  seller: []
};
let joinedRecords = [];

/**
 * Perform Q1 and Q2 Left Joins and calculate commissions
 */
let agencyGrowthRate = 0;
let agencyQoqRate = 0;
let agencyIncentiveRate = 0;
let agencyQ1QoqTotal = 0;
let agencyQ2QoqTotal = 0;
let q1QoqTotalFromFile = 503708883;

// Q1 Group bases for chart comparison
let q1Group1AdSpend = 5506674;
let q1Group2AdSpend = 17493317;
let q1Group3AdSpend = 490090134;
let q1Group4AdSpend = 21605441;

// Baseline tracking variables
let currentQ1QoqTotal = 242727142; // default fallback
let expectedQ1QoqTotal = 0;
let expectedQ2QoqTotal = 0;
let q1VendorLevelData = {}; // vendorId -> { q1AdSpend, q1ColQ }
let mayVendorlistSet = new Set(); // vendorId target list


function performJoinsAndCalculations() {
  console.log('Calculating commissions and joining sheets...');
  
  // Create quick lookup mappings
  const q1Lookup = {};
  rawSheets.q1.forEach(row => {
    const code = String(row.vendorid || '').trim();
    if (code) {
      q1Lookup[code] = {
        q1Revenue: Number(row['이번 분기 광고비']) || 0,
        team: String(row['팀'] || '').trim(),
        marketer: String(row['마케터'] || '').trim()
      };
    }
  });

  const sellerLookup = {};
  rawSheets.seller.forEach(row => {
    const vendorId = String(row.vendorid || '').trim();
    if (vendorId) {
      sellerLookup[vendorId] = String(row['vendor name'] || '').trim();
    }
  });

  // Calculate group-level aggregates from Fact Table (rawSheets.q2)
  let q1_qoq_total = 0;
  let q2_qoq_total = 0;
  
  const processedVendors = rawSheets.q2.map(row => {
    const vendorId = String(row.vendorid || '').trim();
    const commissionGroup = String(row['커미션 그룹'] || '').trim();
    
    const q2Revenue = Number(row['2026_Q2_광고비(~260513)']) || 0;
    const q2AorRevenue = Number(row['2026_Q2_권한설정 광고비']) || 0;
    const q290DayRevenue = Number(row['2026_Q2_90일 커미션대상 광고비']) || 0;
    const q2QoqRevenue = Number(row['2026_Q2_QoQ 대상 권한설정 광고비']) || 0;
    const q1QoqBase = Number(row['직전분기 광고비(D열의 동기간)']) || 0;

    // Accumulate for agency-wide growth: Group 2 and Group 3 QoQ revenues
    if (commissionGroup === 'group 3') {
      q2_qoq_total += q2Revenue; // Use 2Q total spend for Group 3 since they are QoQ-eligible
      q1_qoq_total += q1QoqBase;
    }
    if (commissionGroup === 'group 2') {
      if (q2QoqRevenue > 0) {
        q2_qoq_total += q2QoqRevenue;
        q1_qoq_total += q1QoqBase;
      }
    }
    
    return {
      row,
      vendorId,
      commissionGroup,
      q2Revenue,
      q2AorRevenue,
      q290DayRevenue,
      q2QoqRevenue,
      q1QoqBase
    };
  });

  // Calculate Agency-wide growth and rates (compared actual 2Q QoQ sales directly to actual 1Q QoQ sales)
  agencyGrowthRate = currentQ1QoqTotal > 0 ? (q2_qoq_total / currentQ1QoqTotal - 1) : 0;
  const agency_growth_percent = agencyGrowthRate * 100;
  agencyQ1QoqTotal = currentQ1QoqTotal;
  agencyQ2QoqTotal = q2_qoq_total;
  
  agencyQoqRate = 0;
  agencyIncentiveRate = 0;
  
  if (agencyGrowthRate > 0) {
    if (agency_growth_percent >= 0 && agency_growth_percent < 10) {
      agencyQoqRate = 0.05;
      agencyIncentiveRate = 0.02;
    } else if (agency_growth_percent >= 10 && agency_growth_percent < 20) {
      agencyQoqRate = 0.15;
      agencyIncentiveRate = 0.05;
    } else if (agency_growth_percent >= 20 && agency_growth_percent < 30) {
      agencyQoqRate = 0.175;
      agencyIncentiveRate = 0.07;
    } else if (agency_growth_percent >= 30) {
      agencyQoqRate = 0.20;
      agencyIncentiveRate = 0.10;
    }
  }

  // Calculate dynamic baseline metrics based on the new rules
  expectedQ1QoqTotal = q1QoqTotalFromFile; // sum of QoQ growth 벤더 이번 분기 광고비 (₩503,708,883)
  
  let expectedQ2Sum = 0;
  processedVendors.forEach(vendor => {
    const vendorId = vendor.vendorId;
    const group = vendor.commissionGroup;
    let vendorQoqBase = 0;
    if (group === 'group 3') {
      vendorQoqBase = vendor.q2Revenue;
    } else if (group === 'group 2') {
      vendorQoqBase = vendor.q2QoqRevenue;
    }
    
    if (mayVendorlistSet.has(vendorId)) {
      expectedQ2Sum += vendorQoqBase * 91 / 43;
    } else {
      expectedQ2Sum += vendorQoqBase;
    }
  });
  expectedQ2QoqTotal = Math.round(expectedQ2Sum);

  // Join Fact Table (processedVendors) with Dimensions (LEFT JOIN)
  joinedRecords = processedVendors.map((vendor, index) => {
    const vendorId = vendor.vendorId;
    
    // Left Join dimensions
    const q1Data = q1Lookup[vendorId] || {};
    const vendorName = sellerLookup[vendorId] || '미배정';
    const team = q1Data.team || '미배정';
    const marketer = q1Data.marketer || '미배정';
    const q1Revenue = q1Data.q1Revenue || 0;
    const q2Revenue = vendor.q2Revenue;
    
    // Growth metrics (vendor-level)
    const growthAmount = q2Revenue - q1Revenue;
    const growthRatePercent = q1Revenue > 0 ? (growthAmount / q1Revenue) * 100 : (q2Revenue > 0 ? 100.0 : 0.0);
    
    // 15% new commission: Group 1, Group 2: [2Q 90일 커미션대상 광고비] * 0.15 (rest are 0)
    const isNewCommission = ['group 1', 'group 2'].includes(vendor.commissionGroup);
    const commission15 = isNewCommission ? vendor.q290DayRevenue * 0.15 : 0;
    
    // QoQ Commission eligibility
    const isQoqEligible = ['group 2', 'group 3', 'group 4'].includes(vendor.commissionGroup);
    
    // Expected calculated amounts (using agency-wide rates)
    const expectedQoqCommission = isQoqEligible ? vendor.q2QoqRevenue * agencyQoqRate : 0;

    // Marketer incentives:
    // 1) New commission: 5% of New Target sales (comm90Target * 0.05) paid regardless of growth
    const marketerNewIncentive = isNewCommission ? vendor.q290DayRevenue * 0.05 : 0;
    // 2) QoQ incentive: QoQ Revenue * agencyIncentiveRate
    const marketerQoqIncentive = isQoqEligible ? vendor.q2QoqRevenue * agencyIncentiveRate : 0;
    const expectedMarketerIncentive = marketerNewIncentive + marketerQoqIncentive;

    const isMayVendor = mayVendorlistSet.has(vendorId);
    const q2PredictedRevenue = isMayVendor ? Math.round(q2Revenue * 91 / 43) : q2Revenue;

    return {
      id: vendorId || `EXCEL_${index}`,
      vendorId,
      vendorName,
      team,
      marketer,
      commissionGroup: vendor.commissionGroup,
      q1Revenue,
      q2Revenue,
      q2PredictedRevenue,
      growthAmount,
      growthRatePercent,
      isNewCommission,
      commission15, // agency 15% new commission
      isQoqEligible,
      qoqCommissionRate: agencyQoqRate,
      expectedQoqCommission, // agency QoQ commission
      marketerIncentiveRate: agencyIncentiveRate,
      marketerNewIncentive,
      marketerGrowthIncentive: marketerQoqIncentive,
      expectedMarketerIncentive, // marketer total incentive
      isUpdated: false // Reset update status since it's now direct fact table upload
    };
  });

  console.log(`Successfully compiled and calculated ${joinedRecords.length} vendor records.`);
}

/**
 * Normalizes header names to handle variances in spreadsheet imports.
 */
function normalizeHeaderName(h) {
  const norm = String(h || '').trim().toLowerCase().replace(/[\s_-]/g, '');
  if (norm.includes('2026q2광고비') || norm.includes('q2광고비') || norm.includes('q2sales') || norm === '광고비' || norm === '매출' || norm === '매출액' || norm === '실적') {
    return '2026_Q2_광고비(~260513)';
  }
  if (norm.includes('직전분기광고비') || norm.includes('직전분기광고비(d열의동기간)') || norm.includes('직전분기매출')) {
    return '직전분기 광고비(D열의 동기간)';
  }
  if (norm.includes('qoq대상권한설정') || norm.includes('2026q2qoq대상권한설정') || norm.includes('qoq대상광고비')) {
    return '2026_Q2_QoQ 대상 권한설정 광고비';
  }
  if (norm.includes('권한설정광고비') || norm.includes('2026q2권한설정광고비')) {
    return '2026_Q2_권한설정 광고비';
  }
  if (norm.includes('90일커미션') || norm.includes('2026q290일커미션') || norm.includes('90일커미션대상')) {
    return '2026_Q2_90일 커미션대상 광고비';
  }
  if (norm.includes('커미션그룹') || norm === '그룹') {
    return '커미션 그룹';
  }
  if (norm === 'vendorid' || norm === '업체코드' || norm === 'vendor' || norm === 'vendorid코드' || norm === '업체id' || norm === '벤더id') {
    return 'vendorid';
  }
  return h;
}

/**
 * Robust extraction of Vendor ID and Sales values from custom update sheets
 */
function preprocessSalesFile(filePath) {
  const workbook = XLSX.readFile(filePath);
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  
  // Convert worksheet to 2D array
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
  if (rows.length < 2) {
    throw new Error('File has insufficient rows.');
  }

  // 1. Skiprows=1: Skip first row (index 0). Row 2 (index 1) is the candidate header.
  
  // 2. Identify vendorid column by scanning cells for patterns starting with 'A' followed by digits (e.g. A00239521)
  let vendorColIdx = -1;
  let maxMatches = 0;
  for (let c = 0; c < 30; c++) { // scan first 30 columns
    let matches = 0;
    for (let r = 2; r < rows.length; r++) {
      const val = String(rows[r][c] || '').trim();
      if (/^A\d+$/.test(val)) {
        matches++;
      }
    }
    if (matches > maxMatches) {
      maxMatches = matches;
      vendorColIdx = c;
    }
  }

  if (vendorColIdx === -1) {
    // Fallback: look for startsWith 'A00' in data
    for (let c = 0; c < 30; c++) {
      for (let r = 1; r < rows.length; r++) {
        const val = String(rows[r][c] || '').trim();
        if (val.startsWith('A00')) {
          vendorColIdx = c;
          break;
        }
      }
      if (vendorColIdx !== -1) break;
    }
  }

  if (vendorColIdx === -1) {
    throw new Error('벤더 ID 컬럼을 찾을 수 없습니다. (예: A00으로 시작하는 코드)');
  }

  // Find where actual data starts (the first row with a valid vendor ID matching /^A\d+$/)
  let firstVendorRowIdx = -1;
  for (let r = 2; r < rows.length; r++) {
    const val = String(rows[r][vendorColIdx]).trim();
    if (/^A\d+$/.test(val)) {
      firstVendorRowIdx = r;
      break;
    }
  }

  if (firstVendorRowIdx === -1) {
    throw new Error('Coupang 벤더 ID를 포함한 데이터 행이 존재하지 않습니다.');
  }

  // The actual header row for the vendor table is the row right before firstVendorRowIdx
  const rawHeaders = rows[firstVendorRowIdx - 1].map(h => String(h || '').trim());
  rawHeaders[vendorColIdx] = 'vendorid';

  // Normalize header names to guarantee consistency in parsing
  const headers = rawHeaders.map((h, idx) => {
    if (idx === vendorColIdx) return 'vendorid';
    return normalizeHeaderName(h) || `col_${idx}`;
  });

  console.log(`ETL: Detected vendor ID at column ${vendorColIdx}. Data starts at row ${firstVendorRowIdx}. Normalizing headers to:`, headers);

  // Clean data rows
  const cleanRows = [];
  for (let r = firstVendorRowIdx; r < rows.length; r++) {
    const rowVal = rows[r];
    const vendorId = String(rowVal[vendorColIdx] || '').trim();

    // 3. Drop rows where vendorid is null, empty or does not match Coupang ID regex
    if (!vendorId || !/^A\d+$/.test(vendorId)) {
      continue;
    }

    const obj = {};
    headers.forEach((h, idx) => {
      let val = rowVal[idx];

      // 4. Sanitize numeric fields: remove commas and parse as float
      if (typeof val === 'string') {
        const cleanVal = val.replace(/,/g, '').trim();
        if (cleanVal !== '' && !isNaN(cleanVal)) {
          if (h !== 'vendorid' && h !== '최초 권한 설정일' && h !== '권한 해제일') {
            val = parseFloat(cleanVal);
          }
        }
      }
      obj[h] = val;
    });

    cleanRows.push(obj);
  }

  return cleanRows;
}

/**
 * Load Base excel sheet and local updates
 */
function loadAndInitializeData() {
  console.log('--- Initializing Coupang Ads Dashboard Data ---');
  
  if (!fs.existsSync(COUPANG_FILE_PATH)) {
    console.error(`CRITICAL ERROR: 'coupang.xlsx' not found at ${COUPANG_FILE_PATH}`);
    return;
  }

  try {
    console.log(`Reading base dimensions: ${COUPANG_FILE_PATH} (This may take a moment)...`);
    const workbook = XLSX.readFile(COUPANG_FILE_PATH);
    
    // Parse sheet: 1Q
    const q1Sheet = workbook.Sheets['1Q'];
    if (q1Sheet) {
      const q1Rows = XLSX.utils.sheet_to_json(q1Sheet, { defval: null });
      rawSheets.q1 = q1Rows.map(row => {
        const vendorid = String(row['업체코드'] || row['vendorid'] || '').trim();
        return { ...row, vendorid };
      }).filter(row => /^A\d+$/.test(row.vendorid));
      console.log(`Loaded sheet '1Q': ${rawSheets.q1.length} cleaned rows`);
    }

    // Parse sheet: 2Q seller
    const sellerSheet = workbook.Sheets['2Q seller'];
    if (sellerSheet) {
      const sellerRows = XLSX.utils.sheet_to_json(sellerSheet, { defval: null });
      rawSheets.seller = sellerRows.map(row => {
        const vendorid = String(row['vendor id'] || row['vendorid'] || '').trim();
        return { ...row, vendorid };
      }).filter(row => /^A\d+$/.test(row.vendorid));
      console.log(`Loaded sheet '2Q seller': ${rawSheets.seller.length} cleaned rows`);
    }
  } catch (error) {
    console.error('Error reading and parsing coupang.xlsx:', error);
  }

  // Parse Q1 Agency commission file to get Q1 QoQ base and vendor dimensions
  const q1AgencyPath = path.join(__dirname, '2026Q1 Agency commission-MP(PA).xlsx');
  if (fs.existsSync(q1AgencyPath)) {
    try {
      console.log(`Reading Q1 Agency file: ${q1AgencyPath}...`);
      const q1Wb = XLSX.readFile(q1AgencyPath);
      
      // Parse summary sheet
      const q1SummarySheet = q1Wb.Sheets['Total summary'];
      if (q1SummarySheet) {
        const q1SummaryRows = XLSX.utils.sheet_to_json(q1SummarySheet);
        let sumQoq = 0;
        
        let g1New = 0;
        let g2New = 0, g2Qoq = 0;
        let g3Qoq = 0;
        let g4Qoq = 0;

        q1SummaryRows.forEach(row => {
          const group = String(row['커미션 그룹'] || '').trim();
          const val = Number(row['QoQ growth 벤더 이번 분기 광고비']) || 0;
          if (group === 'group 2' || group === 'group 3') {
            sumQoq += val;
          }

          const newVendorsAdSpend = Number(row['90일 이내 신규/재활성화 벤더 광고비']) || 0;
          const qoqAdSpend = Number(row['QoQ growth 수수료 대상 광고비']) || 0;
          
          if (group === 'group 1') {
            g1New = newVendorsAdSpend;
          } else if (group === 'group 2') {
            g2New = newVendorsAdSpend;
            g2Qoq = qoqAdSpend;
          } else if (group === 'group 3') {
            g3Qoq = qoqAdSpend;
          } else if (group === 'group 4') {
            g4Qoq = qoqAdSpend;
          }
        });
        if (sumQoq > 0) {
          q1QoqTotalFromFile = sumQoq;
          console.log(`Q1 QoQ base loaded from Q1 Agency file: ${q1QoqTotalFromFile}`);
        }

        if (g1New > 0) q1Group1AdSpend = g1New;
        if (g2New > 0 || g2Qoq > 0) q1Group2AdSpend = g2New + g2Qoq;
        if (g3Qoq > 0) q1Group3AdSpend = g3Qoq;
        if (g4Qoq > 0) q1Group4AdSpend = g4Qoq;

        console.log(`Chart Q1 bases loaded dynamically: G1=${q1Group1AdSpend}, G2=${q1Group2AdSpend}, G3=${q1Group3AdSpend}, G4=${q1Group4AdSpend}`);
      }

      // Load and cache Q1 Vendor level columns for baseline calculation
      const q1VendorSheet = q1Wb.Sheets['Vendor level'];
      if (q1VendorSheet) {
        const q1VendorRows = XLSX.utils.sheet_to_json(q1VendorSheet, { header: 1, defval: '' });
        const q1HeaderRow = q1VendorRows[0];
        const vendorColIdx = q1HeaderRow.indexOf('업체코드');
        const colPIdx = q1HeaderRow.indexOf('이번 분기 광고비');
        const colQIdx = q1HeaderRow.indexOf('이번 분기 권한 설정 광고비');
        
        q1VendorLevelData = {};
        for (let i = 1; i < q1VendorRows.length; i++) {
          const row = q1VendorRows[i];
          const vendorId = String(row[vendorColIdx] || '').trim();
          if (vendorId && /^A\d+$/.test(vendorId)) {
            q1VendorLevelData[vendorId] = {
              q1AdSpend: Number(row[colPIdx]) || 0,
              q1ColQ: Number(row[colQIdx]) || 0
            };
          }
        }
        console.log(`Cached Q1 Vendor level data for ${Object.keys(q1VendorLevelData).length} vendors.`);
      }
    } catch (err) {
      console.error('Error reading Q1 Agency file:', err.message);
    }
  }

  // Parse May vendorlist file
  const mayVendorlistPath = path.join(__dirname, '2026_Q2_May_vendorlist_MP.xlsx');
  if (fs.existsSync(mayVendorlistPath)) {
    try {
      console.log(`Reading May vendorlist file: ${mayVendorlistPath}...`);
      const mayWb = XLSX.readFile(mayVendorlistPath);
      const maySheet = mayWb.Sheets[mayWb.SheetNames[0]];
      const mayRows = XLSX.utils.sheet_to_json(maySheet, { header: 1, defval: '' });
      const mayHeader = mayRows[1];
      const mayVendorColIdx = mayHeader.indexOf('vendor id') !== -1 ? mayHeader.indexOf('vendor id') : mayHeader.indexOf('vender id');
      
      mayVendorlistSet = new Set();
      for (let i = 2; i < mayRows.length; i++) {
        const vId = String(mayRows[i][mayVendorColIdx] || '').trim();
        if (/^A\d+$/.test(vId)) {
          mayVendorlistSet.add(vId);
        }
      }
      console.log(`Loaded ${mayVendorlistSet.size} target vendor IDs from May vendor list.`);
    } catch (err) {
      console.error('Error reading May vendorlist file:', err.message);
    }
  }

  // Load Fact Table (either latest uploaded or default 260515_MP.xlsx)
  const latestSalesPath = path.join(__dirname, 'uploads', 'latest_sales.xlsx');
  let factFilePath = COUPANG_FILE_PATH; // fallback
  let factLoaded = false;

  if (fs.existsSync(latestSalesPath)) {
    factFilePath = latestSalesPath;
    console.log(`Found uploaded sales update at ${latestSalesPath}`);
  } else {
    const defaultSalesPath = path.join(__dirname, '260515_MP.xlsx');
    if (fs.existsSync(defaultSalesPath)) {
      factFilePath = defaultSalesPath;
      console.log(`Found default sales file at ${defaultSalesPath}`);
    } else {
      console.warn('Neither uploaded nor default sales file found. Falling back to coupang.xlsx sheet 2Q.');
    }
  }

  try {
    if (factFilePath === COUPANG_FILE_PATH) {
      const workbook = XLSX.readFile(COUPANG_FILE_PATH);
      const q2Sheet = workbook.Sheets['2Q'];
      if (q2Sheet) {
        const rows = XLSX.utils.sheet_to_json(q2Sheet, { defval: null });
        rawSheets.q2 = rows.map(row => {
          const vendorid = String(row.vendorid || '').trim();
          return { ...row, vendorid };
        }).filter(row => /^A\d+$/.test(row.vendorid));
        factLoaded = true;
      }
    } else {
      console.log(`Preprocessing sales Fact Table from ${factFilePath}...`);
      rawSheets.q2 = preprocessSalesFile(factFilePath);

      // Extract currentQ1QoqTotal dynamically from the summary section of the Q2 sales file
      try {
        const workbook = XLSX.readFile(factFilePath);
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        let actualColIdx = -1;
        for (let r = 0; r < Math.min(10, rows.length); r++) {
          const row = rows[r];
          for (let c = 0; c < row.length; c++) {
            const val = String(row[c]).trim();
            if (val.includes('% QoQ PA growth (QTD) (actual)')) {
              actualColIdx = c;
              break;
            }
          }
          if (actualColIdx !== -1) {
            if (rows[r + 1]) {
              const rawVal = Number(rows[r + 1][actualColIdx]);
              if (!isNaN(rawVal) && rawVal > 0) {
                currentQ1QoqTotal = rawVal;
                console.log(`Found currentQ1QoqTotal dynamically: ${currentQ1QoqTotal}`);
              }
            }
            break;
          }
        }
      } catch (errVal) {
        console.warn('Could not extract currentQ1QoqTotal dynamically, using default fallback.', errVal.message);
      }

      factLoaded = true;
    }
  } catch (err) {
    console.error(`Error loading Fact Table from ${factFilePath}:`, err.message);
  }

  if (factLoaded) {
    performJoinsAndCalculations();
  } else {
    console.error('CRITICAL: Failed to load Fact Table.');
  }
}

// Perform initial data loading
loadAndInitializeData();

// Create uploads folder if not exists
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir);
}

// ----------------------------------------------------
// API ENDPOINTS
// ----------------------------------------------------

/**
 * GET /api/base-data
 * Returns the cached, fully joined and calculated rows
 */
app.get('/api/base-data', (req, res) => {
  if (joinedRecords.length === 0) {
    return res.status(500).json({ error: 'Base data not loaded or failed to parse coupang.xlsx' });
  }
  const latestSalesPath = path.join(__dirname, 'uploads', 'latest_sales.xlsx');
  const hasUpload = fs.existsSync(latestSalesPath);
  res.json({
    records: joinedRecords,
    agencyGrowthRate,
    agencyQoqRate,
    agencyIncentiveRate,
    q1QoqTotal: currentQ1QoqTotal,
    q2QoqTotal: agencyQ2QoqTotal,
    currentQ1QoqTotal,
    currentQ2QoqTotal: agencyQ2QoqTotal,
    expectedQ1QoqTotal,
    expectedQ2QoqTotal,
    lastUpdated: hasUpload ? fs.statSync(latestSalesPath).mtime : null,
    totalUpdatesCount: hasUpload ? 1 : 0,
    q1ChartValues: {
      'group 1': q1Group1AdSpend,
      'group 2': q1Group2AdSpend,
      'group 3': q1Group3AdSpend,
      'group 4': q1Group4AdSpend
    }
  });
});

/**
 * POST /api/upload-update
 * Endpoint to upload a single update file (Excel/CSV)
 */
app.post('/api/upload-update', upload.single('updateFile'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }

  const filePath = req.file.path;
  const targetPath = path.join(__dirname, 'uploads', 'latest_sales.xlsx');
  
  try {
    console.log(`Processing uploaded sales file: ${req.file.originalname}`);
    const preprocessedRows = preprocessSalesFile(filePath);
    
    // Clean up existing file if any
    if (fs.existsSync(targetPath)) {
      fs.unlinkSync(targetPath);
    }
    
    // Move/rename temp file to latest_sales.xlsx
    fs.renameSync(filePath, targetPath);
    
    // Set Fact Table in memory
    rawSheets.q2 = preprocessedRows;
    
    // Perform joins and calculations
    performJoinsAndCalculations();
    
    res.json({
      success: true,
      updatedCount: preprocessedRows.length,
      totalUpdatesCount: 1,
      records: joinedRecords,
      agencyGrowthRate,
      agencyQoqRate,
      agencyIncentiveRate,
      q1QoqTotal: currentQ1QoqTotal,
      q2QoqTotal: agencyQ2QoqTotal,
      currentQ1QoqTotal,
      currentQ2QoqTotal: agencyQ2QoqTotal,
      expectedQ1QoqTotal,
      expectedQ2QoqTotal,
      q1ChartValues: {
        'group 1': q1Group1AdSpend,
        'group 2': q1Group2AdSpend,
        'group 3': q1Group3AdSpend,
        'group 4': q1Group4AdSpend
      }
    });
  } catch (error) {
    console.error('Error processing sales update upload:', error);
    res.status(500).json({ error: `File processing failed: ${error.message}` });
    
    // Clean up temporary upload file
    try {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    } catch (err) {}
  }
});

/**
 * POST /api/reset-updates
 * Clear all uploaded updates and restore original values from coupang.xlsx / 260515_MP.xlsx
 */
app.post('/api/reset-updates', (req, res) => {
  try {
    const latestSalesPath = path.join(__dirname, 'uploads', 'latest_sales.xlsx');
    if (fs.existsSync(latestSalesPath)) {
      fs.unlinkSync(latestSalesPath);
    }
    
    // Reload initial data from disk
    loadAndInitializeData();
    
    res.json({
      success: true,
      records: joinedRecords,
      agencyGrowthRate,
      agencyQoqRate,
      agencyIncentiveRate,
      q1QoqTotal: currentQ1QoqTotal,
      q2QoqTotal: agencyQ2QoqTotal,
      currentQ1QoqTotal,
      currentQ2QoqTotal: agencyQ2QoqTotal,
      expectedQ1QoqTotal,
      expectedQ2QoqTotal,
      q1ChartValues: {
        'group 1': q1Group1AdSpend,
        'group 2': q1Group2AdSpend,
        'group 3': q1Group3AdSpend,
        'group 4': q1Group4AdSpend
      }
    });
  } catch (error) {
    res.status(500).json({ error: `Reset failed: ${error.message}` });
  }
});

// Start Express Server
app.listen(port, () => {
  console.log(`Coupang Ads Dashboard backend listening at http://localhost:${port}`);
});

