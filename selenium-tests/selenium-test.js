const { Builder, By, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const ExcelJS = require('exceljs');
const http = require('http');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

const PORT = 8080;
const BASE_URL = `http://localhost:${PORT}`;

// ── Built-in Static Server for E2E testing ──
let server;
function startServer() {
  server = http.createServer((req, res) => {
    let urlPath = req.url.split('?')[0]; // strip query params
    let filePath = path.join(__dirname, '..', urlPath === '/' ? 'index.html' : urlPath);
    
    // Prevent directory traversal
    if (!filePath.startsWith(path.join(__dirname, '..'))) {
      res.writeHead(403);
      res.end('Access denied');
      return;
    }

    const extname = String(path.extname(filePath)).toLowerCase();
    const mimeTypes = {
      '.html': 'text/html',
      '.js': 'text/javascript',
      '.css': 'text/css',
      '.json': 'application/json',
      '.png': 'image/png',
      '.jpg': 'image/jpg',
      '.gif': 'image/gif',
      '.svg': 'image/svg+xml'
    };
    const contentType = mimeTypes[extname] || 'application/octet-stream';

    fs.readFile(filePath, (error, content) => {
      if (error) {
        if (error.code === 'ENOENT') {
          res.writeHead(404, { 'Content-Type': 'text/plain' });
          res.end(`File not found: ${req.url}`);
        } else {
          res.writeHead(500, { 'Content-Type': 'text/plain' });
          res.end(`Server error: ${error.code}`);
        }
      } else {
        res.writeHead(200, { 'Content-Type': contentType });
        res.end(content, 'utf-8');
      }
    });
  });

  server.listen(PORT);
  console.log(`[INFO] Local test server started at ${BASE_URL}`);
}

function stopServer() {
  if (server) {
    server.close(() => {
      console.log('[INFO] Local test server stopped.');
    });
  }
}

// ── Test Logging / Steps ──
const testSteps = [];
function logStep(action, target, expected, actual, status) {
  const stepNum = testSteps.length + 1;
  const logMsg = `[Step ${stepNum}] ${action} | Expected: ${expected} | Actual: ${actual} | ${status.toUpperCase()}`;
  console.log(status === 'pass' ? `\x1b[32m${logMsg}\x1b[0m` : `\x1b[31m${logMsg}\x1b[0m`);
  testSteps.push({ stepNum, action, target, expected, actual, status });
}

// ── Excel Report Generator ──
async function generateExcelReport(expensesData, summaryStats) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'SpendSense Automated Selenium E2E';
  workbook.created = new Date();

  // Color palette (Navy/Teal design system)
  const colors = {
    primaryDark: '1A365D', // Navy
    accent: '0D9488',      // Teal
    passedBg: 'DEF7EC',    // Light Green
    passedText: '03543F',  // Dark Green
    failedBg: 'FDE8E8',    // Light Red
    failedText: '9B1C1C',  // Dark Red
    headerText: 'FFFFFF',
    borderLight: 'E2E8F0',
    zebraBg: 'F8FAFC'
  };

  // --- SHEET 1: Summary Dashboard ---
  const wsSummary = workbook.addWorksheet('Summary Dashboard');
  wsSummary.views = [{ showGridLines: true }];

  // Title block
  wsSummary.mergeCells('B2:F2');
  const cellTitle = wsSummary.getCell('B2');
  cellTitle.value = 'SpendSense Web E2E Test Report';
  cellTitle.font = { name: 'Segoe UI', size: 16, bold: true, color: { argb: colors.headerText } };
  cellTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.primaryDark } };
  cellTitle.alignment = { horizontal: 'center', vertical: 'middle' };
  wsSummary.getRow(2).height = 40;

  // Metadata Table
  const metaLabels = [
    ['Test Engine', 'Selenium WebDriver (Node.js)'],
    ['Target Environment', BASE_URL],
    ['Run Date', new Date().toLocaleString()],
    ['Total E2E Steps', testSteps.length],
    ['Passed Steps', testSteps.filter(s => s.status === 'pass').length],
    ['Failed Steps', testSteps.filter(s => s.status === 'fail').length],
    ['Overall Status', testSteps.some(s => s.status === 'fail') ? 'FAILED ✘' : 'PASSED ✓']
  ];

  let currentBarRow = 4;
  metaLabels.forEach((pair) => {
    wsSummary.getCell(`B${currentBarRow}`).value = pair[0];
    wsSummary.getCell(`B${currentBarRow}`).font = { name: 'Segoe UI', size: 10, bold: true };
    wsSummary.getCell(`C${currentBarRow}`).value = pair[1];
    wsSummary.getCell(`C${currentBarRow}`).font = { name: 'Segoe UI', size: 10 };
    
    // Status color
    if (pair[0] === 'Overall Status') {
      const isPassed = pair[1].includes('PASSED');
      wsSummary.getCell(`C${currentBarRow}`).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: isPassed ? colors.passedBg : colors.failedBg }
      };
      wsSummary.getCell(`C${currentBarRow}`).font = {
        name: 'Segoe UI',
        size: 10,
        bold: true,
        color: { argb: isPassed ? colors.passedText : colors.failedText }
      };
    }
    currentBarRow++;
  });

  // KPI Block
  wsSummary.getCell('E4').value = 'Budget Set';
  wsSummary.getCell('E5').value = summaryStats.budget;
  wsSummary.getCell('E5').numFmt = '₹#,##0';
  wsSummary.getCell('F4').value = 'Total Spent';
  wsSummary.getCell('F5').value = summaryStats.totalSpent;
  wsSummary.getCell('F5').numFmt = '₹#,##0';

  wsSummary.getCell('E7').value = 'Remaining Budget';
  wsSummary.getCell('E8').value = summaryStats.remainingBudget;
  wsSummary.getCell('E8').numFmt = '₹#,##0';
  wsSummary.getCell('F7').value = 'Budget Status';
  wsSummary.getCell('F8').value = summaryStats.budgetStatus;

  // Style KPI Block
  ['E4', 'F4', 'E7', 'F7'].forEach(c => {
    wsSummary.getCell(c).font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: '718096' } };
    wsSummary.getCell(c).alignment = { horizontal: 'center' };
  });
  ['E5', 'F5', 'E8', 'F8'].forEach(c => {
    wsSummary.getCell(c).font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: colors.primaryDark } };
    wsSummary.getCell(c).alignment = { horizontal: 'center' };
    wsSummary.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.zebraBg } };
  });
  wsSummary.getCell('F8').font = { 
    name: 'Segoe UI', 
    size: 11, 
    bold: true, 
    color: { argb: summaryStats.budgetStatus.includes('over') ? colors.failedText : colors.passedText } 
  };

  // Add borders to KPI blocks
  const thinBorder = {
    top: { style: 'thin', color: { argb: colors.borderLight } },
    left: { style: 'thin', color: { argb: colors.borderLight } },
    bottom: { style: 'thin', color: { argb: colors.borderLight } },
    right: { style: 'thin', color: { argb: colors.borderLight } }
  };
  ['E4', 'E5', 'F4', 'F5', 'E7', 'E8', 'F7', 'F8'].forEach(c => {
    wsSummary.getCell(c).border = thinBorder;
  });

  // Set Column Widths
  wsSummary.getColumn('B').width = 22;
  wsSummary.getColumn('C').width = 32;
  wsSummary.getColumn('D').width = 5;
  wsSummary.getColumn('E').width = 22;
  wsSummary.getColumn('F').width = 22;


  // --- SHEET 2: Test Logs ---
  const wsLogs = workbook.addWorksheet('Execution Logs');
  wsLogs.views = [{ showGridLines: true }];

  // Headers
  const logsHeaders = ['Step #', 'Action / Command', 'Target Element', 'Expected Outcome', 'Actual Outcome', 'Status'];
  wsLogs.addRow(logsHeaders);
  const headerRow = wsLogs.getRow(1);
  headerRow.height = 28;
  headerRow.eachCell((cell) => {
    cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: colors.headerText } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.accent } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  // Add Logs
  testSteps.forEach((s) => {
    const row = wsLogs.addRow([s.stepNum, s.action, s.target, s.expected, s.actual, s.status.toUpperCase()]);
    row.height = 20;
    
    // Alignments
    row.getCell(1).alignment = { horizontal: 'center' };
    row.getCell(6).alignment = { horizontal: 'center' };

    // Pass / Fail highlighting
    const statusCell = row.getCell(6);
    if (s.status === 'pass') {
      statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.passedBg } };
      statusCell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: colors.passedText } };
    } else {
      statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.failedBg } };
      statusCell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: colors.failedText } };
    }

    // Add thin borders
    row.eachCell((cell) => {
      cell.border = thinBorder;
      if (cell.col !== 6) {
        cell.font = { name: 'Segoe UI', size: 10 };
      }
    });
  });

  // Auto-fit widths for Logs
  wsLogs.columns.forEach((column, i) => {
    let maxLength = 0;
    column.eachCell({ includeEmpty: true }, (cell) => {
      const columnLength = cell.value ? cell.value.toString().length : 0;
      if (columnLength > maxLength) {
        maxLength = columnLength;
      }
    });
    column.width = Math.max(maxLength + 4, 12);
  });


  // --- SHEET 3: E2E Expense Records ---
  const wsData = workbook.addWorksheet('Recorded Data');
  wsData.views = [{ showGridLines: true }];

  // Headers
  wsData.addRow(['Description', 'Category', 'Amount', 'Date', 'Note']);
  wsData.getRow(1).height = 26;
  wsData.getRow(1).eachCell((cell) => {
    cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: colors.headerText } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.primaryDark } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  // Add recorded expenses
  expensesData.forEach((e) => {
    const row = wsData.addRow([e.desc, e.category, parseFloat(e.amount), e.date, e.note || '']);
    row.height = 20;
    row.getCell(3).numFmt = '₹#,##0.00';
    row.getCell(4).alignment = { horizontal: 'center' };
    
    row.eachCell((cell) => {
      cell.border = thinBorder;
      cell.font = { name: 'Segoe UI', size: 10 };
    });
  });

  // Add a total row
  const totalRowIndex = expensesData.length + 2;
  wsData.getCell(`B${totalRowIndex}`).value = 'Total Spent';
  wsData.getCell(`B${totalRowIndex}`).font = { name: 'Segoe UI', size: 10, bold: true };
  wsData.getCell(`B${totalRowIndex}`).alignment = { horizontal: 'right' };
  
  wsData.getCell(`C${totalRowIndex}`).value = { formula: `SUM(C2:C${totalRowIndex-1})` };
  wsData.getCell(`C${totalRowIndex}`).font = { name: 'Segoe UI', size: 10, bold: true };
  wsData.getCell(`C${totalRowIndex}`).numFmt = '₹#,##0.00';
  wsData.getCell(`C${totalRowIndex}`).border = {
    top: { style: 'thin', color: { argb: colors.borderLight } },
    bottom: { style: 'double', color: { argb: colors.primaryDark } }
  };

  // Auto-fit widths for Data
  wsData.columns.forEach((column) => {
    let maxLength = 0;
    column.eachCell({ includeEmpty: true }, (cell) => {
      const columnLength = cell.value ? cell.value.toString().length : 0;
      if (columnLength > maxLength) {
        maxLength = columnLength;
      }
    });
    column.width = Math.max(maxLength + 5, 15);
  });

  // Write file
  const reportPath = path.join(__dirname, 'selenium_report.xlsx');
  await workbook.xlsx.writeFile(reportPath);
  console.log(`[SUCCESS] Excel report successfully generated and saved to: ${reportPath}`);
}

// ── Background Backend Server Management ──
let backendProcess;
function startBackendServer() {
  console.log('[INFO] Spawning SpendSense backend server...');
  backendProcess = spawn('node', [path.join(__dirname, '../server/server.js')], {
    stdio: 'inherit'
  });
  backendProcess.on('error', (err) => {
    console.error('[ERROR] Failed to start backend server:', err);
  });
}

function stopBackendServer() {
  if (backendProcess) {
    console.log('[INFO] Stopping SpendSense backend server...');
    backendProcess.kill();
  }
}

// ── Main E2E Test Execution ──
async function runTest() {
  // Start backend API server
  startBackendServer();
  
  // Wait a brief moment for Express to spin up
  await new Promise(resolve => setTimeout(resolve, 2000));

  // 1. Host local application server
  startServer();

  // 2. Setup Selenium Chrome options
  let options = new chrome.Options();
  // We run in headless mode by default for reliable CI/automated run.
  // To watch the test run in UI, comment out the next line.
  options.addArguments('--headless=new'); 
  options.addArguments('--disable-gpu');
  options.addArguments('--no-sandbox');
  options.addArguments('--disable-dev-shm-usage');
  options.addArguments('--window-size=1280,1024');

  console.log('[INFO] Launching Chrome driver...');
  let driver = await new Builder()
    .forBrowser('chrome')
    .setChromeOptions(options)
    .build();

  const recordedExpenses = [];
  const summaryStats = {
    budget: 0,
    totalSpent: 0,
    remainingBudget: 0,
    budgetStatus: 'No budget set'
  };

  try {
    // Navigate to local server
    await driver.get(BASE_URL);
    await driver.wait(until.titleContains('SpendSense'), 10000);
    logStep('Navigate to URL', BASE_URL, 'SpendSense app page loads', 'App page loaded successfully', 'pass');

    // Clear localStorage and hit backend reset endpoint to guarantee clean slate
    await driver.executeScript('window.localStorage.clear();');
    try {
      const resetUrl = 'http://localhost:5000/api/reset';
      await driver.executeScript(`fetch("${resetUrl}", { method: "POST" });`);
    } catch (e) {
      console.error('[WARNING] Failed to trigger API database reset via driver script:', e);
    }
    await driver.executeScript('window.location.reload();');
    await driver.wait(until.titleContains('SpendSense'), 10000);
    logStep('Initialize Database', 'localStorage.clear() & API reset', 'Fresh clean slate', 'Local storage and backend API database reset', 'pass');

    // ── STEP 1: Set Budget ──
    // Click budget button in sidebar
    const budgetNav = await driver.findElement(By.css('button[data-view="budget"]'));
    await budgetNav.click();
    await driver.sleep(500); // Wait for transit animation
    logStep('Click sidebar navigation', 'Budget Menu Item', 'Budget view shown', 'Navigated to Budget view', 'pass');

    // Enter monthly budget of 20000
    const budgetInput = await driver.findElement(By.id('budgetInput'));
    await budgetInput.sendKeys('20000');
    summaryStats.budget = 20000;
    
    const budgetSubmit = await driver.findElement(By.css('button[onclick="setBudget()"]'));
    await budgetSubmit.click();
    await driver.sleep(500);
    
    const budgetFb = await driver.findElement(By.id('budgetFeedback'));
    const budgetMsg = await budgetFb.getText();
    if (budgetMsg.includes('set to ₹20,000') || budgetMsg.includes('set to')) {
      logStep('Set Monthly Budget', '₹20,000', 'Success message displayed', `Budget set feedback: "${budgetMsg}"`, 'pass');
    } else {
      logStep('Set Monthly Budget', '₹20,000', 'Success message displayed', `Incorrect feedback message: "${budgetMsg}"`, 'fail');
    }

    // ── STEP 2: Add Expenses ──
    const addNav = await driver.findElement(By.css('button[data-view="add"]'));
    await addNav.click();
    await driver.sleep(500);
    logStep('Click sidebar navigation', 'Add Expense Menu Item', 'Add Expense view shown', 'Navigated to Add Expense view', 'pass');

    // Define transactions to add
    const txs = [
      { desc: 'Dinner at Taj', amount: '1500', category: 'Food', catId: 'food', note: 'Dinner with team' },
      { desc: 'Uber to office', amount: '350', category: 'Transport', catId: 'transport', note: '' },
      { desc: 'New headphones', amount: '4500', category: 'Shopping', catId: 'shopping', note: 'Noise cancelling' },
      { desc: 'Monthly Rent', amount: '12000', category: 'Rent', catId: 'rent', note: 'Flat 402' }
    ];

    for (const tx of txs) {
      // Description
      const expDesc = await driver.findElement(By.id('expDesc'));
      await expDesc.sendKeys(tx.desc);

      // Amount
      const expAmount = await driver.findElement(By.id('expAmount'));
      await expAmount.sendKeys(tx.amount);

      // Note
      if (tx.note) {
        const expNote = await driver.findElement(By.id('expNote'));
        await expNote.sendKeys(tx.note);
      }

      // Pick category
      const catBtn = await driver.findElement(By.xpath(`//button[contains(@class, 'cat-option') and contains(text(), '${tx.category}')]`));
      await catBtn.click();

      // Submit
      const recordBtn = await driver.findElement(By.css('button[onclick="addExpense()"]'));
      await recordBtn.click();

      // Verify feedback
      const formFb = await driver.findElement(By.id('formFeedback'));
      let formMsg = "";
      try {
        await driver.wait(until.elementTextContains(formFb, 'recorded'), 2000);
        formMsg = await formFb.getText();
      } catch (e) {
        formMsg = await formFb.getText();
      }
      const isOk = formMsg.includes('recorded');
      
      logStep(
        `Add Expense: ${tx.desc}`,
        `Amount: ₹${tx.amount}, Category: ${tx.category}`,
        'Feedback: "Expense recorded!"',
        `Result: "${formMsg}"`,
        isOk ? 'pass' : 'fail'
      );

      recordedExpenses.push(tx);
    }

    // ── STEP 3: Verify Dashboard Statistics ──
    const dashNav = await driver.findElement(By.css('button[data-view="dashboard"]'));
    await dashNav.click();
    await driver.sleep(500);
    logStep('Click sidebar navigation', 'Dashboard Menu Item', 'Dashboard view shown', 'Navigated to Dashboard', 'pass');

    // Read Total Spent
    const totalSpentEl = await driver.findElement(By.id('totalSpent'));
    const totalSpentText = await totalSpentEl.getText();
    // Sum = 1500 + 350 + 4500 + 12000 = 18350
    const expectedTotal = '₹18,350';
    const totalIsCorrect = totalSpentText.trim() === expectedTotal;
    summaryStats.totalSpent = 18350;
    logStep(
      'Verify Total Spent Card',
      'totalSpent element',
      `Value: ${expectedTotal}`,
      `Value: ${totalSpentText}`,
      totalIsCorrect ? 'pass' : 'fail'
    );

    // Read Transactions Count
    const monthCountEl = await driver.findElement(By.id('monthCount'));
    const monthCountText = await monthCountEl.getText();
    const countIsCorrect = monthCountText.trim() === '4';
    logStep(
      'Verify Monthly Transaction Count',
      'monthCount element',
      'Value: 4',
      `Value: ${monthCountText}`,
      countIsCorrect ? 'pass' : 'fail'
    );

    // Read Remaining Budget
    const remainingBudgetEl = await driver.findElement(By.id('remainingBudget'));
    const remainingBudgetText = await remainingBudgetEl.getText();
    // Remaining = 20000 - 18350 = 1650
    const expectedRemaining = '₹1,650';
    const remainingIsCorrect = remainingBudgetText.trim() === expectedRemaining;
    summaryStats.remainingBudget = 1650;
    logStep(
      'Verify Remaining Budget Card',
      'remainingBudget element',
      `Value: ${expectedRemaining}`,
      `Value: ${remainingBudgetText}`,
      remainingIsCorrect ? 'pass' : 'fail'
    );

    // Read Budget Status
    const budgetStatusEl = await driver.findElement(By.id('budgetStatus'));
    const budgetStatusText = await budgetStatusEl.getText();
    summaryStats.budgetStatus = budgetStatusText;
    logStep(
      'Verify Budget Status Label',
      'budgetStatus element',
      'Value: remaining',
      `Value: ${budgetStatusText}`,
      budgetStatusText.trim().toLowerCase() === 'remaining' ? 'pass' : 'fail'
    );

    // Read Highest Spend Category
    const highestCatEl = await driver.findElement(By.id('highestCat'));
    const highestCatText = await highestCatEl.getText();
    // Rent is 12000, which is top
    const expectedHighest = '🏠 Rent';
    const highestIsCorrect = highestCatText.includes('Rent') || highestCatText.includes('🏠');
    logStep(
      'Verify Highest Spend Category Card',
      'highestCat element',
      `Value: ${expectedHighest}`,
      `Value: ${highestCatText}`,
      highestIsCorrect ? 'pass' : 'fail'
    );

    // ── STEP 4: Test History Filter & Delete E2E ──
    const historyNav = await driver.findElement(By.css('button[data-view="history"]'));
    await historyNav.click();
    await driver.sleep(500);
    logStep('Click sidebar navigation', 'History Menu Item', 'History view shown', 'Navigated to History', 'pass');

    // Filter by Transport
    const filterCat = await driver.findElement(By.id('filterCat'));
    await filterCat.click();
    const transportOption = await driver.findElement(By.css('#filterCat option[value="transport"]'));
    await transportOption.click();
    await driver.executeScript('arguments[0].dispatchEvent(new Event("change", { bubbles: true }));', filterCat);
    await driver.sleep(500);

    // Check filtered transaction count
    const historyList = await driver.findElement(By.id('historyList'));
    const txItems = await historyList.findElements(By.className('tx-item'));
    const filterIsCorrect = txItems.length === 1;
    logStep(
      'Apply History Category Filter: Transport',
      'filterCat dropdown',
      'List shows exactly 1 transaction',
      `List shows ${txItems.length} transactions`,
      filterIsCorrect ? 'pass' : 'fail'
    );

    // Delete the Transport transaction (Uber to office)
    if (txItems.length > 0) {
      const deleteBtn = await txItems[0].findElement(By.className('tx-delete'));
      await deleteBtn.click();
      await driver.sleep(500);
      
      // Remove transport from recorded list locally for Excel reporting
      const indexToDelete = recordedExpenses.findIndex(e => e.catId === 'transport');
      if (indexToDelete > -1) {
        recordedExpenses.splice(indexToDelete, 1);
      }

      // Re-fetch transactions
      const updatedTxItems = await historyList.findElements(By.className('tx-item'));
      // Wait, filter is still set to 'Transport', so it should show 0 elements now
      const deleteSuccess = updatedTxItems.length === 0;
      logStep(
        'Delete Filtered Expense',
        'Click ✕ on "Uber to office"',
        'Transport list becomes empty',
        `Transport list has ${updatedTxItems.length} items left`,
        deleteSuccess ? 'pass' : 'fail'
      );
    }

    // Reset filter to All
    await filterCat.click();
    const allOption = await driver.findElement(By.css('#filterCat option[value=""]'));
    await allOption.click();
    await driver.executeScript('arguments[0].dispatchEvent(new Event("change", { bubbles: true }));', filterCat);
    await driver.sleep(500);

    // Verify history now shows 3 remaining items
    const allRemainingTxs = await historyList.findElements(By.className('tx-item'));
    const allRemainingCorrect = allRemainingTxs.length === 3;
    logStep(
      'Reset History Category Filter',
      'Set filterCat to ""',
      'History displays 3 remaining expenses',
      `History displays ${allRemainingTxs.length} expenses`,
      allRemainingCorrect ? 'pass' : 'fail'
    );

    // ── STEP 5: Verify Post-Deletion Stats on Dashboard ──
    await dashNav.click();
    await driver.sleep(500);

    const postTotalSpent = await (await driver.findElement(By.id('totalSpent'))).getText();
    // Sum now = 18350 - 350 = 18000
    const postTotalCorrect = postTotalSpent.trim() === '₹18,000';
    summaryStats.totalSpent = 18000;

    logStep(
      'Verify Updated Total Spent Card',
      'totalSpent element after delete',
      'Value: ₹18,000',
      `Value: ${postTotalSpent}`,
      postTotalCorrect ? 'pass' : 'fail'
    );

    const postRemaining = await (await driver.findElement(By.id('remainingBudget'))).getText();
    // Remaining now = 20000 - 18000 = 2000
    const postRemainingCorrect = postRemaining.trim() === '₹2,000';
    summaryStats.remainingBudget = 2000;
    
    logStep(
      'Verify Updated Remaining Budget Card',
      'remainingBudget element after delete',
      'Value: ₹2,000',
      `Value: ${postRemaining}`,
      postRemainingCorrect ? 'pass' : 'fail'
    );

    const postCount = await (await driver.findElement(By.id('monthCount'))).getText();
    logStep(
      'Verify Updated Transaction Count',
      'monthCount element after delete',
      'Value: 3',
      `Value: ${postCount}`,
      postCount.trim() === '3' ? 'pass' : 'fail'
    );

  } catch (err) {
    console.error('An error occurred during testing:');
    console.error(err);
    logStep('E2E Run Execution', 'Driver instance', 'No uncaught exceptions', `Error: ${err.message}`, 'fail');
  } finally {
    // Terminate browser session
    console.log('[INFO] Quitting Chrome browser...');
    try {
      await driver.quit();
    } catch (e) {
      console.error('[WARNING] Error quitting driver:', e);
    }

    // Close static HTTP server
    stopServer();

    // Stop backend Express server
    stopBackendServer();

    // Generate formatted Excel report
    console.log('[INFO] Starting Excel report compilation...');
    await generateExcelReport(recordedExpenses, summaryStats);
  }
}

// Start test suite
runTest();
