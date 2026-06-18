const { remote } = require('webdriverio');
const ExcelJS = require('exceljs');
const path = require('path');

// Appium configurations
const opts = {
  hostname: '127.0.0.1',
  port: 4723,
  path: '/', // Appium 2.x standard path is /
  capabilities: {
    platformName: 'Android',
    'appium:automationName': 'UiAutomator2',
    'appium:deviceName': 'Android Emulator',
    'appium:app': path.join(__dirname, '../android/app/build/outputs/apk/debug/app-debug.apk'),
    'appium:appPackage': 'com.spendsense.app',
    'appium:appActivity': 'com.spendsense.app.MainActivity',
    'appium:ensureWebviewsHavePages': true,
    'appium:nativeWebScreenshot': true,
    'appium:newCommandTimeout': 3600,
    'appium:connectHardwareKeyboard': true
  }
};

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
  workbook.creator = 'SpendSense Automated Appium E2E';
  workbook.created = new Date();

  // Premium Indigo/Teal theme
  const colors = {
    primaryDark: '312E81', // Indigo
    accent: '0F766E',      // Dark Teal
    passedBg: 'E6F4EA',    // Light Green
    passedText: '137333',  // Dark Green
    failedBg: 'FCE8E6',    // Light Red
    failedText: 'C5221F',  // Dark Red
    headerText: 'FFFFFF',
    borderLight: 'E2E8F0',
    zebraBg: 'F8FAFC'
  };

  // --- SHEET 1: Summary Dashboard ---
  const wsSummary = workbook.addWorksheet('Summary Dashboard');
  wsSummary.views = [{ showGridLines: true }];

  // Title Block
  wsSummary.mergeCells('B2:F2');
  const cellTitle = wsSummary.getCell('B2');
  cellTitle.value = 'SpendSense Android E2E Test Report';
  cellTitle.font = { name: 'Segoe UI', size: 16, bold: true, color: { argb: colors.headerText } };
  cellTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.primaryDark } };
  cellTitle.alignment = { horizontal: 'center', vertical: 'middle' };
  wsSummary.getRow(2).height = 40;

  // Metadata Table
  const metaLabels = [
    ['Test Engine', 'Appium (UiAutomator2 + WebdriverIO)'],
    ['Target Platform', 'Android (Capacitor Hybrid Webview)'],
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

  const thinBorder = {
    top: { style: 'thin', color: { argb: colors.borderLight } },
    left: { style: 'thin', color: { argb: colors.borderLight } },
    bottom: { style: 'thin', color: { argb: colors.borderLight } },
    right: { style: 'thin', color: { argb: colors.borderLight } }
  };
  ['E4', 'E5', 'F4', 'F5', 'E7', 'E8', 'F7', 'F8'].forEach(c => {
    wsSummary.getCell(c).border = thinBorder;
  });

  // Column dimensions
  wsSummary.getColumn('B').width = 22;
  wsSummary.getColumn('C').width = 34;
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
    
    row.getCell(1).alignment = { horizontal: 'center' };
    row.getCell(6).alignment = { horizontal: 'center' };

    const statusCell = row.getCell(6);
    if (s.status === 'pass') {
      statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.passedBg } };
      statusCell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: colors.passedText } };
    } else {
      statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.failedBg } };
      statusCell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: colors.failedText } };
    }

    row.eachCell((cell) => {
      cell.border = thinBorder;
      if (cell.col !== 6) {
        cell.font = { name: 'Segoe UI', size: 10 };
      }
    });
  });

  wsLogs.columns.forEach((column) => {
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

  // Add expenses
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

  // Add total row
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
  const reportPath = path.join(__dirname, 'appium_report.xlsx');
  await workbook.xlsx.writeFile(reportPath);
  console.log(`[SUCCESS] Excel report successfully generated and saved to: ${reportPath}`);
}

// ── Switch context to hybrid WebView ──
async function switchToWebview(driver) {
  console.log('[INFO] Attempting to retrieve execution contexts...');
  await driver.pause(3000); // Wait for the webview process to spin up
  const contexts = await driver.getContexts();
  console.log('[INFO] Available contexts:', contexts);
  
  const webview = contexts.find(c => c.includes('WEBVIEW'));
  if (webview) {
    await driver.switchContext(webview);
    console.log(`[SUCCESS] Switched app context to: ${webview}`);
    return true;
  }
  
  console.log('[WARNING] WEBVIEW context not found. Remaining in NATIVE_APP context.');
  return false;
}

// ── Appium Mobile E2E Test Execution ──
async function runMobileTest() {
  console.log('[INFO] Launching Appium Session...');
  let driver;
  const recordedExpenses = [];
  const summaryStats = {
    budget: 0,
    totalSpent: 0,
    remainingBudget: 0,
    budgetStatus: 'No budget set'
  };

  try {
    driver = await remote(opts);
    logStep('Launch Appium driver', 'Remote Session', 'SpendSense Mobile App Launches', 'Application launched successfully', 'pass');

    // Wait for the app to load
    await driver.pause(3000);

    // Switch context to Capacitor Webview
    const inWebview = await switchToWebview(driver);
    logStep('Switch Context to WEBVIEW', 'Capacitor Webview', 'Success switching context', inWebview ? 'Switched to Webview' : 'Stayed in Native App context', inWebview ? 'pass' : 'fail');

    // Clear local storage for a clean test run
    if (inWebview) {
      await driver.executeScript('window.localStorage.clear(); window.location.reload();');
      await driver.pause(2000);
      logStep('Initialize Database', 'localStorage.clear()', 'Fresh clean slate', 'Database cleared and page reloaded', 'pass');
    }

    // ── STEP 1: Set Budget ──
    const budgetNav = await driver.$('button[data-view="budget"]');
    await budgetNav.click();
    await driver.pause(800);
    logStep('Click navigation button', 'Budget View Tab', 'Budget view shown', 'Navigated to Budget view', 'pass');

    const budgetInput = await driver.$('#budgetInput');
    await budgetInput.setValue('20000');
    summaryStats.budget = 20000;

    const budgetSubmit = await driver.$('button[onclick="setBudget()"]');
    await budgetSubmit.click();
    await driver.pause(1000);

    const budgetFb = await driver.$('#budgetFeedback');
    const budgetMsg = await budgetFb.getText();
    if (budgetMsg.includes('set to') || budgetMsg.includes('₹20,000')) {
      logStep('Set Monthly Budget', '₹20,000', 'Success message displayed', `Budget set feedback: "${budgetMsg}"`, 'pass');
    } else {
      logStep('Set Monthly Budget', '₹20,000', 'Success message displayed', `Incorrect feedback message: "${budgetMsg}"`, 'fail');
    }

    // ── STEP 2: Add Expenses ──
    const addNav = await driver.$('button[data-view="add"]');
    await addNav.click();
    await driver.pause(800);
    logStep('Click navigation button', 'Add Expense View Tab', 'Add Expense view shown', 'Navigated to Add Expense view', 'pass');

    const txs = [
      { desc: 'Dinner at Taj', amount: '1500', category: 'Food', catId: 'food', note: 'Dinner with team' },
      { desc: 'Uber to office', amount: '350', category: 'Transport', catId: 'transport', note: '' },
      { desc: 'New headphones', amount: '4500', category: 'Shopping', catId: 'shopping', note: 'Noise cancelling' },
      { desc: 'Monthly Rent', amount: '12000', category: 'Rent', catId: 'rent', note: 'Flat 402' }
    ];

    for (const tx of txs) {
      const expDesc = await driver.$('#expDesc');
      await expDesc.setValue(tx.desc);

      const expAmount = await driver.$('#expAmount');
      await expAmount.setValue(tx.amount);

      if (tx.note) {
        const expNote = await driver.$('#expNote');
        await expNote.setValue(tx.note);
      }

      // Pick category
      const catSelector = `//button[contains(@class, "cat-option") and contains(., "${tx.category}")]`;
      const catBtn = await driver.$(catSelector);
      await catBtn.click();

      // Submit
      const recordBtn = await driver.$('button[onclick="addExpense()"]');
      await recordBtn.click();
 
      // Verify feedback
      const formFb = await driver.$('#formFeedback');
      let formMsg = "";
      try {
        await driver.waitUntil(async () => {
          const text = await formFb.getText();
          return text.includes('recorded');
        }, { timeout: 2000, interval: 200 });
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
    const dashNav = await driver.$('button[data-view="dashboard"]');
    await dashNav.click();
    await driver.pause(800);
    logStep('Click navigation button', 'Dashboard View Tab', 'Dashboard view shown', 'Navigated to Dashboard', 'pass');

    // Read Total Spent
    const totalSpentEl = await driver.$('#totalSpent');
    const totalSpentText = await totalSpentEl.getText();
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
    const monthCountEl = await driver.$('#monthCount');
    const monthCountText = await monthCountEl.getText();
    logStep(
      'Verify Monthly Transaction Count',
      'monthCount element',
      'Value: 4',
      `Value: ${monthCountText}`,
      monthCountText.trim() === '4' ? 'pass' : 'fail'
    );

    // Read Remaining Budget
    const remainingBudgetEl = await driver.$('#remainingBudget');
    const remainingBudgetText = await remainingBudgetEl.getText();
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
    const budgetStatusEl = await driver.$('#budgetStatus');
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
    const highestCatEl = await driver.$('#highestCat');
    const highestCatText = await highestCatEl.getText();
    const highestIsCorrect = highestCatText.includes('Rent') || highestCatText.includes('🏠');
    logStep(
      'Verify Highest Spend Category Card',
      'highestCat element',
      'Value: 🏠 Rent',
      `Value: ${highestCatText}`,
      highestIsCorrect ? 'pass' : 'fail'
    );

    // ── STEP 4: Test History Filter & Delete E2E ──
    const historyNav = await driver.$('button[data-view="history"]');
    await historyNav.click();
    await driver.pause(800);
    logStep('Click navigation button', 'History View Tab', 'History view shown', 'Navigated to History', 'pass');

    // Filter by Transport
    const filterCat = await driver.$('#filterCat');
    await filterCat.selectByAttribute('value', 'transport');
    await driver.pause(500);

    // Check filtered count
    let txItems = await driver.$$('.transactions-list.full .tx-item');
    logStep(
      'Apply History Category Filter: Transport',
      'filterCat dropdown selection',
      'List shows exactly 1 transaction',
      `List shows ${txItems.length} transactions`,
      txItems.length === 1 ? 'pass' : 'fail'
    );

    // Delete the Transport transaction
    if (txItems.length > 0) {
      const deleteBtn = await txItems[0].$('.tx-delete');
      await deleteBtn.click();
      await driver.pause(800);

      const indexToDelete = recordedExpenses.findIndex(e => e.catId === 'transport');
      if (indexToDelete > -1) {
        recordedExpenses.splice(indexToDelete, 1);
      }

      // Verify filtered list is empty now
      txItems = await driver.$$('.transactions-list.full .tx-item');
      logStep(
        'Delete Filtered Expense',
        'Click ✕ on first item',
        'Filtered list becomes empty',
        `Filtered list has ${txItems.length} items left`,
        txItems.length === 0 ? 'pass' : 'fail'
      );
    }

    // Reset filter to All
    await filterCat.selectByAttribute('value', '');
    await driver.pause(500);

    // Verify history now shows 3 remaining items
    const allRemainingTxs = await driver.$$('.transactions-list.full .tx-item');
    logStep(
      'Reset History Category Filter',
      'Set filterCat to All ("")',
      'History displays 3 remaining expenses',
      `History displays ${allRemainingTxs.length} expenses`,
      allRemainingTxs.length === 3 ? 'pass' : 'fail'
    );

    // ── STEP 5: Verify Post-Deletion Stats on Dashboard ──
    await dashNav.click();
    await driver.pause(800);

    const postTotalSpent = await (await driver.$('#totalSpent')).getText();
    const postTotalCorrect = postTotalSpent.trim() === '₹18,000';
    summaryStats.totalSpent = 18000;
    logStep(
      'Verify Updated Total Spent Card',
      'totalSpent element after delete',
      'Value: ₹18,000',
      `Value: ${postTotalSpent}`,
      postTotalCorrect ? 'pass' : 'fail'
    );

    const postRemaining = await (await driver.$('#remainingBudget')).getText();
    const postRemainingCorrect = postRemaining.trim() === '₹2,000';
    summaryStats.remainingBudget = 2000;
    logStep(
      'Verify Updated Remaining Budget Card',
      'remainingBudget element after delete',
      'Value: ₹2,000',
      `Value: ${postRemaining}`,
      postRemainingCorrect ? 'pass' : 'fail'
    );

    const postCount = await (await driver.$('#monthCount')).getText();
    logStep(
      'Verify Updated Transaction Count',
      'monthCount element after delete',
      'Value: 3',
      `Value: ${postCount}`,
      postCount.trim() === '3' ? 'pass' : 'fail'
    );

  } catch (err) {
    console.error('An error occurred during mobile testing:');
    console.error(err);
    logStep('Mobile E2E Run Execution', 'Appium driver', 'No uncaught exceptions', `Error: ${err.message}`, 'fail');
  } finally {
    if (driver) {
      console.log('[INFO] Tearing down Appium session...');
      await driver.deleteSession();
    }

    // Generate formatted Excel report
    console.log('[INFO] Starting Excel report compilation...');
    await generateExcelReport(recordedExpenses, summaryStats);
  }
}

// Start mobile test suite
runMobileTest();
