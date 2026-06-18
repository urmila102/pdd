# SpendSense Web E2E Testing Suite (Selenium)

This folder contains the automated end-to-end (E2E) testing suite for the SpendSense web application. The testing suite uses **Selenium WebDriver** in Node.js to spin up a headless Chrome browser, interact with the application, verify stats and core business logic, and generate a styled Excel report.

---

## What the Test Suite Does

1. **Launches a Local Server**: Spins up a local Node.js server hosting the static assets (`index.html`, `app.js`, `style.css`) on `http://localhost:8080`.
2. **Initializes Database**: Clears local storage to guarantee a clean state.
3. **Sets Budget**: Enters the Budget view, configures a monthly budget of `₹20,000`, and verifies the feedback message.
4. **Adds Expenses**: Records 4 transactions:
   - **Food**: *Dinner at Taj* — ₹1,500
   - **Transport**: *Uber to office* — ₹350
   - **Shopping**: *New headphones* — ₹4,500
   - **Rent**: *Monthly Rent* — ₹12,000
5. **Verifies Dashboard Stats**: Asserts that total spent equals `₹18,350`, remaining budget is `₹1,650`, transaction count is `4`, and the highest spend category is `🏠 Rent`.
6. **Filters History**: Navigates to the History view, filters by "Transport", and verifies only the Uber transaction is shown.
7. **Deletes Expense**: Deletes the Transport transaction, then clears the filter and asserts that 3 transactions remain.
8. **Verifies Dashboard After Deletion**: Returns to the dashboard to verify that stats updated correctly (Total Spent: `₹18,000`, Remaining Budget: `₹2,000`, Transaction Count: `3`).
9. **Compiles Report**: Generates `selenium_report.xlsx` inside this folder containing the execution logs and data summaries.
10. **Cleans Up**: Quits the browser and shuts down the test server.

---

## Prerequisites

1. **Node.js**: Make sure you have Node.js (v16+) installed.
2. **Google Chrome**: Google Chrome must be installed on your system. Selenium Manager will automatically download the correct ChromeDriver binary.

---

## Installation & Running

1. **Install Dependencies**:
   Open a terminal in this directory (`selenium-tests/`) and run:
   ```bash
   npm install
   ```

2. **Run E2E Tests**:
   Execute the testing suite by running:
   ```bash
   npm test
   ```
   *(Note: By default, it runs in **headless** mode so you won't see a browser popup. If you want to watch the browser actions, open `selenium-test.js` and comment out the `options.addArguments('--headless=new');` line.)*

---

## Outputs

- **Console logs**: Detailed real-time logs of every step, colored green (Pass) or red (Fail).
- **Excel Report**: Generates **`selenium_report.xlsx`** inside this folder. The report is fully styled and features:
  - **Summary Dashboard**: Overview of test execution status, steps run, and overall financial outcome.
  - **Execution Logs**: Complete step-by-step table showing the action, expected outcome, actual outcome, and status of each step.
  - **Recorded Data**: Detailed table showing all transactions recorded during the test, including sums.
