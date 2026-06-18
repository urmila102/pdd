# SpendSense Mobile E2E Testing Suite (Appium)

This folder contains the automated mobile end-to-end (E2E) testing suite for the SpendSense Android mobile application. The testing suite uses **Appium** with the **WebdriverIO** client in Node.js to launch an Android emulator, switch into the Capacitor application's **WEBVIEW** context, and perform complete user actions to generate a styled Excel report.

---

## What the Test Suite Does

1. **Launches Appium Session**: Connects to the local Appium server on `127.0.0.1:4723` and loads the SpendSense APK onto the target Android device/emulator.
2. **Context Switching**: Automatically switches context from the native Android wrapper (`NATIVE_APP`) to the hybrid web context (`WEBVIEW_com.spendsense.app`). This allows for highly reliable web CSS selector targeting.
3. **Database Reset**: Clears the application's local storage to start from a clean state.
4. **Sets Budget**: Navigate to the Budget tab, inputs a budget of `₹20,000`, and verifies correct submission feedback.
5. **Adds Expenses**: Inputs and records 4 distinct expenses (Food, Transport, Shopping, Rent).
6. **Verifies Dashboard Stats**: Navigates to the Dashboard and asserts that the Total Spent, Transaction Count, Remaining Budget, and Highest Category calculations are accurate.
7. **Filters History**: Navigates to the History tab, filters by category "Transport", and validates the record.
8. **Deletes Expense**: Deletes the Transport item and asserts that 3 records remain.
9. **Verifies Dashboard Re-calculations**: Validates the updated dashboard budget metrics.
10. **Generates Excel Analysis**: Creates the styled **`appium_report.xlsx`** spreadsheet within this directory.
11. **Cleans Up**: Deletes the mobile session and closes the connections.

---

## Prerequisites

To run these mobile tests, your system needs:

1. **Node.js**: Version 16+.
2. **Android Studio & SDK**:
   - Verify `ANDROID_HOME` environment variable points to your Android SDK location (e.g., `C:\Users\<username>\AppData\Local\Android\Sdk`).
   - Add Platform Tools to your Path.
3. **Android Emulator/Device**:
   - Set up an emulator running Android (API 30+ recommended).
   - Keep the emulator open and active before running tests.
4. **Appium Server**:
   - Install Appium globally: `npm install -g appium`
   - Install the UiAutomator2 driver: `appium driver install uiautomator2`
   - Start the Appium server using: `appium` (runs on default port `4723`)

---

## Build the App APK

Before running the Appium tests, you must compile the application's debug APK.

1. Install capacitor assets & sync the project (from root workspace directory):
   ```bash
   npm run sync
   ```
2. Build the Android application using the Gradle wrapper (from `android/` directory):
   ```bash
   cd android
   .\gradlew.bat assembleDebug
   ```
   *This builds the debug APK at `android/app/build/outputs/apk/debug/app-debug.apk` which is the path configured in the Appium capabilities.*

---

## Installation & Running

1. **Install Dependencies**:
   Open a terminal in this directory (`appium-tests/`) and run:
   ```bash
   npm install
   ```

2. **Ensure Emulator & Appium Server are running**:
   - Open your Android Emulator.
   - Run `appium` in a separate terminal.

3. **Run Mobile E2E Tests**:
   Execute the testing suite by running:
   ```bash
   npm test
   ```

---

## Outputs

- **Console logs**: Detailed real-time logs of every step, colored green (Pass) or red (Fail).
- **Excel Report**: Generates **`appium_report.xlsx`** inside this folder. The report is fully styled and features:
  - **Summary Dashboard**: Overview of test execution status, steps run, and overall financial outcome.
  - **Execution Logs**: Complete step-by-step table showing the action, expected outcome, actual outcome, and status of each step.
  - **Recorded Data**: Detailed table showing all transactions recorded during the test, including sums.
