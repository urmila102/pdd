const fs = require('fs').promises;
const path = require('path');

const DB_FILE = path.join(__dirname, 'db.json');

// Helper to ensure database JSON file exists
async function initDb() {
  try {
    await fs.access(DB_FILE);
  } catch (error) {
    // File doesn't exist, create it with default structure
    const defaultData = {
      expenses: [],
      budget: 0
    };
    await fs.writeFile(DB_FILE, JSON.stringify(defaultData, null, 2), 'utf8');
  }
}

// Read database contents
async function readDb() {
  await initDb();
  try {
    const content = await fs.readFile(DB_FILE, 'utf8');
    return JSON.parse(content);
  } catch (error) {
    console.error('Error reading DB file, returning empty state:', error);
    return { expenses: [], budget: 0 };
  }
}

// Write database contents
async function writeDb(data) {
  await initDb();
  await fs.writeFile(DB_FILE, JSON.stringify(data, null, 2), 'utf8');
}

// --- Data Operations API ---

async function getExpenses() {
  const db = await readDb();
  return db.expenses || [];
}

async function addExpense(expense) {
  const db = await readDb();
  if (!db.expenses) db.expenses = [];
  
  // unshift to match frontend behavior (latest first)
  db.expenses.unshift(expense);
  await writeDb(db);
  return expense;
}

async function deleteExpense(id) {
  const db = await readDb();
  if (!db.expenses) db.expenses = [];
  
  const originalLength = db.expenses.length;
  db.expenses = db.expenses.filter(e => e.id !== parseInt(id) && e.id !== id);
  
  const deleted = db.expenses.length < originalLength;
  if (deleted) {
    await writeDb(db);
  }
  return deleted;
}

async function clearExpenses() {
  const db = await readDb();
  db.expenses = [];
  await writeDb(db);
  return true;
}

async function getBudget() {
  const db = await readDb();
  return db.budget || 0;
}

async function setBudget(budgetVal) {
  const db = await readDb();
  db.budget = parseFloat(budgetVal) || 0;
  await writeDb(db);
  return db.budget;
}

module.exports = {
  getExpenses,
  addExpense,
  deleteExpense,
  clearExpenses,
  getBudget,
  setBudget,
  // Helper to reset entire database
  resetDb: async () => {
    await writeDb({ expenses: [], budget: 0 });
  }
};
