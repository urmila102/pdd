const express = require('express');
const cors = require('cors');
const db = require('./database');

const app = express();
const PORT = process.env.PORT || 5000;

// Enable CORS and JSON body parser
app.use(cors());
app.use(express.json());

// Logging middleware
app.use((req, res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.url}`);
  next();
});

// --- API ROUTES ---

// GET: Fetch all expenses
app.get('/api/expenses', async (req, res) => {
  try {
    const expenses = await db.getExpenses();
    res.json(expenses);
  } catch (error) {
    res.status(500).json({ error: 'Failed to retrieve expenses' });
  }
});

// POST: Add new expense
app.post('/api/expenses', async (req, res) => {
  try {
    const { desc, amount, date, category, note } = req.body;
    
    if (!desc || !amount || !date || !category) {
      return res.status(400).json({ error: 'Missing required fields (desc, amount, date, category)' });
    }

    const newExpense = {
      id: Date.now(), // Generate unique ID
      desc,
      amount: parseFloat(amount),
      date,
      category,
      note: note || ''
    };

    const saved = await db.addExpense(newExpense);
    res.status(201).json(saved);
  } catch (error) {
    res.status(500).json({ error: 'Failed to record expense' });
  }
});

// DELETE: Delete a single expense
app.delete('/api/expenses/:id', async (req, res) => {
  try {
    const success = await db.deleteExpense(req.params.id);
    if (success) {
      res.json({ success: true, message: 'Expense deleted successfully' });
    } else {
      res.status(404).json({ error: 'Expense not found' });
    }
  } catch (error) {
    res.status(500).json({ error: 'Failed to delete expense' });
  }
});

// DELETE: Clear all expenses
app.delete('/api/expenses', async (req, res) => {
  try {
    await db.clearExpenses();
    res.json({ success: true, message: 'All expenses deleted' });
  } catch (error) {
    res.status(500).json({ error: 'Failed to clear expenses' });
  }
});

// GET: Fetch monthly budget
app.get('/api/budget', async (req, res) => {
  try {
    const budgetVal = await db.getBudget();
    res.json({ budget: budgetVal });
  } catch (error) {
    res.status(500).json({ error: 'Failed to retrieve budget' });
  }
});

// POST: Set monthly budget
app.post('/api/budget', async (req, res) => {
  try {
    const { budget } = req.body;
    if (budget === undefined || parseFloat(budget) < 0) {
      return res.status(400).json({ error: 'Invalid or missing budget value' });
    }

    const setVal = await db.setBudget(budget);
    res.json({ success: true, budget: setVal });
  } catch (error) {
    res.status(500).json({ error: 'Failed to configure budget' });
  }
});

// POST: Reset entire database (used in testing scripts)
app.post('/api/reset', async (req, res) => {
  try {
    await db.resetDb();
    res.json({ success: true, message: 'Database reset successfully' });
  } catch (error) {
    res.status(500).json({ error: 'Failed to reset database' });
  }
});

// Default Root endpoint
app.get('/', (req, res) => {
  res.json({ message: 'SpendSense API is online' });
});

// Start Server
app.listen(PORT, () => {
  console.log(`[INFO] SpendSense API Server running at http://localhost:${PORT}`);
});
