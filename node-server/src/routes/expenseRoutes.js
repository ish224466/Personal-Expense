import express from "express";
import { addExpense, EXPENSE_CATEGORIES } from "../services/sheetService.js";

const router = express.Router();

router.post("/add-expense", async (req, res) => {
  try {
    const { amount, date, category, notes } = req.body;

    if (!amount || amount <= 0) {
      return res.status(400).json({ message: "Invalid amount" });
    }

    if (!date) {
      return res.status(400).json({ message: "Date required" });
    }

    if (!category || !EXPENSE_CATEGORIES.includes(category)) {
      return res.status(400).json({ message: "Invalid category" });
    }

    const result = await addExpense(amount, date, category, notes);

    res.json({ message: result });
  } catch (err) {
    console.error("Failed to save expense:", err?.response?.data || err);
    res.status(500).json({ error: err.message });
  }
});

import { getMonthlyExpenses } from "../services/sheetService.js";

router.get("/month-expenses", async (req, res) => {
  try {
    const data = await getMonthlyExpenses(req.query.month);
    res.json(data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

export default router;