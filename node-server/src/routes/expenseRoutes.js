import express from "express";
import { addExpense } from "../services/sheetService.js";

const router = express.Router();

router.post("/add-expense", async (req, res) => {
  try {
    const { amount, date } = req.body;

    if (!amount || amount <= 0) {
      return res.status(400).json({ message: "Invalid amount" });
    }

    if (!date) {
      return res.status(400).json({ message: "Date required" });
    }

    const result = await addExpense(amount, date);

    res.json({ message: result });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

import { getMonthlyExpenses } from "../services/sheetService.js";

router.get("/month-expenses", async (req, res) => {
  try {
    const data = await getMonthlyExpenses();
    res.json(data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

export default router;