package com.tanmay.expense_tracker.controller;

import com.tanmay.expense_tracker.service.ExpenseService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import com.tanmay.expense_tracker.dto.ExpenseRequest;


@RestController
@RequestMapping("/api")
public class MainController {

    @Autowired
    private ExpenseService expenseService;

    @PostMapping("/add-expense")
    public String addExpense(@RequestBody ExpenseRequest req) {

        if (req.amount <= 0) {
            throw new RuntimeException("Invalid amount");
        }

        return expenseService.processExpense(req.amount, req.date);
    }
}
