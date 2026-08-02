package com.pavithra.washingstation.controller;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import com.pavithra.washingstation.entity.Transaction;
import com.pavithra.washingstation.service.TransactionService;

@RestController
@RequestMapping("/api/transactions")
@CrossOrigin(origins = {"http://localhost:5173", "https://washing-station-5a1g.vercel.app"})
public class TransactionController {

    @Autowired
    private TransactionService transactionService;

    @PostMapping
    public Transaction addTransaction(@RequestBody Transaction transaction) {
        return transactionService.addTransaction(transaction);
    }

    @GetMapping
    public List<Transaction> getAllTransactions() {
        return transactionService.getAllTransactions();
    }

    @GetMapping("/worker/{workerId}")
    public List<Transaction> getWorkerTransactions(@PathVariable Long workerId) {
        return transactionService.getWorkerTransactions(workerId);
    }
}