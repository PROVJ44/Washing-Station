package com.pavithra.washingstation.service;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.pavithra.washingstation.entity.Transaction;
import com.pavithra.washingstation.repository.TransactionRepository;

@Service
public class TransactionService {

    @Autowired
    private TransactionRepository transactionRepository;

    public Transaction addTransaction(Transaction transaction) {
        return transactionRepository.save(transaction);
    }

    public List<Transaction> getAllTransactions() {
        return transactionRepository.findAll();
    }

    public List<Transaction> getWorkerTransactions(Long workerId) {
        return transactionRepository.findByWorkerId(workerId);
    }
    public double getTotalDeductions(Long workerId) {

        List<Transaction> transactions = transactionRepository.findByWorkerId(workerId);

        return transactions.stream()
                .filter(t ->
                        t.getType().equals("Advance") ||
                        t.getType().equals("Food") ||
                        t.getType().equals("Rent") ||
                        t.getType().equals("Medical") ||
                        t.getType().equals("Bills"))
                .mapToDouble(Transaction::getAmount)
                .sum();
    }
}