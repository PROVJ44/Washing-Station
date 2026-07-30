package com.pavithra.washingstation.controller;

import java.util.HashMap;
import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.pavithra.washingstation.entity.Transaction;
import com.pavithra.washingstation.entity.Worker;
import com.pavithra.washingstation.repository.TransactionRepository;
import com.pavithra.washingstation.repository.WorkerRepository;

@RestController
@RequestMapping("/api/dashboard")
@CrossOrigin(origins = "http://localhost:8443")
public class DashboardController {

    @Autowired
    private WorkerRepository workerRepository;

    @Autowired
    private TransactionRepository transactionRepository;

    @GetMapping
    public Map<String, Object> getDashboard() {

        Map<String, Object> data = new HashMap<>();

        // Total Workers
        long totalWorkers = workerRepository.count();

        // Active Workers
        long activeWorkers = workerRepository.findAll()
                .stream()
                .filter(worker -> "Active".equalsIgnoreCase(worker.getStatus()))
                .count();

        // Total Salary
        double totalSalary = workerRepository.findAll()
                .stream()
                .mapToDouble(Worker::getSalary)
                .sum();

        // Total Transactions Amount
        double totalTransactions = transactionRepository.findAll()
                .stream()
                .mapToDouble(Transaction::getAmount)
                .sum();

        data.put("totalWorkers", totalWorkers);
        data.put("activeWorkers", activeWorkers);
        data.put("totalSalary", totalSalary);
        data.put("totalTransactions", totalTransactions);

        return data;
    }
}