package com.pavithra.washingstation.service;

import java.time.LocalDate;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.pavithra.washingstation.entity.SalarySettlement;
import com.pavithra.washingstation.entity.Transaction;
import com.pavithra.washingstation.entity.Worker;
import com.pavithra.washingstation.repository.SalarySettlementRepository;
import com.pavithra.washingstation.repository.TransactionRepository;
import com.pavithra.washingstation.repository.WorkerRepository;

@Service
public class SalarySettlementService {

    @Autowired
    private SalarySettlementRepository settlementRepository;

    @Autowired
    private WorkerRepository workerRepository;

    @Autowired
    private TransactionRepository transactionRepository;

    public List<SalarySettlement> getAll() {
        return settlementRepository.findAll();
    }

    public SalarySettlement settleSalary(Long workerId) {

        Worker worker = workerRepository.findById(workerId).orElseThrow();

        int month = LocalDate.now().getMonthValue();
        int year = LocalDate.now().getYear();

        if (settlementRepository
                .findByWorkerIdAndMonthAndYear(workerId, month, year)
                .isPresent()) {

            throw new org.springframework.web.server.ResponseStatusException(
                    org.springframework.http.HttpStatus.CONFLICT,
                    "Salary already settled for this month."
            );
        }

        LocalDate today = LocalDate.now();

        LocalDate startOfMonth = today.withDayOfMonth(1);
        LocalDate endOfMonth = today.withDayOfMonth(today.lengthOfMonth());

        List<Transaction> transactions =
                transactionRepository.findByWorkerIdAndDateBetween(
                        workerId,
                        startOfMonth,
                        endOfMonth
                );

        double deductions = transactions.stream()
                .filter(t ->
                        t.getType().equals("Advance") ||
                        t.getType().equals("Food") ||
                        t.getType().equals("Rent") ||
                        t.getType().equals("Medical") ||
                        t.getType().equals("Bills"))
                .mapToDouble(Transaction::getAmount)
                .sum();

        SalarySettlement settlement = new SalarySettlement();

        settlement.setWorker(worker);
        settlement.setMonth(month);
        settlement.setYear(year);
        settlement.setTotalSalary(worker.getSalary());
        settlement.setDeductions(deductions);

        double remainingSalary = worker.getSalary() - deductions;

        settlement.setPaidSalary(remainingSalary);
        settlement.setSettlementDate(LocalDate.now());

        // Create final salary payment transaction
        Transaction salaryTransaction = new Transaction();
        salaryTransaction.setWorker(worker);
        salaryTransaction.setType("Salary Settlement");
        salaryTransaction.setAmount(remainingSalary);
        salaryTransaction.setDate(LocalDate.now());
        salaryTransaction.setRemarks("Monthly salary settled");

        System.out.println("==============");
        System.out.println("Saving transaction...");
        System.out.println("Worker : " + worker.getId());
        System.out.println("Amount : " + remainingSalary);

        transactionRepository.save(salaryTransaction);

        System.out.println("Transaction saved!");
        System.out.println("==============");

        return settlementRepository.save(settlement);
    }
}