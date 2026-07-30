package com.pavithra.washingstation.repository;

import java.time.LocalDate;
import java.util.List;

import org.springframework.data.jpa.repository.JpaRepository;

import com.pavithra.washingstation.entity.Transaction;

public interface TransactionRepository extends JpaRepository<Transaction, Long> {

    List<Transaction> findByWorkerId(Long workerId);

    // NEW - Current month transactions
    List<Transaction> findByWorkerIdAndDateBetween(
            Long workerId,
            LocalDate startDate,
            LocalDate endDate
    );

    void deleteByWorkerId(Long workerId);
}