package com.pavithra.washingstation.repository;

import com.pavithra.washingstation.entity.Expense;
import org.springframework.data.jpa.repository.JpaRepository;

public interface ExpenseRepository extends JpaRepository<Expense, Long> {
}