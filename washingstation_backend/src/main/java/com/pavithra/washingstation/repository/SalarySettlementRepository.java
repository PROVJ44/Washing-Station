package com.pavithra.washingstation.repository;

import java.util.List;
import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;

import com.pavithra.washingstation.entity.SalarySettlement;

public interface SalarySettlementRepository extends JpaRepository<SalarySettlement, Long> {

    List<SalarySettlement> findByWorkerId(Long workerId);

    Optional<SalarySettlement> findByWorkerIdAndMonthAndYear(
            Long workerId,
            Integer month,
            Integer year
    );
}