package com.pavithra.washingstation.repository;

import java.util.List;

import org.springframework.data.jpa.repository.JpaRepository;

import com.pavithra.washingstation.entity.Earning;

public interface EarningRepository extends JpaRepository<Earning, Long> {

    List<Earning> findByPaymentMode(String paymentMode);

}