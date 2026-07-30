package com.pavithra.washingstation.repository;

import org.springframework.data.jpa.repository.JpaRepository;

import com.pavithra.washingstation.entity.Worker;

public interface WorkerRepository extends JpaRepository<Worker, Long> {
	
}