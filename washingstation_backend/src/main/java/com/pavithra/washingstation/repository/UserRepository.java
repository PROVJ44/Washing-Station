package com.pavithra.washingstation.repository;

import java.util.Optional;

import org.springframework.data.jpa.repository.JpaRepository;

import com.pavithra.washingstation.entity.user;

public interface UserRepository extends JpaRepository<user, Long> {

    Optional<user> findByUsername(String username);

}