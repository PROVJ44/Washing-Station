package com.pavithra.washingstation.entity;

import java.time.LocalDate;

import jakarta.persistence.*;

@Entity
@Table(name = "salary_settlements")
public class SalarySettlement {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    private Integer month;

    private Integer year;

    private Double totalSalary;

    private Double deductions;

    private Double paidSalary;

    private LocalDate settlementDate;

    @ManyToOne
    @JoinColumn(name = "worker_id")
    private Worker worker;

    public SalarySettlement() {
    }

    public Long getId() {
        return id;
    }

    public Integer getMonth() {
        return month;
    }

    public void setMonth(Integer month) {
        this.month = month;
    }

    public Integer getYear() {
        return year;
    }

    public void setYear(Integer year) {
        this.year = year;
    }

    public Double getTotalSalary() {
        return totalSalary;
    }

    public void setTotalSalary(Double totalSalary) {
        this.totalSalary = totalSalary;
    }

    public Double getDeductions() {
        return deductions;
    }

    public void setDeductions(Double deductions) {
        this.deductions = deductions;
    }

    public Double getPaidSalary() {
        return paidSalary;
    }

    public void setPaidSalary(Double paidSalary) {
        this.paidSalary = paidSalary;
    }

    public LocalDate getSettlementDate() {
        return settlementDate;
    }

    public void setSettlementDate(LocalDate settlementDate) {
        this.settlementDate = settlementDate;
    }

    public Worker getWorker() {
        return worker;
    }

    public void setWorker(Worker worker) {
        this.worker = worker;
    }
}