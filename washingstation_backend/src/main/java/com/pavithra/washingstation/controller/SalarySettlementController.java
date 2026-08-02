package com.pavithra.washingstation.controller;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import com.pavithra.washingstation.entity.SalarySettlement;
import com.pavithra.washingstation.service.SalarySettlementService;

@RestController
@RequestMapping("/api/salary-settlements")
@CrossOrigin(origins = {"http://localhost:5173", "https://washing-station-5a1g.vercel.app"})
public class SalarySettlementController {

    @Autowired
    private SalarySettlementService service;

    @GetMapping
    public List<SalarySettlement> getAll() {
        return service.getAll();
    }

    @PostMapping("/{workerId}")
    public SalarySettlement settleSalary(@PathVariable Long workerId) {
        return service.settleSalary(workerId);
    }
}