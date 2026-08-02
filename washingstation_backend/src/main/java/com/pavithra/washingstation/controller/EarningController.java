package com.pavithra.washingstation.controller;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import com.pavithra.washingstation.entity.Earning;
import com.pavithra.washingstation.service.EarningService;

@RestController
@RequestMapping("/api/earnings")
@CrossOrigin(origins = {"http://localhost:5173", "https://washing-station-5a1g.vercel.app"})
public class EarningController {

    @Autowired
    private EarningService service;

    @GetMapping
    public List<Earning> getAll() {
        return service.getAll();
    }
    @GetMapping("/pending")
    public List<Earning> getPendingPayments() {
        return service.getPendingPayments();
    }
    @PutMapping("/{id}")
    public Earning updateEarning(@PathVariable Long id,
                                 @RequestBody Earning earning) {
        return service.updateEarning(id, earning);
    }
    @DeleteMapping("/{id}")
    public void deleteEarning(@PathVariable Long id) {
        service.deleteEarning(id);
    }

    @PostMapping
    public Earning save(@RequestBody Earning earning) {

        System.out.println("Date = " + earning.getDate());
        System.out.println("Customer = " + earning.getCustomerName());
        System.out.println("Vehicle = " + earning.getVehicleType());
        System.out.println("Amount = " + earning.getAmount());

        return service.save(earning);
    }
}