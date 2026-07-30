package com.pavithra.washingstation.service;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.pavithra.washingstation.entity.Earning;
import com.pavithra.washingstation.repository.EarningRepository;

@Service
public class EarningService {

    @Autowired
    private EarningRepository repository;
    public void deleteEarning(Long id) {
        repository.deleteById(id);
    }

    public List<Earning> getAll() {
        return repository.findAll();
    }
    public List<Earning> getPendingPayments() {
        return repository.findByPaymentMode("Pending");
    }
    public Earning updateEarning(Long id, Earning updatedEarning) {

        Earning earning = repository.findById(id).orElse(null);

        if (earning == null) {
            return null;
        }

        earning.setDate(updatedEarning.getDate());
        earning.setCustomerName(updatedEarning.getCustomerName());
        earning.setVehicleType(updatedEarning.getVehicleType());
        earning.setWashType(updatedEarning.getWashType());
        earning.setAmount(updatedEarning.getAmount());
        earning.setPaymentMode(updatedEarning.getPaymentMode());
        earning.setStatus(updatedEarning.getStatus());
        earning.setRemarks(updatedEarning.getRemarks());

        return repository.save(earning);
    }

    public Earning save(Earning earning) {
        return repository.save(earning);
    }
}