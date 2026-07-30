
package com.pavithra.washingstation.service;
import com.pavithra.washingstation.repository.TransactionRepository;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.pavithra.washingstation.entity.Worker;
import com.pavithra.washingstation.repository.WorkerRepository;
import org.springframework.transaction.annotation.Transactional;
@Service
public class WorkerService {

    @Autowired
    private WorkerRepository workerRepository;
    @Autowired
    private TransactionRepository transactionRepository;

    // Add Worker
    public Worker addWorker(Worker worker) {
        return workerRepository.save(worker);
    }

    // Get All Workers
    public List<Worker> getAllWorkers() {
        return workerRepository.findAll();
    }

    // Get Worker by ID
    public Worker getWorkerById(Long id) {
        return workerRepository.findById(id).orElse(null);
    }

    // Update Worker
    public Worker updateWorker(Long id, Worker updatedWorker) {

        Worker worker = workerRepository.findById(id).orElse(null);

        if (worker != null) {
            worker.setName(updatedWorker.getName());
            worker.setSalary(updatedWorker.getSalary());
            worker.setJoinDate(updatedWorker.getJoinDate());
            worker.setStatus(updatedWorker.getStatus());

            return workerRepository.save(worker);
        }

        return null;
    }

    // Delete Worker
 // Delete Worker
    @Transactional
    public void deleteWorker(Long id) {

        Worker worker = workerRepository.findById(id)
                .orElseThrow(() -> new RuntimeException("Worker not found"));

        transactionRepository.deleteByWorkerId(id);

        workerRepository.delete(worker);
    }
}