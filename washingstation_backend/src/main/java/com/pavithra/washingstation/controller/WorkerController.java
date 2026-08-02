package com.pavithra.washingstation.controller;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import com.pavithra.washingstation.entity.Worker;
import com.pavithra.washingstation.service.WorkerService;
import org.springframework.http.ResponseEntity;
@RestController
@RequestMapping("/api/workers")
@CrossOrigin(origins = {"http://localhost:8443", "https://washing-station-5a1g.vercel.app"})
public class WorkerController {

    @Autowired
    private WorkerService workerService;

    // Get all workers
    @GetMapping
    public List<Worker> getAllWorkers() {
        return workerService.getAllWorkers();
    }
    @DeleteMapping("/{id}")
    public ResponseEntity<String> deleteWorker(@PathVariable Long id) {

        workerService.deleteWorker(id);

        return ResponseEntity.ok("Worker deleted successfully");
    }

    // Get worker by ID
    @GetMapping("/{id}")
    public Worker getWorkerById(@PathVariable Long id) {
        return workerService.getWorkerById(id);
    }

    // Add worker
    @PostMapping
    public Worker addWorker(@RequestBody Worker worker) {
        return workerService.addWorker(worker);
    }

    // Update worker
    @PutMapping("/{id}")
    public Worker updateWorker(@PathVariable Long id,
                               @RequestBody Worker updatedWorker) {

        return workerService.updateWorker(id, updatedWorker);
    }

}