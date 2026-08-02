package com.pavithra.washingstation.controller;

import java.util.Map;
import java.util.Optional;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import com.pavithra.washingstation.entity.user;
import com.pavithra.washingstation.repository.UserRepository;

@RestController
@RequestMapping("/api/auth")
@CrossOrigin(origins = {"http://localhost:5173", "https://washing-station-5a1g.vercel.app"})
public class AuthController {

    @Autowired
    private UserRepository userRepository;

    @PostMapping("/login")
    public String login(@RequestBody Map<String, String> request) {

        String username = request.get("username");
        String password = request.get("password");

        Optional<user> user = userRepository.findByUsername(username);

        if (user.isPresent() && user.get().getPassword().equals(password)) {
            return "SUCCESS";
        }

        return "INVALID";
    }
    @PutMapping("/change-password")
    public String changePassword(@RequestBody Map<String, String> request) {

        String username = request.get("username");
        String currentPassword = request.get("currentPassword");
        String newPassword = request.get("newPassword");

        Optional<user> user = userRepository.findByUsername(username);

        if (user.isEmpty()) {
            return "USER_NOT_FOUND";
        }

        if (!user.get().getPassword().equals(currentPassword)) {
            return "INVALID_PASSWORD";
        }

        user u = user.get();
        u.setPassword(newPassword);

        userRepository.save(u);

        return "SUCCESS";
    }
}