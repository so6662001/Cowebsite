package com.puxun.website.controller;

import com.puxun.website.model.ContactRequest;
import com.puxun.website.service.ContactService;
import jakarta.validation.Valid;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.Map;

@RestController
@RequestMapping("/api/contact")
public class ContactController {

    private final ContactService contactService;

    public ContactController(ContactService contactService) {
        this.contactService = contactService;
    }

    @PostMapping
    public ResponseEntity<Map<String, Object>> submitContact(@Valid @RequestBody ContactRequest request) {
        ContactRequest saved = contactService.submit(request);
        return ResponseEntity.ok(Map.of(
            "success", true,
            "message", "合作意向已提交，我们将尽快与您联系！",
            "id", saved.getId()
        ));
    }
}
