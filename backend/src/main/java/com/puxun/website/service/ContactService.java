package com.puxun.website.service;

import com.puxun.website.model.ContactRequest;
import com.puxun.website.repository.ContactRequestRepository;
import org.springframework.stereotype.Service;

@Service
public class ContactService {

    private final ContactRequestRepository repository;

    public ContactService(ContactRequestRepository repository) {
        this.repository = repository;
    }

    public ContactRequest submit(ContactRequest request) {
        return repository.save(request);
    }
}
