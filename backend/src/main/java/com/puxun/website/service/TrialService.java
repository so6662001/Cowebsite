package com.puxun.website.service;

import com.puxun.website.model.TrialApplication;
import com.puxun.website.repository.TrialApplicationRepository;
import org.springframework.stereotype.Service;

@Service
public class TrialService {

    private final TrialApplicationRepository repository;

    public TrialService(TrialApplicationRepository repository) {
        this.repository = repository;
    }

    public TrialApplication submit(TrialApplication application) {
        return repository.save(application);
    }
}
