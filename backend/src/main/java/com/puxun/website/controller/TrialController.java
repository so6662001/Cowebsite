package com.puxun.website.controller;

import com.puxun.website.model.TrialApplication;
import com.puxun.website.service.TrialService;
import jakarta.validation.Valid;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.Map;

@RestController
@RequestMapping("/api/trial")
public class TrialController {

    private final TrialService trialService;

    public TrialController(TrialService trialService) {
        this.trialService = trialService;
    }

    @PostMapping
    public ResponseEntity<Map<String, Object>> submitTrial(@Valid @RequestBody TrialApplication application) {
        TrialApplication saved = trialService.submit(application);
        return ResponseEntity.ok(Map.of(
            "success", true,
            "message", "试用申请提交成功，我们将在24小时内与您联系！",
            "id", saved.getId()
        ));
    }
}
