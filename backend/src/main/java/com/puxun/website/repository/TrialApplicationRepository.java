package com.puxun.website.repository;

import com.puxun.website.model.TrialApplication;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface TrialApplicationRepository extends JpaRepository<TrialApplication, Long> {
    List<TrialApplication> findByStatusOrderByCreatedAtDesc(String status);
    boolean existsByPhone(String phone);
}
