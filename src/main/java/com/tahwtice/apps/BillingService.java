package com.tahwtice.apps;

import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

import com.tahwtice.apps.model.Billing;

import lombok.Data;

@Data
public class BillingService {
    private List<Billing> billingList;

    public Optional<Billing> findBillingByGuid(String guid) {
        return this.billingList.stream().filter(billing -> billing.getGuid().equals(guid)).collect(Collectors.toList())
                .stream().findAny();
    }
}
