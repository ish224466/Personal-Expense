package com.tanmay.expense_tracker.config;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.Collections;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.json.gson.GsonFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.auth.http.HttpCredentialsAdapter;
import com.google.auth.oauth2.GoogleCredentials;

@Configuration
public class GoogleSheetsConfig {

    @Bean
    public Sheets sheetsService() throws Exception {

        String json = System.getenv("GOOGLE_CREDENTIALS");

        if (json == null || json.isEmpty()) {
            throw new RuntimeException("GOOGLE_CREDENTIALS env variable not set");
        }

        GoogleCredentials credentials = GoogleCredentials
                .fromStream(new ByteArrayInputStream(json.getBytes(StandardCharsets.UTF_8)))
                .createScoped(Collections.singleton("https://www.googleapis.com/auth/spreadsheets"));

        return new Sheets.Builder(
                GoogleNetHttpTransport.newTrustedTransport(),
                GsonFactory.getDefaultInstance(),
                new HttpCredentialsAdapter(credentials)
        ).setApplicationName("Expense Tracker").build();
    }
}