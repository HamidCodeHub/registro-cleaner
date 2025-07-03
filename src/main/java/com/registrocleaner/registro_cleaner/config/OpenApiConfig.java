package com.registrocleaner.registro_cleaner.config;

import io.swagger.v3.oas.models.OpenAPI;
import io.swagger.v3.oas.models.info.Contact;
import io.swagger.v3.oas.models.info.Info;
import io.swagger.v3.oas.models.info.License;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class OpenApiConfig {

    @Bean
    public OpenAPI customOpenAPI() {
        return new OpenAPI()
                .info(new Info()
                        .title("Investor Bulk API")
                        .description("API for bulk insertion and management of investors")
                        .version("1.0.0")
                        .contact(new Contact()
                                .name("Ucapital")
                                .email("info@ucapital.com")
                                .url("https://ucapital.com"))
                        .license(new License()
                                .name("Proprietary")
                                .url("https://ucapital.com/license")));
    }
}