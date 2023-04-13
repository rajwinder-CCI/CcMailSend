package com.mail.send;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import com.mail.server.MailService;
import com.mail.server.MailServiceImp;

@Configuration
public class Config {
	@Bean
	  public MailService mailService() {
	    return (MailService)new MailServiceImp();
	  }

}
