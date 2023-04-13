package com.mail.server;

import javax.mail.MessagingException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.EnableAutoConfiguration;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.event.ApplicationReadyEvent;
import org.springframework.context.event.EventListener;

@SpringBootApplication
@EnableAutoConfiguration
public class MailServiceImp {
	
	@Autowired
	private MailService mailService;
	  
	  public static void main(String[] args) {
	    SpringApplication.run(com.mail.send.MailSendApplication.class, args);
	  }
	  
	  @EventListener({ApplicationReadyEvent.class})
	  public void mail() throws MessagingException {
	    this.mailService.mailSend("rajwinder.kaur@caeliusconsulting.com");
	  }

}
