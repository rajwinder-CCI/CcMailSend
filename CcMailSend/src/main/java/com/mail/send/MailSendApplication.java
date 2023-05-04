package com.mail.send;

import javax.mail.MessagingException;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.event.ApplicationReadyEvent;
import org.springframework.context.event.EventListener;

import com.mail.server.MailService;
import com.mail.server.MailServiceImp;

@SpringBootApplication
public class MailSendApplication implements CommandLineRunner {

	@Autowired
	private MailService mailService;

	public static void main(String[] args) {
		SpringApplication.run(com.mail.send.MailSendApplication.class, args);
	}

//	@EventListener({ ApplicationReadyEvent.class })
//	public void mail() throws MessagingException {
//		mailService.mailSend("hr@caeliusconsulting.com");
//		}
	
	@Override
	public void run(String... args) throws MessagingException {
		mailService.mailSend("hr@caeliusconsulting.com");
	}

//	@Override
//	@EventListener({ ApplicationReadyEvent.class })
//	public void run(String... args) throws MessagingException {
//		mailService.mailSend("hr@caeliusconsulting.com");
//}
	
}
