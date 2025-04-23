package com.mailSender;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Properties;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.core.io.FileSystemResource;
import org.springframework.mail.javamail.JavaMailSenderImpl;
import org.springframework.mail.javamail.MimeMessageHelper;

import jakarta.mail.MessagingException;
import jakarta.mail.internet.MimeMessage;

@SpringBootApplication
public class MailSenderApplication {
	
	public static void main(String[] args) {
		SpringApplication.run(MailSenderApplication.class, args);
		try {
			SpreadsheetService sheet=new SpreadsheetService();
			String filePath="TrainingCenter.xlsx";
			List<Mail> data;
			if (filePath != null && !filePath.isEmpty()) {
			FileSystemResource file = new FileSystemResource(filePath);
			if (file.exists()) {
				data =sheet.readSpreadsheet(filePath);
			}
		}
			
		
		
		} catch (Exception e) {

	}
	}


}
