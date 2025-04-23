package com.mailSender;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.mail.javamail.JavaMailSenderImpl;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;

import jakarta.mail.internet.MimeMessage;

public class SpreadsheetService {

//	public List<Mail> readSpreadsheet(String filePath) throws IOException {
//		List<Mail> persons = new ArrayList<>();
//		String fileName = "MailStatus.xlsx";
//		 File file = new File(fileName);
//		 // Check if the file exists and delete it
//	        if (file.exists()) {
//	            if (file.delete()) {
//	                System.out.println("Existing file deleted: " + fileName);
//	            } else {
//	                System.err.println("Failed to delete the existing file: " + fileName);
//	                
//	            }
//	        }
//		// Create a workbook and sheet
//		Workbook workbookStatus = new XSSFWorkbook();
//		Sheet StatusSheet = workbookStatus.createSheet("User Data");
//		// Create the header row
//		Row headerRow = StatusSheet.createRow(0); // Row 0 is the header row
//		// Define headers
//		String[] headers = { "Name", "Email", "Status" };
//		// Write headers to the row
//		for (int i = 0; i < headers.length; i++) {
//			Cell cell = headerRow.createCell(i);
//			cell.setCellValue(headers[i]);
//			System.out.println(headers[i]);
//			// Apply styling to header cells
//			CellStyle headerStyle = workbookStatus.createCellStyle();
//			Font headerFont = workbookStatus.createFont();
//			headerFont.setBold(true);
//			headerFont.setFontHeightInPoints((short) 12);
//			headerStyle.setFont(headerFont);
//			cell.setCellStyle(headerStyle);
//		}
//
//		try (FileInputStream fis = new FileInputStream(new File(filePath));
//				Workbook workbook = WorkbookFactory.create(fis)) {
//
//			Sheet sheet = workbook.getSheetAt(0); // Read the first sheet
//
//			for (Row row : sheet) {
//				if (row.getRowNum() == 0)
//					continue; // Skip header row
//				String name = row.getCell(0).getStringCellValue();
//				String email = row.getCell(1).getStringCellValue();
//				String status =this.sendMail(name, email);
//
//
//				persons.add(new Mail(name, email, status));
//			}
//
//			// Write data from the List<Mail> to the sheet
//			int rowCount = 1; // Start after header
//			for (Mail mail : persons) {
//				Row row = sheet.createRow(rowCount++);
//				row.createCell(0).setCellValue(mail.getName());
//				row.createCell(1).setCellValue(mail.getEmail());
//				row.createCell(2).setCellValue(mail.getStatus());
//			}
//			// Auto size the columns
//			for (int j = 0; j < headers.length; j++) {
//				sheet.autoSizeColumn(j);
//			}
//			// Write the output to a file
//			try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
//				workbook.write(fileOut);
//				System.out.println("Excel file created successfully: " + fileName);
//			} catch (IOException e) {
//				System.err.println("Error writing to file: " + e.getMessage());
//			} finally {
//				try {
//					workbook.close();
//				} catch (IOException e) {
//					e.printStackTrace();
//				}
//			}
//
//		}
//		return persons;
//	}
	public List<Mail> readSpreadsheet(String filePath) throws IOException {
	    List<Mail> persons = new ArrayList<>();
	    String fileName = "MailStatus.xlsx";
	    File file = new File(fileName);

	    // Check if the file exists and delete it
	    if (file.exists()) {
	        if (file.delete()) {
	            System.out.println("Existing file deleted: " + fileName);
	        } else {
	            System.err.println("Failed to delete the existing file: " + fileName);
	        }
	    }

	    // Create a workbook and sheet for output
	    Workbook workbookStatus = new XSSFWorkbook();
	    Sheet statusSheet = workbookStatus.createSheet("User Data");

	    // Create the header row
	    Row headerRow = statusSheet.createRow(0);
	    String[] headers = {"Name", "Email", "Status"};

	    // Write headers to the row
	    for (int i = 0; i < headers.length; i++) {
	        Cell cell = headerRow.createCell(i);
	        cell.setCellValue(headers[i]);
	        CellStyle headerStyle = workbookStatus.createCellStyle();
	        Font headerFont = workbookStatus.createFont();
	        headerFont.setBold(true);
	        headerFont.setFontHeightInPoints((short) 12);
	        headerStyle.setFont(headerFont);
	        cell.setCellStyle(headerStyle);
	    }

	    // Read the input file
	    try (FileInputStream fis = new FileInputStream(new File(filePath));
	         Workbook workbook = WorkbookFactory.create(fis)) {

	        Sheet inputSheet = workbook.getSheetAt(0); // Read the first sheet

	        for (Row row : inputSheet) {
	            if (row.getRowNum() == 0) continue; // Skip header row

	            String name = row.getCell(0).getStringCellValue();
	            String email = row.getCell(1).getStringCellValue();
	            String status = this.sendMail(name, email); // Assuming sendMail returns a status

	            persons.add(new Mail(name, email, status));
	        }
	    }

	    // Write data to the new sheet
	    int rowCount = 1; // Start after header
	    for (Mail mail : persons) {
	        Row row = statusSheet.createRow(rowCount++);
	        row.createCell(0).setCellValue(mail.getName());
	        row.createCell(1).setCellValue(mail.getEmail());
	        row.createCell(2).setCellValue(mail.getStatus());
	    }

	    // Auto size the columns
	    for (int j = 0; j < headers.length; j++) {
	        statusSheet.autoSizeColumn(j);
	    }

	    // Write the output to a file
	    try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
	        workbookStatus.write(fileOut);
	        System.out.println("Excel file created successfully: " + fileName);
	    } catch (IOException e) {
	        System.err.println("Error writing to file: " + e.getMessage());
	    } finally {
	        workbookStatus.close();
	    }

	    return persons;
	}
	public String sendMail(String recipientName, String ReciverMailId) {

		JavaMailSenderImpl mailSender = new JavaMailSenderImpl();
		mailSender.setHost("smtp.hostinger.com"); // Set host
		mailSender.setPort(587); // Set port
		mailSender.setUsername("info@vsmartengine.com"); // Set username
		mailSender.setPassword("$Meganar123"); // Set password

		// Email properties
		Properties props = mailSender.getJavaMailProperties();
		props.put("mail.transport.protocol", "smtp");
		props.put("mail.smtp.auth", "true");
		props.put("mail.smtp.starttls.enable", "true");
		props.put("mail.debug", "true");
		try {
			MimeMessage mimeMessage = mailSender.createMimeMessage();
			MimeMessageHelper helper = new MimeMessageHelper(mimeMessage, true, "UTF-8");

			// Set email metadata
			helper.setFrom("info@vsmartengine.com");
			helper.setTo(ReciverMailId);
			helper.setSubject("LearnHub");
			StringBuilder htmlContent2 = new StringBuilder();
			htmlContent2.append("<!DOCTYPE html>").append("<html lang=\"en\">").append("<head>")
					.append("<meta charset=\"UTF-8\">")
					.append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">")
					.append("<title>LearnHub Email</title>").append("<style>")
					.append("body { font-family: Arial, sans-serif; background-color: #f5f6fa; color: #333; margin: 0; padding: 0; }")
					.append(".container { max-width: 600px; margin: 0 auto; padding: 20px; background-color: #fff; border: 1px solid #e0e0e0; border-radius: 8px; }")
					.append(".banner img { width: 100%; max-width: 600px; border-radius: 8px; }")
					.append(".content h2 { margin-top: 0; color: #2c3e50; }").append(".content h3 { color: #2980b9; }")
					.append(".content ul { padding-left: 20px; }")
					.append(".content ul li { margin-bottom: 10px; line-height: 1.6; }")
					.append(".content p { line-height: 1.8; }")
					.append(".image-center { text-align: center; margin: 20px 0; }")
					.append(".image-center img { width: 80%; max-width: 600px; border-radius: 8px; }")
					.append(".footer { border-top: 1px solid #ddd; padding-top: 20px; margin-top: 20px; text-align: center; color: #777; font-size: 14px; }")
					.append(".footer a { color: #3498db; text-decoration: none; }")
					.append(".footer .social-icons { margin-top: 15px; }")
					.append(".footer .social-icons a { margin: 0 10px; display: inline-block; }")
					.append(".footer .social-icons img { width: 24px; vertical-align: middle; }")
					.append(".footer p { margin-top: 20px; }").append("</style>").append("</head>").append("<body>")
					.append("<div class=\"container\">").append("<div class=\"banner\">")
					.append("<img src=\"https://imgur.com/RdKXEVL.jpg\" alt=\"LearnHub Banner\" />").append("</div>")
					.append("<div class=\"content\">")
					.append("<h2>Boost Your Business with LearnHub &ndash; The Ultimate Online Training Platform</h2>")
					.append("<h3>Hello, Dear ").append(recipientName).append("</h3>") // Dynamically adding recipient's
																						// name
					.append("<p>We&rsquo;re excited to introduce ")
					.append("<a href=\"https://learnhub.vsmartengine.com\" style=\"color: #3498db;\">LearnHub</a>, ")
					.append("a powerful online platform designed to elevate your training business and maximize your revenue potential. ")
					.append("Whether you're a training institute, expert tutor, or specialized trainer, LearnHub offers tools to manage and grow your online course offerings effortlessly.")
					.append("</p>").append("<h3>Why Choose LearnHub?</h3>").append("<div class=\"image-center\">")
					.append("<img  src=\"https://lh3.googleusercontent.com/pw/AP1GczMF4fLW8TQDxq9RxhV1rOpDswerz1dIDbqipF5SwIuYMt8SQssfou0yqKy9T--cWLg7O3wqq7x7RWcfBthUG10cAhtDhk5hwHOR3c50l5FS9sGQrGYz4RP-lLusmlmQicox4DPvvW-wjHSpeLFOFCk6VQ=w1273-h716-s-no-gm?authuser=0\" />")
					.append("</div>").append("<ul>")
					.append("<li><strong>Expand Your Reach:</strong> Build a professional online presence to attract a larger audience.</li>")
					.append("<li><strong>Boost Your Profits:</strong> Streamlined management and flexible pricing to increase earnings.</li>")
					.append("<li><strong>Advanced Features:</strong> Secure content delivery, scheduling tools, and analytics for performance tracking.</li>")
					.append("</ul>")
					.append("<p>Let LearnHub transform your business and unlock new revenue opportunities. Ready to get started? ")
					.append("Visit ")
					.append("<a href=\"https://lhdemo.vsmartengine.com/\" style=\"color: #3498db;\">LearnHub</a> ")
					.append("to learn more.</p>").append("<p>Warm regards,</p>")
					.append("<p>Varun Karthick R<br />VSmartEngine Team</p>").append("</div>")
					.append("<div class=\"footer\">").append("<p>Have questions? ")
					.append("<a href=\"mailto:salesinfo@vsmartengine.com\">Contact Support</a> ")
					.append("or visit our ")
					.append("<a href=\"https://learnhub.vsmartengine.com/#contact\">Help Center</a>.</p>")
					.append("<div class=\"social-icons\">").append("<a href=\"https://x.com/vsmartengine\">")
					.append("<img src=\"https://i.imgur.com/IZ2a8vM.png\" alt=\"X\" />").append("</a>")
					.append("<a href=\"https://www.facebook.com/profile.php?id=61563355049634\">")
					.append("<img src=\"https://i.imgur.com/hbJNKsm.png\" alt=\"Facebook\" />").append("</a>")
					.append("<a href=\"#\">")
					.append("<img src=\"https://cdn-icons-png.flaticon.com/512/2111/2111646.png\" alt=\"Telegram\" />")
					.append("</a>").append("<a href=\"https://www.linkedin.com/company/101757366/\">")
					.append("<img src=\"https://i.imgur.com/aGY4UAZ.png\" alt=\"LinkedIn\" />").append("</a>")
					.append("<a href=\"https://www.instagram.com/vsmartengine/\">")
					.append("<img src=\"https://i.imgur.com/fMseZg4.png\" alt=\"Instagram\" />").append("</a>")
					.append("</div>").append("<p>&copy; 2024 LearnHub. All rights reserved.</p>").append("</div>")
					.append("</div>").append("</body>").append("</html>");

			String emailContent = htmlContent2.toString();

			helper.setText(emailContent, true);
			// Send the email
			mailSender.send(mimeMessage);
			return "Success";
		} catch (Exception e) {
			System.out.println(e.getMessage());
			return "Error: " + e.getMessage();
		}
	}
}
