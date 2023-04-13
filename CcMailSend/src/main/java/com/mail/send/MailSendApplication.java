package com.mail.send;

import com.mail.server.MailService;
import java.io.File;
import java.io.FileInputStream;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.time.Month;
import java.util.ArrayList;
import java.util.Date;
import java.util.GregorianCalendar;
import javax.mail.internet.MimeMessage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.thymeleaf.TemplateEngine;
import org.thymeleaf.context.Context;
import org.thymeleaf.context.IContext;
public class MailSendApplication implements MailService {
	
	
	static String currentDir = System.getProperty("user.dir");
	  
	  @Autowired
	  private JavaMailSender javaMailSender;
	  
	  @Autowired
	  private TemplateEngine templateEngine;
	  
	  public void mailSend(String From) {
	    try {
	      FileInputStream names = new FileInputStream(new File("C:\\Users\\ITCae\\OneDrive\\Desktop\\Meal.xlsx"));
	      XSSFWorkbook workbook = new XSSFWorkbook(names);
	      XSSFSheet mealSheet = workbook.getSheetAt(0);
	      
	      FileInputStream ids = new FileInputStream(new File("C:\\Users\\ITCae\\OneDrive\\Desktop\\Ids.xlsx"));
	      XSSFWorkbook wb = new XSSFWorkbook(ids);
	      XSSFSheet idsSheet = wb.getSheetAt(0);
	      MimeMessage message = this.javaMailSender.createMimeMessage();
	      MimeMessageHelper helper = new MimeMessageHelper(message, 3, StandardCharsets.UTF_8.name());
	      DataFormatter formatter = new DataFormatter();
	      int date = 1, date1 = 2, date2 = 3, date3 = 4, date4 = 5, date5 = 6, date6 = 7, date7 = 8, date8 = 9;
	      int date9 = 10;
	      int date10 = 11, date11 = 12, date12 = 13, date13 = 14, date14 = 15, date15 = 16, date16 = 17, date17 = 18;
	      int date18 = 19, date19 = 20;
	      int date20 = 21, date21 = 22, date22 = 23, date23 = 24, date24 = 25, date25 = 26, date26 = 27, date27 = 28;
	      int date28 = 29, date29 = 30, date30 = 31;
	      String day = new String();
	      String intime = new String();
	      String outtime = new String();
	      String totaltime = new String();
	      String day1 = new String();
	      String intime1 = new String();
	      String outtime1 = new String();
	      String totaltime1 = new String();
	      String day2 = new String();
	      String intime2 = new String();
	      String outtime2 = new String();
	      String totaltime2 = new String();
	      String day3 = new String();
	      String intime3 = new String();
	      String outtime3 = new String();
	      String totaltime3 = new String();
	      String day4 = new String();
	      String intime4 = new String();
	      String outtime4 = new String();
	      String totaltime4 = new String();
	      String day5 = new String();
	      String intime5 = null;
	      String outtime5 = new String();
	      String totaltime5 = new String();
	      String day6 = new String();
	      String intime6 = new String();
	      String outtime6 = new String();
	      String totaltime6 = new String();
	      String day7 = new String();
	      String intime7 = new String();
	      String outtime7 = new String();
	      String totaltime7 = new String();
	      String day8 = new String();
	      String intime8 = new String();
	      String outtime8 = new String();
	      String totaltime8 = new String();
	      String day9 = new String();
	      String intime9 = new String();
	      String outtime9 = new String();
	      String totaltime9 = new String();
	      String day10 = new String();
	      String intime10 = new String();
	      String outtime10 = new String();
	      String totaltime10 = new String();
	      String day11 = new String();
	      String intime11 = new String();
	      String outtime11 = new String();
	      String totaltime11 = new String();
	      String day12 = new String();
	      String intime12 = new String();
	      String outtime12 = new String();
	      String totaltime12 = new String();
	      String day13 = new String();
	      String intime13 = new String();
	      String outtime13 = new String();
	      String totaltime13 = new String();
	      String day14 = new String();
	      String intime14 = new String();
	      String outtime14 = new String();
	      String totaltime14 = new String();
	      String day15 = new String();
	      String intime15 = new String();
	      String outtime15 = new String();
	      String totaltime15 = new String();
	      String day16 = new String();
	      String intime16 = new String();
	      String outtime16 = new String();
	      String totaltime16 = new String();
	      String day17 = new String();
	      String intime17 = new String();
	      String outtime17 = new String();
	      String totaltime17 = new String();
	      String day18 = new String();
	      String intime18 = new String();
	      String outtime18 = new String();
	      String totaltime18 = new String();
	      String day19 = new String();
	      String intime19 = new String();
	      String outtime19 = new String();
	      String totaltime19 = new String();
	      String day20 = new String();
	      String intime20 = new String();
	      String outtime20 = new String();
	      String totaltime20 = new String();
	      String day21 = new String();
	      String intime21 = new String();
	      String outtime21 = new String();
	      String totaltime21 = new String();
	      String day22 = new String();
	      String intime22 = new String();
	      String outtime22 = new String();
	      String totaltime22 = new String();
	      String day23 = new String();
	      String intime23 = new String();
	      String outtime23 = new String();
	      String totaltime23 = new String();
	      String day24 = new String();
	      String intime24 = new String();
	      String outtime24 = new String();
	      String totaltime24 = new String();
	      String day25 = new String();
	      String intime25 = new String();
	      String outtime25 = new String();
	      String totaltime25 = new String();
	      String day26 = new String();
	      String intime26 = new String();
	      String outtime26 = new String();
	      String totaltime26 = new String();
	      String day27 = new String();
	      String intime27 = new String();
	      String outtime27 = new String();
	      String totaltime27 = new String();
	      String day28 = new String();
	      String intime28 = new String();
	      String outtime28 = new String();
	      String totaltime28 = new String();
	      String day29 = new String();
	      String intime29 = new String();
	      String outtime29 = new String();
	      String totaltime29 = new String();
	      String day30 = new String();
	      String intime30 = new String();
	      String outtime30 = new String();
	      String totaltime30 = new String();
	      String name2 = new String();
	      String email = new String();
	      
	      SimpleDateFormat simpleformat = new SimpleDateFormat(" MMMM");
	      String strMonth = simpleformat.format(new Date());
	      GregorianCalendar c = (GregorianCalendar)GregorianCalendar.getInstance();
	      SimpleDateFormat format = new SimpleDateFormat("dd");
	      String dates = format.format(new Date());
	      String mm = new String();
	     
	      if (dates.equals("01") || dates.equals("02") || dates.equals("03")) {
	        c.add(2, -1);
	        int month = c.get(2) + 1;
	        Month m = Month.of(month);
	        mm = m.toString();
	        System.out.println(mm);
	      } else {
	        mm = strMonth;
	      } 
	      ArrayList<String> list2 = new ArrayList<>();
	      for (int i = 1; i <= idsSheet.getLastRowNum(); i++) {
	        name2 = idsSheet.getRow(i).getCell(0).getStringCellValue();
	        list2.add(name2);
	      } 
	      Context context = new Context();
	      for (int j = 1; j <= mealSheet.getLastRowNum(); j++) {
	        String name1 = mealSheet.getRow(j).getCell(1).getStringCellValue();
	        if (list2.contains(name1)) {
	          try {
	            email = idsSheet.getRow(list2.indexOf(name1) + 1).getCell(1).getStringCellValue();
	          } catch (Exception e) {
	             
	          } 
	          int k = 2;
	          int days = 2;
	          context.setVariable("Name", name1);
	          System.out.println(name1);
	          try {
	            day = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(2);
	            days += 4;
	            intime = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day", day);
	            context.setVariable("Date", Integer.valueOf(date));
	            context.setVariable("InTime", intime);
	            context.setVariable("OutTime", outtime);
	            context.setVariable("Total", totaltime);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day1 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(2);
	            days += 4;
	            intime1 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime1 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime1 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day1", day1);
	            context.setVariable("Date1", Integer.valueOf(date1));
	            context.setVariable("InTime1", intime1);
	            context.setVariable("OutTime1", outtime1);
	            context.setVariable("Total1", totaltime1);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day2 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(2);
	            days += 4;
	            intime2 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime2 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime2 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day2", day2);
	            context.setVariable("Date2", Integer.valueOf(date2));
	            context.setVariable("InTime2", intime2);
	            context.setVariable("OutTime2", outtime2);
	            context.setVariable("Total2", totaltime2);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day3 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(2);
	            days += 4;
	            intime3 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime3 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime3 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day3", day3);
	            context.setVariable("Date3", Integer.valueOf(date3));
	            context.setVariable("InTime3", intime3);
	            context.setVariable("OutTime3", outtime3);
	            context.setVariable("Total3", totaltime3);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day4 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(2);
	            days += 4;
	            intime4 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime4 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime4 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day4", day4);
	            context.setVariable("Date4", Integer.valueOf(date4));
	            context.setVariable("InTime4", intime4);
	            context.setVariable("OutTime4", outtime4);
	            context.setVariable("Total4", totaltime4);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day5 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(2);
	            days += 4;
	            intime5 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime5 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime5 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day5", day5);
	            context.setVariable("Date5", Integer.valueOf(date5));
	            context.setVariable("InTime5", intime5);
	            context.setVariable("OutTime5", outtime5);
	            context.setVariable("Total5", totaltime5);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day6 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(2);
	            days += 4;
	            intime6 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime6 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime6 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day6", day6);
	            context.setVariable("Date6", Integer.valueOf(date6));
	            context.setVariable("InTime6", intime6);
	            context.setVariable("OutTime6", outtime6);
	            context.setVariable("Total6", totaltime6);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day7 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(2);
	            days += 4;
	            intime7 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime7 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime7 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day7", day7);
	            context.setVariable("Date7", Integer.valueOf(date7));
	            context.setVariable("InTime7", intime7);
	            context.setVariable("OutTime7", outtime7);
	            context.setVariable("Total7", totaltime7);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day8 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(2);
	            days += 4;
	            intime8 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime8 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime8 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day8", day8);
	            context.setVariable("Date8", Integer.valueOf(date8));
	            context.setVariable("InTime8", intime8);
	            context.setVariable("OutTime8", outtime8);
	            context.setVariable("Total8", totaltime8);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day9 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime9 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime9 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime9 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day9", day9);
	            context.setVariable("Date9", Integer.valueOf(date9));
	            context.setVariable("InTime9", intime9);
	            context.setVariable("OutTime9", outtime9);
	            context.setVariable("Total9", totaltime9);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day10 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime10 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime10 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime10 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day10", day10);
	            context.setVariable("Date10", Integer.valueOf(date10));
	            context.setVariable("InTime10", intime10);
	            context.setVariable("OutTime10", outtime10);
	            context.setVariable("Total10", totaltime10);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day11 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime11 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime11 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime11 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day11", day11);
	            context.setVariable("Date11", Integer.valueOf(date11));
	            context.setVariable("InTime11", intime11);
	            context.setVariable("OutTime11", outtime11);
	            context.setVariable("Total11", totaltime11);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day12 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime12 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime12 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime12 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day12", day12);
	            context.setVariable("Date12", Integer.valueOf(date12));
	            context.setVariable("InTime12", intime12);
	            context.setVariable("OutTime12", outtime12);
	            context.setVariable("Total12", totaltime12);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day13 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime13 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime13 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime13 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day13", day13);
	            context.setVariable("Date13", Integer.valueOf(date13));
	            context.setVariable("InTime13", intime13);
	            context.setVariable("OutTime13", outtime13);
	            context.setVariable("Total13", totaltime13);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day14 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime14 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime14 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime14 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day14", day14);
	            context.setVariable("Date14", Integer.valueOf(date14));
	            context.setVariable("InTime14", intime14);
	            context.setVariable("OutTime14", outtime14);
	            context.setVariable("Total14", totaltime14);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day15 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime15 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime15 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime15 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day15", day15);
	            context.setVariable("Date15", Integer.valueOf(date15));
	            context.setVariable("InTime15", intime15);
	            context.setVariable("OutTime15", outtime15);
	            context.setVariable("Total15", totaltime15);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day16 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime16 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime16 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime16 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day16", day16);
	            context.setVariable("Date16", Integer.valueOf(date16));
	            context.setVariable("InTime16", intime16);
	            context.setVariable("OutTime16", outtime16);
	            context.setVariable("Total16", totaltime16);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day17 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime17 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime17 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime17 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day17", day17);
	            context.setVariable("Date17", Integer.valueOf(date17));
	            context.setVariable("InTime17", intime17);
	            context.setVariable("OutTime17", outtime17);
	            context.setVariable("Total17", totaltime17);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day18 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime18 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime18 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime18 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day18", day18);
	            context.setVariable("Date18", Integer.valueOf(date18));
	            context.setVariable("InTime18", intime18);
	            context.setVariable("OutTime18", outtime18);
	            context.setVariable("Total18", totaltime18);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day19 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime19 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime19 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime19 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day19", day19);
	            context.setVariable("Date19", Integer.valueOf(date19));
	            context.setVariable("InTime19", intime19);
	            context.setVariable("OutTime19", outtime19);
	            context.setVariable("Total19", totaltime19);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day20 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime20 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime20 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime20 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day20", day20);
	            context.setVariable("Date20", Integer.valueOf(date20));
	            context.setVariable("InTime20", intime20);
	            context.setVariable("OutTime20", outtime20);
	            context.setVariable("Total20", totaltime20);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day21 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime21 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime21 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime21 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day21", day21);
	            context.setVariable("Date21", Integer.valueOf(date21));
	            context.setVariable("InTime21", intime21);
	            context.setVariable("OutTime21", outtime21);
	            context.setVariable("Total21", totaltime21);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day22 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime22 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime22 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime22 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day22", day22);
	            context.setVariable("Date22", Integer.valueOf(date22));
	            context.setVariable("InTime22", intime22);
	            context.setVariable("OutTime22", outtime22);
	            context.setVariable("Total22", totaltime22);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day23 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime23 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime23 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime23 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day23", day23);
	            context.setVariable("Date23", Integer.valueOf(date23));
	            context.setVariable("InTime23", intime23);
	            context.setVariable("OutTime23", outtime23);
	            context.setVariable("Total23", totaltime23);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day24 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime24 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime24 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime24 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day24", day24);
	            context.setVariable("Date24", Integer.valueOf(date24));
	            context.setVariable("InTime24", intime24);
	            context.setVariable("OutTime24", outtime24);
	            context.setVariable("Total24", totaltime24);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day25 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime25 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime25 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime25 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day25", day25);
	            context.setVariable("Date25", Integer.valueOf(date25));
	            context.setVariable("InTime25", intime25);
	            context.setVariable("OutTime25", outtime25);
	            context.setVariable("Total25", totaltime25);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day26 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime26 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime26 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime26 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day26", day26);
	            context.setVariable("Date26", Integer.valueOf(date26));
	            context.setVariable("InTime26", intime26);
	            context.setVariable("OutTime26", outtime26);
	            context.setVariable("Total26", totaltime26);
	          } catch (Exception e) {
	            e.getMessage();
	          } 
	          try {
	            day27 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime27 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime27 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime27 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day27", day27);
	            context.setVariable("Date27", Integer.valueOf(date27));
	            context.setVariable("InTime27", intime27);
	            context.setVariable("OutTime27", outtime27);
	            context.setVariable("Total27", totaltime27);
	          } catch (Exception e) {
	            e.getMessage();
	          } 
	          try {
	            day28 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime28 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime28 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime28 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day28", day28);
	            context.setVariable("Date28", Integer.valueOf(date28));
	            context.setVariable("InTime28", intime28);
	            context.setVariable("OutTime28", outtime28);
	            context.setVariable("Total28", totaltime28);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day29 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime29 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime29 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime29 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day29", day29);
	            context.setVariable("Date29", Integer.valueOf(date29));
	            context.setVariable("InTime29", intime29);
	            context.setVariable("OutTime29", outtime29);
	            context.setVariable("Total29", totaltime29);
	          } catch (Exception e) {
	             
	          } 
	          try {
	            day30 = mealSheet.getRow(0).getCell(days).getStringCellValue().substring(3);
	            days += 4;
	            intime30 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            outtime30 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k++;
	            totaltime30 = formatter.formatCellValue(mealSheet.getRow(j).getCell(k));
	            k += 2;
	            context.setVariable("Day30", day30);
	            context.setVariable("Date30", Integer.valueOf(date30));
	            context.setVariable("InTime30", intime30);
	            context.setVariable("OutTime30", outtime30);
	            context.setVariable("Total30", totaltime30);
	          } catch (Exception e) {
	             
	          } 
	          String html = this.templateEngine.process("Attendanceformat", (IContext)context);
	          helper.setText(html, true);
	          helper.setFrom(From);
	          String subject = "Attendance Record " + mm;
	          helper.setTo(email);
	          helper.setSubject(subject);
	          this.javaMailSender.send(message);
	        } 
	      } 
	      System.out.println("send mail");
	      workbook.close();
	      wb.close();
	    } catch (Exception e) {
	      
	    } 
	  }
	}


