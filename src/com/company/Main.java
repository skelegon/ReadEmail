package com.company;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import javax.mail.BodyPart;
import javax.mail.Session;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.io.*;
import java.util.Date;
import java.util.Properties;

public class Main {

    public static void main(String[] args) throws Exception {

        System.out.println("Identifying the current directory");
        File curDir = new File(".");

        System.out.println("Checking for e-mails in current directory");
        File[] emails = getFiles(curDir, ".eml", ".msg"); //TODO tagastab ainult esimese argumendi?
        if (emails.length==0) {
            System.out.println("No e-mails found in current directory. Closing the program.");
            System.exit(0);
        } else {
            System.out.println("Total e-mails found in current folder: " + emails.length );
            createXlsFile();
            for (File email: emails) {
                readEmailData(email);
            }
        }
    }

    private static void createXlsFile() throws IOException {
        String excelFileName = "Test.xls";//name of excel file
        String sheetName = "Sheet1";//name of sheet
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet(sheetName) ;
        FileOutputStream fileOut = new FileOutputStream(excelFileName);
        //write this workbook to an Outputstream.
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
    }

    private static void readEmailData(File email) throws Exception{
        Properties props = System.getProperties();
        props.put("mail.host", "smtp.dummydomain.com");
        props.put("mail.transport.protocol", "smtp");

        Session mailSession = Session.getDefaultInstance(props, null);
        InputStream source = new FileInputStream(email);
        MimeMessage message = new MimeMessage(mailSession, source);

        String textMessage = getTextFromMessage(message);
        String subject = message.getSubject();
        String from = String.valueOf(message.getFrom()[0]);
        Date sentDate = message.getSentDate();

        insertToExcelTable(sentDate,from,subject,textMessage);
    }

    private static void insertToExcelTable(Date sentDate, String from, String subject, String textMessage) throws IOException{
        String excelFileName = "Test.xls";//name of excel file
        InputStream ExcelFileToRead = new FileInputStream("Test.xls");
        HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);

        HSSFSheet sheet=wb.getSheetAt(0);
        System.out.println(sheet.getLastRowNum());
        HSSFRow row = sheet.createRow(sheet.getLastRowNum()+1);

        String [] eMailData = {sentDate.toString(), from, subject, textMessage};

        for (int i = 0; i < 4; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellValue(eMailData[i]);
        }

        FileOutputStream fileOut = new FileOutputStream(excelFileName);

        //write this workbook to an Outputstream.
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
    }

    private static File[] getFiles(File curDir, String ... args) {
        File[] files = curDir.listFiles(new FilenameFilter() {
            public boolean accept(File dir, String name) {
                for (String arg: args) {
                    return name.toLowerCase().endsWith(arg);
                } return true;
                //TODO Miks return siin?
            }
        });
        return files;
    }

    private static String getTextFromMessage(MimeMessage message) throws Exception {
        String result = "";
        if (message.isMimeType("text/plain")) {
            result = message.getContent().toString();
        } else if (message.isMimeType("multipart/*")) {
            MimeMultipart mimeMultipart = (MimeMultipart) message.getContent();
            result = getTextFromMimeMultipart(mimeMultipart);
        }
        return result;
    }

    private static String getTextFromMimeMultipart(MimeMultipart mimeMultipart) throws Exception{
        String result = "";
        int count = mimeMultipart.getCount();
        for (int i = 0; i < count; i++) {
            BodyPart bodyPart = mimeMultipart.getBodyPart(i);
            if (bodyPart.isMimeType("text/plain")) {
                result = result + "\n" + bodyPart.getContent();
                break; // without break same text appears twice in my tests
            } else if (bodyPart.isMimeType("text/html")) {
                String html = (String) bodyPart.getContent();
                result = result + "\n" + org.jsoup.Jsoup.parse(html).text();
            } else if (bodyPart.getContent() instanceof MimeMultipart){
                result = result + getTextFromMimeMultipart((MimeMultipart)bodyPart.getContent());
            }
        }
        return result;
    }
}