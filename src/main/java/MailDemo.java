/*
    Description:
    Author:zeratul
    Time: 2022/4/14-上午11:49
*/


import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.mail.*;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import java.io.*;


import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

public class MailDemo {
    public static void main(String[] args) {
        Properties pros = new Properties();
        MailDemo mail = new MailDemo();
        InputStream is = null;

        try {
//            is = new FileInputStream("MailUtil.properties");
            ClassLoader classLoader = MailDemo.class.getClassLoader();
            is = classLoader.getResourceAsStream("MailUtil.properties");
            BufferedInputStream excelIs = new BufferedInputStream(new FileInputStream(new File("/home/zeratul/companyInfo.xlsx")));
            XSSFWorkbook workbook = new XSSFWorkbook(excelIs);
            BufferedInputStream wordIs = new BufferedInputStream(new FileInputStream(new File("/home/zeratul/content.docx")));
            FileInputStream wordFile = new FileInputStream("/home/zeratul/content.docx");
            String content = mail.getContent(wordFile);
            pros.load(is);
            Map<String, String> recipients = mail.recipients(workbook);
            mail.sendGroupEmails(pros,recipients,content);


        } catch (FileNotFoundException | MessagingException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            try {
                if (is != null) {
                    is.close();

                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }


    }

    public void sendGroupEmails(Properties pros,Map<String,String> map,String content) throws MessagingException {
        String uname = pros.getProperty("username");
        String password = pros.getProperty("password");
        String address =pros.getProperty("internetAddress");
        Authenticator authenticator = new Authenticator(){
            @Override
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(uname,password);
            }
        };

        Session session = Session.getInstance(pros, authenticator);

        MimeMessage message = new MimeMessage(session);
        message.setFrom(new InternetAddress(address));
        for (Map.Entry<String,String> entry : map.entrySet()){
            String comName = entry.getKey();
            String comAddress = entry.getValue();
            message.setRecipient(Message.RecipientType.TO,new InternetAddress(comAddress));
            message.setSubject("邮件");

            message.setContent(comName + ":\n" + content,"text/html;charset=utf-8");
            Transport.send(message);
        }


    }

    public Map<String,String> recipients(Workbook workbook) {
        HashMap<String, String> map = new HashMap<>();
        Sheet sheet = workbook.getSheet("Sheet1");
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        for (int i = firstRowNum;i<=lastRowNum;i++){
            Row row = sheet.getRow(i);
            Cell comname = row.getCell(0);
            Cell email = row.getCell(2);
            map.put(comname.getStringCellValue(),email.getStringCellValue());

        }
        return map;
    }

    public String getContent(InputStream is) throws IOException {
        StringBuilder stringBuilder = new StringBuilder();
        is = FileMagic.prepareToCheckMagic(is);


        if (FileMagic.valueOf(is) == FileMagic.OLE2){
            WordExtractor wordExtractor = new WordExtractor(is);
            String line = System.getProperty("line.separator");
            stringBuilder.append("<br>").append(wordExtractor.getText());
            wordExtractor.close();
        }else if (FileMagic.valueOf(is) == FileMagic.OOXML){
            XWPFDocument xwpfDocument = new XWPFDocument(is);
            XWPFWordExtractor xwpfWordExtractor = new XWPFWordExtractor(xwpfDocument);
            String line = System.getProperty("line.separator");
            stringBuilder.append("<br>").append(xwpfWordExtractor.getText());
            xwpfDocument.close();
        }

        return stringBuilder.toString();
    }
}
