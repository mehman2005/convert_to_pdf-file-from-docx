/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.rabitabank.docxtopdfconverter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 *
 * @author Mehman Abasov
 */
public class MainClass {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        try {
            // TODO code application logic here
            String SOURCE_FILE = "D:\\template.docx";
            String OUTPUT_FILE_WORD =  "D:\\template_new.docx";
            String OUTPUT_FILE_PDF =  "D:\\template_new.pdf";
            
            WordReplaceText instance = new WordReplaceText();
            XWPFDocument workbook;
            workbook = instance.openDocument(SOURCE_FILE);
            
            if (workbook != null) {
                OutputStream out = new FileOutputStream(new File(OUTPUT_FILE_PDF));
                workbook = instance.replaceText(workbook, "p_template_word", "new word");
                
                // Save new word file
                instance.saveDocument(workbook, OUTPUT_FILE_WORD);
                
                // Save pdf file
                instance.createPDF(workbook, out);
            }
        } catch (Exception ex) {
            Logger.getLogger(MainClass.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
}
