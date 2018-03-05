package com.rabitabank.docxtopdfconverter;

import com.itextpdf.text.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import com.itextpdf.text.pdf.*;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;



public class WordReplaceText {
    public XWPFDocument replaceText(XWPFDocument doc, String findText, String replaceText) {
        for (XWPFParagraph p : doc.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    if (text != null && text.contains(findText)) {
                        text = text.replace(findText, replaceText);
                        r.setText(text, 0);
                    }
                }
            }
        }
       /* for (XWPFTable tbl : doc.getTables()) {
           for (XWPFTableRow row : tbl.getRows()) {
              for (XWPFTableCell cell : row.getTableCells()) {
                 for (XWPFParagraph p : cell.getParagraphs()) {
                    for (XWPFRun r : p.getRuns()) {
                      String text = r.getText(0);
                      if (text != null && text.contains("needle")) {
                        text = text.replace("needle", "haystack");
                        r.setText(text,0);
                      }
                    }
                 }
              }
           }
        }*/
        return doc;
    }

    public XWPFDocument openDocument(String filename) throws Exception {
        try{
            File file = new File(filename);
            FileInputStream fis = new FileInputStream(file);
            XWPFDocument document = null;
            if (fis != null) {
                document = new XWPFDocument(fis);
            }
            return document;
        } catch (Exception e) {
             e.printStackTrace();
             return null;
        }
    }

    public void saveDocument(XWPFDocument doc, String file) {
        try (FileOutputStream out = new FileOutputStream(file)) {
            doc.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void createPDF(XWPFDocument document, OutputStream out) {
        try {
            PdfOptions options = PdfOptions.getDefault();
            PdfConverter.getInstance().convert(document, out,options);
        } catch (Throwable e) {
            e.printStackTrace();
        }
    }
    
    public void manipulatePdf(String src,  OutputStream out, String findText, String replaceText) throws IOException, DocumentException {
        try {
        PdfReader reader = new PdfReader(src);
        PdfStamper stamper = new PdfStamper(reader, out);
        AcroFields form = stamper.getAcroFields();
        form.setField(findText, replaceText);
        stamper.close();
        reader.close();
        } catch (Throwable e) {
            e.printStackTrace();
        }
    }
}
