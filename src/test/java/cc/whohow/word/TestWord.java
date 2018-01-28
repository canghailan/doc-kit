package cc.whohow.word;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;

import java.io.*;
import java.util.Arrays;

public class TestWord {
    @Test
    public void testXWPFDocument() throws Exception {
        try (InputStream stream = new FileInputStream("template.docx")) {
            XWPFDocument doc = new XWPFDocument(stream);
            System.out.println(doc);
            System.out.println(doc.getAllEmbedds());
            System.out.println(doc.getAllPackagePictures());
            System.out.println(doc.getAllPictures());
            System.out.println(doc.getBodyElements());
            System.out.println(Arrays.toString(doc.getComments()));
            System.out.println(doc.getFooterList());
            System.out.println(doc.getFootnotes());
            System.out.println(doc.getHeaderFooterPolicy());
            System.out.println(doc.getHeaderList());
            System.out.println(Arrays.toString(doc.getHyperlinks()));
            System.out.println(doc.getNumbering());
            System.out.println(doc.getParagraphs());
            System.out.println(doc.getPart());
            System.out.println(doc.getPartType());
            System.out.println(doc.getStyle());
            System.out.println(doc.getStyles());
            System.out.println(doc.getTables());
            System.out.println(doc.getXWPFDocument());
            System.out.println(doc.getPackagePart());
            System.out.println(doc.getParent());
            System.out.println(doc.getProperties());
            System.out.println(doc.getRelationParts());
            System.out.println(doc.getRelations());
        }
    }

    @Test
    public void testDocToDocx() throws Exception {
        HWPFDocumentToXWPFDocumentConverter converter = new HWPFDocumentToXWPFDocumentConverter();
        try (InputStream input = new FileInputStream("template.doc");
             OutputStream output = new FileOutputStream("template-transform.docx")) {
            converter.apply(new HWPFDocument(input)).write(output);
        }
    }

    @Test
    public void test() throws Exception {
        ObjectMapper objectMapper = new ObjectMapper();
        XWPFDocumentTemplate XWPFDocumentTemplate = new XWPFDocumentTemplate(new File("template.docx").toURI().toURL());
        XWPFDocumentTemplate.process(objectMapper.readTree(new File("data.json")));
    }
}
