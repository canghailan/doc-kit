package cc.whohow.word;

import com.fasterxml.jackson.databind.JsonNode;
import org.apache.poi.xwpf.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.lang.reflect.UndeclaredThrowableException;
import java.net.URL;

public class XWPFDocumentTemplate {
    private URL url;

    public XWPFDocumentTemplate(URL url) {
        this.url = url;
    }

    public XWPFDocument process(JsonNode data) {
        try (InputStream stream = url.openStream()) {

            XWPFDocument template = new XWPFDocument(stream);

            for (XWPFParagraph paragraph : template.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    System.out.println(run.getText(0));
                    System.out.println(run.getText(1));
                }
            }


            for (XWPFTable table : template.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        System.out.println(cell.getText());
                    }
                }
            }

            return null;
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        } catch (Exception e) {
            throw new UndeclaredThrowableException(e);
        }
    }
}
