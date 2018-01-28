package cc.whohow.word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.odftoolkit.simple.TextDocument;

import java.util.function.Function;

public class XWPFDocumentToTextDocumentConverter implements Function<XWPFDocument, TextDocument> {
    @Override
    public TextDocument apply(XWPFDocument source) {
        throw new UnsupportedOperationException();
    }
}
