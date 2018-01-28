package cc.whohow.word;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.util.function.Function;

public class XWPFDocumentToHWPFDocumentConverter implements Function<XWPFDocument, HWPFDocument> {
    @Override
    public HWPFDocument apply(XWPFDocument source) {
        throw new UnsupportedOperationException();
    }
}
