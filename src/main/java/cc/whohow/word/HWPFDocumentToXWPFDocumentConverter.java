package cc.whohow.word;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;

import java.util.function.Function;

public class HWPFDocumentToXWPFDocumentConverter implements Function<HWPFDocument, XWPFDocument> {
    @Override
    public XWPFDocument apply(HWPFDocument source) {
        XWPFDocument target = new XWPFDocument();
        Range range = source.getRange();
        for (int sectionIndex = 0; sectionIndex < range.numSections(); sectionIndex++) {
            Section section = range.getSection(sectionIndex);
            for (int paragraphIndex = 0; paragraphIndex < section.numParagraphs(); paragraphIndex++) {
                Paragraph paragraph = section.getParagraph(paragraphIndex);

                if (paragraph.isInTable()) {
                    Table table = section.getTable(paragraph);
                    paragraphIndex += table.numParagraphs() - 1;
                    copy(source, table, target, target.createTable());
                } else {
                    copy(source, paragraph, target, target.createParagraph());
                }
            }
        }
        return target;
    }

    private void copy(HWPFDocument source, Table sourceTable,
                      XWPFDocument target, XWPFTable targetTable) {
        for (int rowIndex = 0; rowIndex < sourceTable.numRows(); rowIndex++) {
            TableRow tableRow = sourceTable.getRow(rowIndex);
            XWPFTableRow xwpfTableRow = targetTable.createRow();
            for (int cellIndex = 0; cellIndex < tableRow.numCells(); cellIndex++) {
                TableCell tableCell = tableRow.getCell(cellIndex);
                XWPFTableCell xwpfTableCell = xwpfTableRow.createCell();
                for (int i = 0; i < tableCell.numParagraphs(); i++) {
                    copy(source, tableCell.getParagraph(i), target, xwpfTableCell.addParagraph());
                }
            }
        }
    }

    private void copy(HWPFDocument source, Paragraph sourceParagraph,
                      XWPFDocument target, XWPFParagraph targetParagraph) {
        for (int runIndex = 0; runIndex < sourceParagraph.numCharacterRuns(); runIndex++) {
            CharacterRun run = sourceParagraph.getCharacterRun(runIndex);
            XWPFRun xwpfRun = targetParagraph.createRun();
            xwpfRun.setText(run.text());
        }
    }
}
