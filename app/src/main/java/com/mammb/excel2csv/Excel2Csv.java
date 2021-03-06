package com.mammb.excel2csv;

import com.github.mygreen.cellformatter.ObjectCellFormatter;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.Closeable;
import java.io.Flushable;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.Date;
import java.util.Objects;

public class Excel2Csv {

    private final Path xlsx;
    private final Path csv;
    private final String sheetName;

    public Excel2Csv(Path xlsx, Path csv, String sheetName) {
        this.xlsx = Objects.requireNonNull(xlsx);
        this.csv = Objects.requireNonNull(csv);
        this.sheetName = Objects.isNull(sheetName) ? "" : sheetName;
    }

    public void process() throws IOException, OpenXML4JException, SAXException {

        try (final CsvPrinter printer = new CsvPrinter(new CSVPrinter(
                new PrintWriter(csv.toFile(), StandardCharsets.UTF_8), CSVFormat.EXCEL));
             final OPCPackage pkg = OPCPackage.open(xlsx.toFile(), PackageAccess.READ)) {

            final ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);
            final XSSFReader xssfReader = new XSSFReader(pkg);
            final StylesTable styles = xssfReader.getStylesTable();
            final XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData();

            while (sheets.hasNext()) {
                try (InputStream stream = sheets.next()) {
                    if (sheetName.isEmpty() || sheetName.equals(sheets.getSheetName())) {
                        processSheet(
                                styles,
                                strings,
                                new SheetHandler(printer),
                                stream);
                        return;
                    }
                }
            }
            throw new IOException("not found sheet: " + sheetName);
        }
    }

    private void processSheet(
            Styles styles,
            SharedStrings strings,
            XSSFSheetXMLHandler.SheetContentsHandler sheetHandler,
            InputStream sheetInputStream) throws IOException, SAXException {
        try {
            final XMLReader reader = XMLHelper.newXMLReader();
            reader.setContentHandler(new XSSFSheetXMLHandler(
                    styles,
                    null,
                    strings,
                    sheetHandler,
                    new CustomDataFormatter(),
                    false));
            reader.parse(new InputSource(sheetInputStream));
        } catch (ParserConfigurationException e) {
            throw new RuntimeException("sax parser appears to be broken. " + e.getMessage());
        }
    }


    private static class SheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

        private final Printer printer;
        private boolean firstCellOfRow;
        private int currentRow;
        private int currentCol;

        public SheetHandler(Printer printer) {
            this.printer = printer;
            this.firstCellOfRow = false;
            this.currentRow = -1;
            this.currentCol = -1;
        }

        @Override
        public void startRow(int rowNum) {
            for (int i = 0; i < (rowNum - currentRow - 1); i++) {
                printer.println();
            }
            firstCellOfRow = true;
            currentRow = rowNum;
            currentCol = -1;
        }

        @Override
        public void endRow(int rowNum) {
            printer.println();
        }

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            if (firstCellOfRow) {
                firstCellOfRow = false;
            }

            if (cellReference == null) {
                cellReference = new CellAddress(currentRow, currentCol).formatAsString();
            }

            int thisCol = (new CellReference(cellReference)).getCol();
            int missedCols = thisCol - currentCol - 1;
            for (int i = 0; i < missedCols; i++) {
                printer.print("");
            }
            currentCol = thisCol;
            printer.print(formattedValue);
        }
    }


    private static class CustomDataFormatter extends DataFormatter {

        private final ObjectCellFormatter formatter = new ObjectCellFormatter();

        @Override
        public String formatRawCellContents(double value, int formatIndex,
                String formatString, boolean use1904Windowing) {
            if (DateUtil.isADateFormat(formatIndex, formatString)) {
                if (DateUtil.isValidExcelDate(value)) {
                    Date d = DateUtil.getJavaDate(value, use1904Windowing);
                    return formatter.formatAsString(formatString, d);
                }
            }
            return super.formatRawCellContents(value, formatIndex, formatString, use1904Windowing);
        }
    }


    private interface Printer {
        void println();
        void print(Object val);
    }


    private static class CsvPrinter implements Printer, Flushable, Closeable {

        private CSVPrinter printer;

        public CsvPrinter(CSVPrinter printer) {
            this.printer = Objects.requireNonNull(printer);
        }

        @Override
        public void println() {
            try {
                printer.println();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }

        @Override
        public void print(Object val) {
            try {
                printer.print(val);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }

        @Override
        public void close() throws IOException {
            printer.close();
        }

        @Override
        public void flush() throws IOException {
            printer.flush();
        }
    }

}

