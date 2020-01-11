package cn.zhh;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import java.io.InputStream;
import java.lang.invoke.MethodHandles;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Logger;
import java.util.stream.Collectors;

/**
 * POI 导入
 *
 * @author Zhou Huanghua
 * @date 2020/1/11 14:02
 */
public class PoiImport {

    private static final Logger LOGGER = Logger.getLogger(MethodHandles.lookup().lookupClass().getName());

    public static void main(String[] args) throws Exception {
        String filePath = "C:/Users/dell/Desktop/tmp/demo.xlsx";
        // OPCPackage.open(...)有多个重载方法，比如String path、File file、InputStream in等
        try (OPCPackage opcPackage = OPCPackage.open(filePath);) {
            // 创建XSSFReader读取StylesTable和ReadOnlySharedStringsTable
            XSSFReader xssfReader = new XSSFReader(opcPackage);
            StylesTable stylesTable = xssfReader.getStylesTable();
            ReadOnlySharedStringsTable sharedStringsTable = new ReadOnlySharedStringsTable(opcPackage);
            // 创建XMLReader，设置ContentHandler
            XMLReader xmlReader = SAXHelper.newXMLReader();
            xmlReader.setContentHandler(new XSSFSheetXMLHandler(stylesTable, sharedStringsTable, new SimpleSheetContentsHandler(), false));
            // 解析每个Sheet数据
            Iterator<InputStream> sheetsData = xssfReader.getSheetsData();
            while (sheetsData.hasNext()) {
                try (InputStream inputStream = sheetsData.next();) {
                    xmlReader.parse(new InputSource(inputStream));
                }
            }
        }
    }

    /**
     * 内容处理器
     */
    public static class SimpleSheetContentsHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

        protected List<String> row;

        /**
         * A row with the (zero based) row number has started
         *
         * @param rowNum
         */
        @Override
        public void startRow(int rowNum) {
            row = new ArrayList<>();
        }

        /**
         * A row with the (zero based) row number has ended
         *
         * @param rowNum
         */
        @Override
        public void endRow(int rowNum) {
            if (row.isEmpty()) {
                return;
            }
            // 处理数据
            LOGGER.info(row.stream().collect(Collectors.joining("   ")));
        }

        /**
         * A cell, with the given formatted value (may be null),
         * and possibly a comment (may be null), was encountered
         *
         * @param cellReference
         * @param formattedValue
         * @param comment
         */
        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            row.add(formattedValue);
        }

        /**
         * A header or footer has been encountered
         *
         * @param text
         * @param isHeader
         * @param tagName
         */
        @Override
        public void headerFooter(String text, boolean isHeader, String tagName) {
        }
    }
}
