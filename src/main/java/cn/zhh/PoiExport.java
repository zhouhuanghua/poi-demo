package cn.zhh;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.invoke.MethodHandles;
import java.util.logging.Logger;

import static java.lang.String.format;

/**
 * POI 导出
 *
 * @author Zhou Huanghua
 * @date 2020/1/11 14:03
 */
public class PoiExport {

    private static final Logger LOGGER = Logger.getLogger(MethodHandles.lookup().lookupClass().getName());

    public static void main(String[] args) throws Exception {
        long begin = System.currentTimeMillis();
        // keep 100 rows in memory, exceeding rows will be flushed to disk
        try (SXSSFWorkbook wb = new SXSSFWorkbook(100);
             OutputStream os = new FileOutputStream("C:/Users/dell/Desktop/tmp/demo.xlsx")) {
            Sheet sh = wb.createSheet();
            String val = "第%s行第%s列";
            for (int rowNum = 0; rowNum < 100_0000; rowNum++) {
                Row row = sh.createRow(rowNum);
                int realRowNum = rowNum + 1;
                Cell cell1 = row.createCell(0);
                cell1.setCellValue(format(val, realRowNum, 1));
                Cell cell2 = row.createCell(1);
                cell2.setCellValue(format(val, realRowNum, 2));
                Cell cell3 = row.createCell(2);
                cell3.setCellValue(format(val, realRowNum, 3));
                Cell cell4 = row.createCell(3);
                cell4.setCellValue(format(val, realRowNum, 4));
            }
            wb.write(os);
        }
        LOGGER.info("导出100W行数据耗时（秒）：" + (System.currentTimeMillis() - begin)/1000);
    }
}
