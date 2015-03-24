package com.nasdaqomx.jxls;

import com.nasdaqomx.jxls.domain.Person;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Hello world!
 *
 */
public class App {

    private static final Logger LOG = LoggerFactory.getLogger(App.class);

    public static void main(String[] args) throws Exception {
        List persons = new ArrayList();
        for (int a = 0; a < 20000; a++) {
            persons.add(new Person("Ola", "Theander", 5));
            persons.add(new Person("Pelle", "Jansson", 23));
        }

        useJXLS(persons);
        useApachePOI(persons);
    }

    private static void useApachePOI(List persons) {
        final org.apache.poi.xssf.usermodel.XSSFWorkbook wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook();
        final org.apache.poi.xssf.usermodel.XSSFSheet sheet = wb.createSheet("Sample sheet");

        // create 2 cell styles
        final XSSFCellStyle cs = wb.createCellStyle();
        final XSSFCellStyle cs2 = wb.createCellStyle();
        final XSSFDataFormat df = wb.createDataFormat();

        // create 2 fonts objects
        final XSSFFont f = wb.createFont();
        final XSSFFont f2 = wb.createFont();

        // Set font 1 to 12 point type, blue and bold
        f.setFontHeightInPoints((short) 12);
        f.setColor(org.apache.poi.ss.usermodel.IndexedColors.RED.getIndex());
        f.setBoldweight(org.apache.poi.ss.usermodel.Font.BOLDWEIGHT_BOLD);

        // Set font 2 to 10 point type, red and bold
        f2.setFontHeightInPoints((short) 10);
        f2.setColor(org.apache.poi.ss.usermodel.IndexedColors.BLUE.getIndex());
        f2.setBoldweight(org.apache.poi.ss.usermodel.Font.BOLDWEIGHT_BOLD);

        cs.setFont(f);
        cs2.setFont(f2);

        LOG.info("Starting to create Excel using Apache POI.");
        final long start = System.currentTimeMillis();

        int rownum = 0;
        for (Object p : persons) {
            Person person = (Person) p;
            XSSFRow row = sheet.createRow(rownum++);
            XSSFCell cell = row.createCell(0);
            cell.setCellValue(person.getFirstName());
            cell.setCellStyle(cs);
            cell = row.createCell(1);
            cell.setCellValue(person.getSurName());
            cell.setCellStyle(cs);
            cell = row.createCell(2);
            cell.setCellValue(person.getAge());
            cell.setCellStyle(cs2);
        }

        try (FileOutputStream out = new FileOutputStream(new File("C:\\new.xlsx"))) {
            wb.write(out);
            LOG.info("Excel written successfully, took {} ms.",
                    System.currentTimeMillis() - start);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void useJXLS(List persons) throws IOException, ParsePropertyException, InvalidFormatException {
        // initilize list of departments in some way
        Map beans = new HashMap();
        beans.put("persons", persons);
        XLSTransformer transformer = new XLSTransformer();
        LOG.debug("Transforming XLSTransformer.");
        final long start = System.currentTimeMillis();
        transformer.transformXLS("Person.xlsx", beans, "PersonOut.xlsx");
        LOG.info("Transform done, took {} ms.",
                System.currentTimeMillis() - start);
    }
}
