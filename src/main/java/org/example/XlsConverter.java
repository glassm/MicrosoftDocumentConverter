package org.example;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.StringWriter;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;

import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class XlsConverter extends Converter {
    private HSSFWorkbook wb;
    private InputStream inputStream;

    public static void main(String[] args) throws IOException {
        InputStream is = new FileInputStream("example/test.xls");
        OutputStream os = new FileOutputStream("example/out/test_xls.html");


        XlsConverter xlsConverter = new XlsConverter(is);
        //convert Excel to html output stream
        xlsConverter.convertToHtmlStream(os,  "UTF-8");

    }
    
    public XlsConverter(final InputStream xlsStream) {
        this.inputStream = xlsStream;
        try {
            this.wb = new HSSFWorkbook(this.inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    /**
     *  @param htmlOutputStream excel path
     * @param encoding encoding format
     */
    public void convertToHtmlStream(OutputStream htmlOutputStream, String encoding) {
        printPage(htmlOutputStream, encoding, this.wb);
    }

    private void printPage(OutputStream htmlOutputStream, String encoding, HSSFWorkbook wb) {
        StringWriter writer = null;
        try {
            ExcelToHtmlConverter converter = getExcelToHtmlConverter(wb);
            writer = getStringWriter(encoding, converter.getDocument());
            htmlOutputStream.write(writer.toString().getBytes(encoding));
            htmlOutputStream.flush();
        } catch (ParserConfigurationException e) {
            e.printStackTrace();
        } catch (TransformerException | IOException e) {
            e.printStackTrace();
        } finally {
            if (writer != null) {
                try {
                    writer.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private ExcelToHtmlConverter getExcelToHtmlConverter(HSSFWorkbook workBook) throws ParserConfigurationException {
        ExcelToHtmlConverter converter = new ExcelToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
        converter.setOutputColumnHeaders(false);
        converter.setOutputRowNumbers(false);
        converter.processWorkbook(workBook);
        return converter;
    }

}
