package org.example;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import com.itextpdf.html2pdf.ConverterProperties;
import com.itextpdf.html2pdf.HtmlConverter;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;

public class PdfConverter {

    public static void main(String[] args) throws IOException {
        xlsDemo(new FileInputStream("example/test.xls"), new FileOutputStream("example/out/test_xls.pdf"));
        xlsxDemo(new FileInputStream("example/test.xlsx"), new FileOutputStream("example/out/test_xlsx.pdf"));


    }

    private static void xlsDemo(InputStream is, OutputStream os) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();


        XlsConverter xlsConverter = new XlsConverter(is);
        xlsConverter.convertToHtmlStream(baos,  "UTF-8"); //closes input stream automatically

        //create a pdf document and set it's page size
        writeStreamToPdf(is, os, baos);
    }

    private static void xlsxDemo(InputStream is, OutputStream os) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        XlsxConverter converter = new XlsxConverter(is);

        converter.convertToHtmlStream(baos);
        writeStreamToPdf(is, os, baos);

    }

    private static void writeStreamToPdf(InputStream is, OutputStream os, ByteArrayOutputStream baos) throws IOException {
        //create a pdf document and set it's page size
        PdfDocument doc = new PdfDocument(new PdfWriter(os));
        doc.setDefaultPageSize(new PageSize(3000, 2000));


        //convert input stream to a
        ConverterProperties props = new ConverterProperties();
        HtmlConverter.convertToPdf(String.valueOf(baos), doc, props);

        baos.close();
        doc.close();
        is.close();
        os.close();
    }
}
