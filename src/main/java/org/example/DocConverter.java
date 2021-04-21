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

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.w3c.dom.Document;

public class DocConverter extends Converter {
    private HWPFDocument hwpfDocument;
    private InputStream inputStream;

    public static void main(String[] args) throws IOException {
        InputStream is = new FileInputStream("example/resume.doc");
        OutputStream os = new FileOutputStream("example/out/resume_doc.html");


        DocConverter DocConverter = new DocConverter(is);
        //convert Excel to html output stream
        DocConverter.convertToHtmlStream(os,  "UTF-8");

        is.close();
        os.close();

        System.out.println("FINISHED DocConverter");
    }

    public DocConverter(final InputStream xlsStream) {
        this.inputStream = xlsStream;
        try {
            this.hwpfDocument = new HWPFDocument(this.inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     *  @param htmlOutputStream excel path
     * @param encoding encoding format
     */
    public void convertToHtmlStream(OutputStream htmlOutputStream, String encoding) {
        printPage(htmlOutputStream, encoding, this.hwpfDocument);
    }

    private void printPage(OutputStream htmlOutputStream, String encoding, HWPFDocument hwpfDocument) {
        StringWriter writer = null;

        try {
            Document doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
            WordToHtmlConverter converter = new WordToHtmlConverter(doc);
            converter.processDocument(hwpfDocument);
            writer = getStringWriter(encoding, converter.getDocument());
            htmlOutputStream.write(writer.toString().getBytes(encoding));
            htmlOutputStream.flush();
        } catch (ParserConfigurationException | TransformerException | IOException e) {
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


}
