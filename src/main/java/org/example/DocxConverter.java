package org.example;

import java.io.BufferedReader;
import java.io.Closeable;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.nio.charset.StandardCharsets;
import java.util.Formatter;
import java.util.Locale;

import org.apache.commons.codec.CharEncoding;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;

public class DocxConverter extends Converter {
    private InputStream inputStream;
    private XWPFDocument xwpfDocument = null;
    private Appendable output = null;
    private Formatter out;
    private boolean completeHTML;


    public static void main(String[] args) throws IOException {

        DocxConverter docxConverter = new DocxConverter(new FileInputStream("example/resume.docx"));
        docxConverter.convertToHtmlStream(new FileOutputStream("example/out/resume_docx.html"), CharEncoding.UTF_8);

    }

    public DocxConverter(final InputStream docxStream) {
        this.inputStream = docxStream;
    }

    public static DocxConverter create(InputStream in, Appendable output)
            throws IOException, InvalidFormatException {
        XWPFDocument xwpfDocument = new XWPFDocument(OPCPackage.open(in));
        return create(xwpfDocument, output);
    }

    public static DocxConverter create(XWPFDocument xwpfDocument, Appendable output) {
        return new DocxConverter(xwpfDocument, output);
    }

    private DocxConverter(XWPFDocument xwpfDocument, Appendable output) {
        if (xwpfDocument == null) {
            throw new NullPointerException("xwpfDocument");
        }
        if (output == null) {
            throw new NullPointerException("output");
        }
        this.xwpfDocument = xwpfDocument;
        this.output = output;
    }

    public void convertToHtmlStream(OutputStream htmlOutputStream, String encoding) throws IOException {
        try (PrintWriter pw = new PrintWriter(htmlOutputStream)) {
            DocxConverter toHtml = create(this.inputStream, pw);
            toHtml.setCompleteHTML(true);
            toHtml.printPage();

        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    public void setCompleteHTML(boolean completeHTML) {
        this.completeHTML = completeHTML;
    }
    public void printPage() {
        try {
            ensureOut();
            if (completeHTML) {
                out.format(
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" ?>%n");
                out.format("<html>%n");
                out.format("<head>%n");
                printInlineStyle();
                out.format("</head>%n");
                out.format("<body>%n");
            }

            print();

            if (completeHTML) {
                out.format("</body>%n");
                out.format("</html>%n");
            }
        } finally {
            IOUtils.closeQuietly(out);
            if (output instanceof Closeable) {
                IOUtils.closeQuietly((Closeable) output);
            }
        }
    }

    private void ensureOut() {
        if (out == null) {
            out = new Formatter(output, Locale.ROOT);
        }
    }

    public void print() {
        printXwpfDocument();
    }

    private void printXwpfDocument() {
        ensureOut();

        printBody();

    }

    private void printInlineStyle() {
        //out.format("<link href=\"excelStyle.css\" rel=\"stylesheet\" type=\"text/css\">%n");
        out.format("<style type=\"text/css\">%n");
        printStyles();
        out.format("</style>%n");
    }

    public void printStyles() {
        ensureOut();

        // First, copy the base css
        try (BufferedReader in = new BufferedReader(new InputStreamReader(
                getClass().getClassLoader().getResourceAsStream("excelStyle.css"), StandardCharsets.ISO_8859_1))){
            String line;
            while ((line = in.readLine()) != null) {
                out.format("%s%n", line);
            }
        } catch (IOException e) {
            throw new IllegalStateException("Reading standard css", e);
        }

//
//        // now add css for each used style
//        Set<CellStyle> seen = new HashSet<>();
//        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
//            Sheet sheet = wb.getSheetAt(i);
//            Iterator<Row> rows = sheet.rowIterator();
//            while (rows.hasNext()) {
//                Row row = rows.next();
//                for (Cell cell : row) {
//                    CellStyle style = cell.getCellStyle();
//                    if (!seen.contains(style)) {
//                        printStyle(style);
//                        seen.add(style);
//                    }
//                }
//            }
//        }
    }

    public void printBody() {
        ensureOut();
        for(IBodyElement bodyElement: this.xwpfDocument.getBodyElements()) {
            switch(bodyElement.getElementType()) {
                case PARAGRAPH:
                    XWPFParagraph paragraph = (XWPFParagraph)bodyElement;
                    System.out.println(paragraph.getText());
                    System.out.println(paragraph.getElementType().toString());
                    CTPPr pr = paragraph.getCTP().getPPr();
                    if(pr != null && pr.isSetPStyle()) {
                        if (paragraph.getStyle().equalsIgnoreCase("BodyContentStyle")) {
                            out.format("<div style=\"font-weight: normal; color: black; margin: 5px 0px 5px 0px\">%s</div>%n", paragraph.getText());
                        } else if (paragraph.getStyle().equalsIgnoreCase("IntroductionStyle")) {
                            out.format("<div style=\"color: chocolate; font-weight: bolder; font-size: larger;margin: 10px 0px 10px 0px;\">%s</div>%n", paragraph.getText());
                        }
                    } else {
                        out.format("<div>%s</div>%n", paragraph.getText());
                    }
                    break;
                case TABLE:
                    System.out.println("Is Table");
                    XWPFTable table = (XWPFTable)bodyElement;
                    System.out.printf("Rows: %d%n", table.getNumberOfRows());
                    break;
                case CONTENTCONTROL:
                    System.out.println("is Content Control");
                default:
            }

        }
//        for (XWPFParagraph paragraph : this.xwpfDocument.getParagraphs()) {
//            System.out.println(paragraph.getText());
//            System.out.println(paragraph.getElementType().toString());
//            if (paragraph.getStyle().equalsIgnoreCase("BodyContentStyle")) {
//                out.format("<div style=\"font-weight: normal; color: black; margin: 5px 0px 5px 0px\">%s</div>%n", paragraph.getText());
//            } else if (paragraph.getStyle().equalsIgnoreCase("IntroductionStyle")) {
//                out.format("<div style=\"color: chocolate; font-weight: bolder; font-size: larger;margin: 10px 0px 10px 0px;\">%s</div>%n", paragraph.getText());
//            }
//
//
//            System.out.println(paragraph.getStyle());
//
//        }
    }
}
