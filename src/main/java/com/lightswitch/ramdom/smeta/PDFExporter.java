package com.lightswitch.ramdom.smeta;

import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.util.ArrayList;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;

public class PDFExporter {
    public static final String FONT = System.getProperty("user.dir") + "/src/fonts/Arial.ttf";
    Logger logger = LoggerFactory.getLogger(WorkbooksPool.class);

    public void smetaInternal(Stream<ArrayList<String>> rows) throws IOException {

        // |--------------------------------------------------
        // | 3 - Названия групп, в том числе материалы, работы
        // | 4 - Итого работы
        // | 5 - Тоже итого материалы
        // | 6
        // | 7 - Работы
        // | 8 - Материалы

//        rows.filter(row -> {
//            return row.size() == 3;
//        }).forEach(System.out::println);
//
//        boolean a = true;
//        if (a == true) {
//            return;
//        }

        PdfWriter writer = new PdfWriter(System.getProperty("user.dir") + "/test.pdf");
        System.out.println("PDF created.");

        PdfDocument pdfDoc = new PdfDocument(writer);
        pdfDoc.setDefaultPageSize(PageSize.A3.rotate());

        Document doc = new Document(pdfDoc);
//            PdfFont font = PdfFontFactory.createFont(FontConstants.HELVETICA_BOLD);

        PdfFont font = PdfFontFactory.createFont(FONT);

        doc.setFont(font);
        doc.add(new Paragraph("Smeta internal"));

        float[] colWidth = {50f, 20f, 20f, 20f, 20f, 20f, 20f, 20f};
        Table table = new Table(colWidth);

        table.addHeaderCell("Naimenovanie");
        table.addHeaderCell("Ed izm");
        table.addHeaderCell("Qnt");
        table.addHeaderCell("Price per");
        table.addHeaderCell("Price with over");
        table.addHeaderCell("Total");
        table.addHeaderCell("Total by price");
        table.addHeaderCell("For merch");

        doc.add(table);

        AtomicInteger counter = new AtomicInteger();
        rows.forEach(row -> {
            float[] cols = {350f};

            if (row.size() == 3) {
                counter.getAndIncrement();
                if (row.get(0) != "Работы" && row.get(0) != "Материалы") {
                    table.addHeaderCell(row.get(0));
                } else {

                }
                Table tableHeader = new Table(cols);
                tableHeader.addCell(new Paragraph(row.get(0)));
                doc.add(tableHeader);
            }
        });

        // Adding headers

//        rows.forEach(row -> {
//            row.forEach(table::addCell);
//        });

//        doc.add(table);
        doc.close();
        writer.close();
    }
}
