package com.lightswitch.ramdom.smeta;

import com.itextpdf.io.font.PdfEncodings;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.property.TextAlignment;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.stream.Stream;

public class PDFExporter {
    public static final String FONT = System.getProperty("user.dir") + "/src/fonts/calibri.ttf";
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
//            return row.size() == 8;
//        }).forEach(row -> {
//            System.out.println(row);
//        });
//
//        boolean a = true;
//        if (a == true) {
//            return;
//        }

        PdfWriter writer = new PdfWriter(System.getProperty("user.dir") + "/smeta.pdf");
        System.out.println("PDF created.");

        PdfDocument pdfDoc = new PdfDocument(writer);
        pdfDoc.setDefaultPageSize(PageSize.A3.rotate());

        Document doc = new Document(pdfDoc);
//            PdfFont font = PdfFontFactory.createFont(FontConstants.HELVETICA_BOLD);

        PdfFont font = PdfFontFactory.createFont(FONT, PdfEncodings.IDENTITY_H);

        doc.setFont(font);
        doc.add(new Paragraph("Сметный расчет (внутренний)").setFontSize(14).setBold().setTextAlignment(TextAlignment.CENTER));
        doc.add(new Paragraph("\n"));

        float[] colWidth = {50f, 20f, 20f, 20f, 20f, 20f, 20f, 20f};
        Table table = new Table(colWidth);


////        table.setFont(font);
//        table.addHeaderCell("Наименование работ и затрат");
//        table.addHeaderCell("ед. изм");
//        table.addHeaderCell("кол-во");
//        table.addHeaderCell("Цена ед.изм.");
//        table.addHeaderCell("Цена с накруткой");
//        table.addHeaderCell("Общая цена");
//        table.addHeaderCell("Цена по прайсу");
//        table.addHeaderCell("Для закупа материала");
//
//        doc.add(table);

        addTableHeaders(doc, table);

        AtomicReference<InternalSmetaStates> state = new AtomicReference<>(InternalSmetaStates.TITLE);
        AtomicReference<Table> bufTable = new AtomicReference<>(new Table(new float[]{50f, 20f, 20f, 20f, 20f, 20f, 20f}));

        rows.forEach(row -> {

            if (row.size() == 3) {
                // If the prev state was materials or works - we close buf table
                if (state.get() == InternalSmetaStates.MATERIALS || state.get() == InternalSmetaStates.WORKS) {
                    doc.add(bufTable.get());
                }

                if (!Objects.equals(row.get(0), "Работы") && !Objects.equals(row.get(0), "Материалы")) {
                    state.set(InternalSmetaStates.TITLE);
                    doc.add(new Paragraph("\n"));
                    doc.add(new Paragraph(row.get(0)).setFontSize(18).setBold());
                } else {
                    state.set(InternalSmetaStates.SUBTITLE);
                    doc.add(new Paragraph(row.get(0)).setFontSize(14).setBold());
                }

            }else if(row.size() == 7) {
                float[] worksColWidth = {50f, 20f, 20f, 20f, 20f, 20f, 20f};

                if (state.get() == InternalSmetaStates.SUBTITLE || state.get() == InternalSmetaStates.TITLE) {
                    Table t = new Table(worksColWidth);
//                    addTableHeaders(doc, t);
                    bufTable.set(t);
                }
                row.forEach(bufTable.get()::addCell);

                state.set(InternalSmetaStates.WORKS);
//                Table worksTable = new Table(worksColWidth);
//                row.forEach(worksTable::addCell);
//
//                doc.add(worksTable);
            } else if (row.size() == 8) {
                System.out.println(row);
                float[] materialsColWidth = {50f, 20f, 20f, 20f, 20f, 20f, 20f, 20f};
                if (state.get() == InternalSmetaStates.SUBTITLE || state.get() == InternalSmetaStates.TITLE) {
                    Table t = new Table(materialsColWidth);
//                    addTableHeaders(doc, t);
                    bufTable.set(t);
                }

//                Table materialsTable = new Table(materialsColWidth);

                try {
                    row.forEach(bufTable.get()::addCell);
//                    row.forEach(materialsTable::addCell);
                }catch (IllegalArgumentException e) {
//                    materialsTable.addCell(" ");
                    bufTable.get().addCell(" ");
                }

                state.set(InternalSmetaStates.MATERIALS);

//                doc.add(materialsTable);
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

    public void addTableHeaders(Document doc, Table table) {
//        float[] colWidth = {50f, 20f, 20f, 20f, 20f, 20f, 20f, 20f};

//        table.setFont(font);
        table.addHeaderCell(new Paragraph("Наименование работ и затрат").setBold());
        table.addHeaderCell(new Paragraph("ед. изм").setBold());
        table.addHeaderCell(new Paragraph("кол-во").setBold());
        table.addHeaderCell(new Paragraph("Цена ед.изм.").setBold());
        table.addHeaderCell(new Paragraph("Цена с накруткой").setBold());
        table.addHeaderCell(new Paragraph("Общая цена").setBold());
        table.addHeaderCell(new Paragraph("Цена по прайсу").setBold());
        table.addHeaderCell(new Paragraph("Для закупа материала").setBold());

        doc.add(table);
    }
}
