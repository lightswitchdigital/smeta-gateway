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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicReference;

public class PDFExporter {
    public static final String FONT = System.getProperty("user.dir") + "/src/fonts/calibri.ttf";
    Logger logger = LoggerFactory.getLogger(WorkbooksPool.class);
    public static final boolean CLEAR_ZEROS = true;

    public void smetaZak(String path, FormulaEvaluator evaluator, Sheet sheet, ArrayList<ArrayList<String>> rows) throws IOException {

        // |----------------------------------------------
        // | 1 - Всякая хуйня + сабтайтлы материалы и работы
        // | 2 - Тайтлы, сабтайтлы материалы и работы
        // | 3 - непонятно
        // | 4 - строки материалы и работы

        DecimalFormat df = new DecimalFormat("0.00");

//        PdfWriter writer = new PdfWriter(System.getProperty("user.dir") + "/smeta_zak.pdf");
        PdfWriter writer = new PdfWriter(Paths.get(path + "smeta_zak.pdf").toString());

        PdfDocument pdfDoc = new PdfDocument(writer);
        pdfDoc.setDefaultPageSize(PageSize.A4.rotate());

        Document doc = new Document(pdfDoc);

        PdfFont font = PdfFontFactory.createFont(FONT, PdfEncodings.IDENTITY_H);

        doc.setFont(font);
        doc.add(new Paragraph("Сметный расчет для зак укр").setFontSize(14).setBold().setTextAlignment(TextAlignment.CENTER));
        doc.add(new Paragraph("\n"));

        //////////////////////////
        // Adding header information
        Table tableMeta = new Table(new float[]{150f, 100f});
        tableMeta.addCell("Длина дома, м.п.").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMeta.addCell(df.format(this.getCellValue(evaluator, sheet, "F3")));
        tableMeta.addCell("Ширина дома, м.п.").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMeta.addCell(df.format(this.getCellValue(evaluator, sheet, "F4")));
        tableMeta.addCell("Этажность дома").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMeta.addCell(df.format(this.getCellValue(evaluator, sheet, "F5")));
        tableMeta.addCell("S строения общая, м2").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMeta.addCell(df.format(this.getCellValue(evaluator, sheet, "F6")));
        tableMeta.addCell("S строения чистая, м2").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMeta.addCell(df.format(this.getCellValue(evaluator, sheet, "F7")));
        doc.add(tableMeta);

        doc.add(new Paragraph("\n"));

        ArrayList<ArrayList<String>> cleared = new ArrayList<>();
        rows.forEach(row -> {
            if (row.size() == 4) {

                // Clear all zeroed rows
                if (CLEAR_ZEROS) {
                    try {
                        double price = Double.parseDouble(row.get(3));
                        if (price < 0.00001) {
                            return;
                        }
                    } catch (NumberFormatException e) {
                        return;
                    }
                }
            }
            cleared.add(row);
        });

        ////////////////////////////////
        // Removing all excessive titles

        if (CLEAR_ZEROS) {
            ArrayList<ArrayList<String>> toRemove = new ArrayList<>();
            ZakSmetaStates inState = ZakSmetaStates.TITLE;
            for (int i = 0; i < cleared.size(); i++) {
                ArrayList<String> row = cleared.get(i);
                if (row.size() == 1 || row.size() == 2) {
                    if (row.get(0).startsWith("Итого")) {
                        inState = ZakSmetaStates.TOTAL;
                    } else if (!Objects.equals(row.get(0), "Работы") && !Objects.equals(row.get(0), "Материалы")) {
                        inState = ZakSmetaStates.TITLE;
                    } else {
                        inState = ZakSmetaStates.SUBTITLE;
                    }
                } else if (row.size() == 4) {
                    inState = ZakSmetaStates.DATA;
                }

                // Now we are checking if prev state was title or subtitle
//                if (inState == ZakSmetaStates.TOTAL) {
//                    if (cleared.get(i-1).size() == 2 && cleared.get(i-2).size() == 2) {
//                        toRemove.add(row);
//                        toRemove.add(cleared.get(i-1));
//                        toRemove.add(cleared.get(i-2));
//                    } else if (cleared.get(i - 1).size() == 2) {
//                        toRemove.add(row);
//                        toRemove.add(cleared.get(i-1));
//                    }
//                }
            }

            toRemove.forEach(System.out::println);
            toRemove.forEach(cleared::remove);
        }

        AtomicReference<ZakSmetaStates> state = new AtomicReference<>(ZakSmetaStates.TITLE);
        AtomicReference<Table> bufTable = new AtomicReference<>(new Table(new float[]{100f, 50f, 50f, 50f}));

        cleared.forEach(row -> {

            if (row.size() == 2 || row.size() == 1) {

                // If the prev state was materials or works - we close buf table
                if (state.get() == ZakSmetaStates.DATA) {
                    doc.add(bufTable.get());
                }

                if (row.get(0).startsWith("Итого")) {
                    state.set(ZakSmetaStates.TOTAL);
                    Table totalTable = new Table(new float[]{150f, 100f});
                    totalTable.addCell(row.get(0)).setBold();
                    totalTable.addCell(row.get(1));
                    doc.add(totalTable);
                } else if (!Objects.equals(row.get(0), "Работы") && !Objects.equals(row.get(0), "Материалы")) {
                    state.set(ZakSmetaStates.TITLE);
                    doc.add(new Paragraph("\n"));
                    doc.add(new Paragraph(row.get(0)).setFontSize(18).setBold());
                } else {
                    state.set(ZakSmetaStates.SUBTITLE);
                    doc.add(new Paragraph(row.get(0)).setFontSize(14).setBold());
                }

                // Works and materials
            } else if (row.size() == 4) {
//
                // Checking for validity
                if (row.get(0) == null) {
                    return;
                }

                float[] cols = {100f, 50f, 50f, 50f};

                if (state.get() == ZakSmetaStates.TITLE || state.get() == ZakSmetaStates.SUBTITLE) {
                    Table t = new Table(cols);
                    addTableHeaders4(t);
                    bufTable.set(t);
                }
                bufTable.get().addCell(row.get(0));
                bufTable.get().addCell(row.get(1));
                try {
                    double price1 = Double.parseDouble(row.get(2));
                    double price2 = Double.parseDouble(row.get(3));

                    bufTable.get().addCell(df.format(price1));
                    bufTable.get().addCell(df.format(price2));
                } catch (NumberFormatException ee) {
                    return;
                }

                state.set(ZakSmetaStates.DATA);
            }
        });


        //////////////////////////
        // Adding footer information
        Table tableMetaFooter = new Table(new float[]{150f, 100f});
        tableMetaFooter.addCell("Всего материалов, рублей").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "F2348")));
        tableMetaFooter.addCell("Всего работы, рублей").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "F2349")));
        tableMetaFooter.addCell("Всего транспортные расходы, рублей").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "F2350")));
        tableMetaFooter.addCell("Всего дополнительные расходные материалы").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "F2351")));
        tableMetaFooter.addCell("Всего работ и материалов, рублей").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "F2352")));
        doc.add(tableMetaFooter);

        doc.close();
        writer.close();
    }

    private Cell getCell(Sheet sheet, String id) {
        CellReference cr = new CellReference(id);
        Row row = sheet.getRow(cr.getRow());

        return row.getCell(cr.getCol());
    }

    private Double getCellValue(FormulaEvaluator evaluator, Sheet sheet, String id) {
        Cell cell = this.getCell(sheet, id);

        evaluator.clearAllCachedResultValues();
        return evaluator.evaluate(cell).getNumberValue();
    }

    public void smetaInternal(String path, ArrayList<ArrayList<String>> rows) throws IOException {

//        rows.forEach(row -> {
//            if (row.size() == 4) {
//                System.out.println(row);
//            }
//        });
//
//        if (true) {
//            return;
//        }

        // |--------------------------------------------------
        // | 3 - Названия групп, в том числе материалы, работы
        // | 4 - Итого работы
        // | 5 - итого материалы
        // | 6
        // | 7 - Работы
        // | 8 - Материалы

//        PdfWriter writer = new PdfWriter(System.getProperty("user.dir") + "/smeta_internal.pdf");
        PdfWriter writer = new PdfWriter(Paths.get(path + "smeta_internal.pdf").toString());

        PdfDocument pdfDoc = new PdfDocument(writer);
        pdfDoc.setDefaultPageSize(PageSize.A4.rotate());

        Document doc = new Document(pdfDoc);

        PdfFont font = PdfFontFactory.createFont(FONT, PdfEncodings.IDENTITY_H);

        doc.setFont(font);
        doc.add(new Paragraph("Сметный расчет (внутренний)").setFontSize(14).setBold().setTextAlignment(TextAlignment.CENTER));
        doc.add(new Paragraph("\n"));


        // First we clear out all priceless fuck
        ArrayList<ArrayList<String>> cleared = new ArrayList<>();
        rows.forEach(row -> {
            if (row.size() == 7 || row.size() == 8) {
                if (CLEAR_ZEROS) {
                    try {
                        double price = Double.parseDouble(row.get(6));
                        if (price < 0.0001) {
                            return;
                        }

                    } catch (NumberFormatException e) {
                        return;
                    }
                }
            }
            cleared.add(row);
        });

        AtomicReference<InternalSmetaStates> state = new AtomicReference<>(InternalSmetaStates.TITLE);
        AtomicReference<Table> bufTable = new AtomicReference<>(new Table(new float[]{50f, 20f, 20f, 20f, 20f, 20f, 20f}));

//        AtomicInteger counter = new AtomicInteger();

        cleared.forEach(row -> {

//            counter.getAndIncrement();

            if (row.size() == 3) {
                // If the prev state was materials or works - we close buf table
                if (state.get() == InternalSmetaStates.MATERIALS || state.get() == InternalSmetaStates.WORKS) {
                    doc.add(bufTable.get());
                }

                if (!Objects.equals(row.get(0), "Работы") && !Objects.equals(row.get(0), "Материалы") && !Objects.equals(row.get(0), "Техника и дополнительные расходы")) {
                    state.set(InternalSmetaStates.TITLE);
                    doc.add(new Paragraph("\n"));
                    doc.add(new Paragraph(row.get(0)).setFontSize(18).setBold());
                } else {
                    state.set(InternalSmetaStates.SUBTITLE);
                    doc.add(new Paragraph(row.get(0)).setFontSize(14).setBold());
                }

                // Works
            }else if(row.size() == 7) {

                float[] worksColWidth = {50f, 20f, 20f, 20f, 20f, 20f, 20f};

                if (state.get() == InternalSmetaStates.SUBTITLE || state.get() == InternalSmetaStates.TITLE) {
                    Table t = new Table(worksColWidth);
                    addTableHeaders7(t);
                    bufTable.set(t);
                }
                bufTable.get().addCell(row.get(0));
                bufTable.get().addCell(row.get(1));
                bufTable.get().addCell(row.get(2));
                bufTable.get().addCell(row.get(3));
                bufTable.get().addCell(row.get(4));
                bufTable.get().addCell(row.get(5));
                bufTable.get().addCell(row.get(6));
//                row.forEach(bufTable.get()::addCell);

                state.set(InternalSmetaStates.WORKS);

                // Materials
//            }
            } else if (row.size() == 8) {
//
//                // Checking for validity
                if (row.get(0) == null) {
                    return;
                }
//
                // If the first col is can be parsed into double, we emerge
//                try {
//                    Double.parseDouble(row.get(0));
//                    logger.error("could not parse table row name: " + row.get(0));
//                }catch (NumberFormatException e) {
                float[] materialsColWidth = {50f, 20f, 20f, 20f, 20f, 20f, 20f, 20f};
                if (state.get() == InternalSmetaStates.SUBTITLE || state.get() == InternalSmetaStates.TITLE) {
                    Table t = new Table(materialsColWidth);
//                        addTableHeaders8(t);
                    bufTable.set(t);
                }

                row.forEach(bufTable.get()::addCell);
                state.set(InternalSmetaStates.MATERIALS);
//                }
            }
        });

        //////////////////
        // Adding footer


        doc.close();
        writer.close();
    }

    public void addTableHeaders4(Table table) {
        table.addHeaderCell(new Paragraph("Наименование работ и затрат").setBold());
        table.addHeaderCell(new Paragraph("ед. изм").setBold());
        table.addHeaderCell(new Paragraph("кол-во").setBold());
        table.addHeaderCell(new Paragraph("Цена ед. изм.").setBold());
    }


    public void addTableHeaders8(Table table) {
        table.addHeaderCell("Наименование работ и затрат").setBold();
        table.addHeaderCell("ед. изм").setBold();
        table.addHeaderCell("кол-во").setBold();
        table.addHeaderCell("Цена ед.изм.").setBold();
        table.addHeaderCell("Цена с накруткой").setBold();
        table.addHeaderCell("Общая цена").setBold();
        table.addHeaderCell("Цена по прайсу").setBold();
        table.addHeaderCell("Для закупа материала").setBold();
    }

    public void addTableHeaders7(Table table) {
        table.addHeaderCell(new Paragraph("Наименование работ и затрат").setBold());
        table.addHeaderCell(new Paragraph("ед. изм").setBold());
        table.addHeaderCell(new Paragraph("кол-во").setBold());
        table.addHeaderCell(new Paragraph("Цена ед.изм.").setBold());
        table.addHeaderCell(new Paragraph("Цена с накруткой").setBold());
        table.addHeaderCell(new Paragraph("Общая цена").setBold());
        table.addHeaderCell(new Paragraph("Цена по прайсу").setBold());
    }
}
