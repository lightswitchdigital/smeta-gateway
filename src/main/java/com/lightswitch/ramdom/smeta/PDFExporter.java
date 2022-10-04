package com.lightswitch.ramdom.smeta;

import com.itextpdf.io.font.PdfEncodings;
import com.itextpdf.io.image.ImageData;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.color.Color;
import com.itextpdf.kernel.color.DeviceRgb;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.AreaBreak;
import com.itextpdf.layout.element.Image;
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
import java.net.MalformedURLException;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicReference;

public class PDFExporter {
    public static final String FONT = System.getProperty("user.dir") + "/src/fonts/calibri.ttf";
    Logger logger = LoggerFactory.getLogger(WorkbooksPool.class);
    public static final boolean CLEAR_ZEROS = true;

    private static final DeviceRgb GRAY_COLOR = new DeviceRgb(50, 50, 50);

    public void smetaZak(String path, FormulaEvaluator evaluator, Sheet sheet, ArrayList<ArrayList<String>> rows) throws IOException {

        // |----------------------------------------------
        // | 1 - Всякая хуйня + сабтайтлы материалы и работы
        // | 2 - Тайтлы, сабтайтлы материалы и работы
        // | 3 - непонятно
        // | 4 - строки материалы и работы


        DecimalFormat df = new DecimalFormat("0.00");

        PdfWriter writer = new PdfWriter(Paths.get(path, "smeta_zak.pdf").toString());

        PdfDocument pdfDoc = new PdfDocument(writer);
        pdfDoc.setDefaultPageSize(PageSize.A4.rotate());

        Document doc = new Document(pdfDoc);

        PdfFont font = PdfFontFactory.createFont(FONT, PdfEncodings.IDENTITY_H);
        doc.setFont(font);

        ImageData logoData = ImageDataFactory.create(Paths.get(System.getProperty("user.dir"), "/src/img/logo.png").toString());
        Image logoImg = new Image(logoData);
        logoImg.setWidth(75);

        this.addHeaderPage(doc, "Сметный расчет для закупок", logoImg);

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

        ArrayList<ArrayList<String>> cleared = new ArrayList<>();
        rows.forEach(row -> {
            if (row.size() == 4) {
                // Clear all zeroed rows
                try {
                    double price = Double.parseDouble(row.get(3));
                    if (price < 10) {
                        return;
                    }
                } catch (NumberFormatException e) {
                    return;
                }
            } else if (row.size() == 2) {
                if (row.get(0).startsWith("Итого")) {
                    return;
                }
            }
            cleared.add(row);
        });

        // А теперь сука удаляем ненужные хедеры блять
        ArrayList<ArrayList<String>> toDelete = new ArrayList<>();

        for (int i = 0; i < cleared.size(); i++) {
            ArrayList<String> row = cleared.get(i);

            // Если это большой тайтл, то проверяем следующие 3
            if (row.size() == 2 || row.size() == 1 || row.size() == 3) {

                if (!Objects.equals(row.get(0), "Работы") && !Objects.equals(row.get(0), "Материалы") && !Objects.equals(row.get(0), "Техника и дополнительные расходы")) {

                    int headersSize = 0;
                    ArrayList<String> currentRow;

                    boolean safeToDelete = true;

                    // Пытаемся найти блять другой тайтл в пределах 4 элементов
                    for (int j = 1; j < 4; j++) {
                        try {
                            currentRow = cleared.get(i + j);
                            if ((currentRow.size() == 2 || currentRow.size() == 1 || currentRow.size() == 3)
                                    && !Objects.equals(currentRow.get(0), "Работы")
                                    && !Objects.equals(currentRow.get(0), "Материалы")) {
                                headersSize = j;
                            } else if (currentRow.size() == 4) {
                                safeToDelete = false;
                            }
                        } catch (IndexOutOfBoundsException ignored) {

                        }
                    }

                    if (headersSize > 0 && safeToDelete) {
                        for (int j = 0; j < headersSize; j++) {
                            ArrayList<String> rowToDelete = cleared.get(i + j);
                            toDelete.add(rowToDelete);
                        }
                    }
                }
            }
        }

        cleared.removeAll(toDelete);

        // Если последний элемент это тайтл - то удаляем нахуй
        if (cleared.get(cleared.size() - 1).size() == 2 || cleared.get(cleared.size() - 1).size() == 3) {
            cleared.remove(cleared.size() - 1);
        }

        AtomicReference<ZakSmetaStates> state = new AtomicReference<>(ZakSmetaStates.TITLE);
        AtomicReference<Table> bufTable = new AtomicReference<>(new Table(new float[]{100f, 50f, 50f, 50f}));

        cleared.forEach(row -> {

            if (row.size() == 2 || row.size() == 1 || row.size() == 3) {

                // If the prev state was materials or works - we close buf table
                if (state.get() == ZakSmetaStates.DATA) {
                    doc.add(bufTable.get());
                }

                if (!Objects.equals(row.get(0), "Работы") && !Objects.equals(row.get(0), "Материалы")) {
                    state.set(ZakSmetaStates.TITLE);
                    doc.add(new Paragraph("\n"));
                    doc.add(new Paragraph(row.get(0)).setFontSize(18).setBold());
                } else {
                    state.set(ZakSmetaStates.SUBTITLE);
                    doc.add(new Paragraph(row.get(0)).setFontSize(14).setBold());
                }

                // Works and materials
            } else if (row.size() == 4) {

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
                // TODO: возможно добавление лишних клеток
                bufTable.get().addCell(new Paragraph(row.get(0)).setFontColor(GRAY_COLOR));
                bufTable.get().addCell(new Paragraph(row.get(1)).setFontColor(GRAY_COLOR));
                try {
                    double price1 = Double.parseDouble(row.get(2));
                    double price2 = Double.parseDouble(row.get(3));

                    bufTable.get().addCell(new Paragraph(df.format(price1)).setFontColor(GRAY_COLOR));
                    bufTable.get().addCell(new Paragraph(df.format(price2)).setFontColor(GRAY_COLOR));
                } catch (NumberFormatException ee) {
                    return;
                }

                state.set(ZakSmetaStates.DATA);
            }
        });

        doc.add(bufTable.get());
        doc.add(new AreaBreak());

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
        tableMetaFooter.addCell("Всего, командировочные расходы, рублей").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "F2352")));
        tableMetaFooter.addCell("Всего работ и материалов, рублей").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "F2353")));
        doc.add(tableMetaFooter);

        doc.add(new Paragraph("\n"));
        doc.add(new Paragraph("С перечнем работ и материалов, ознакомлен, с итоговой стоимостью согласен.").setBold());
        doc.add(new Paragraph("\n"));
        doc.add(new Paragraph("Подпись Заказчика ______________"));
        doc.add(new Paragraph("\n"));
        doc.add(new Paragraph("Подпись Подрядчика ______________"));
        doc.add(new Paragraph("\n"));

        doc.add(logoImg.setMarginLeft(30));

        doc.close();
        writer.close();
    }

    public void smetaZakRassh(String path, FormulaEvaluator evaluator, Sheet sheet, ArrayList<ArrayList<String>> rows) throws IOException {

        // |----------------------------------------------
        // | 1 - ничего
        // | 2 - Тайтлы, сабтайтлы материалы и работы
        // | 3 - Тайтлы, сабтайтлы материалы и работы
        // | 4 - В основном - итого работы
        // | 5 - Итого материалы
        // | 6 - Непонятно
        // | 7, 8 - Данные

        DecimalFormat df = new DecimalFormat("0.00");

        PdfWriter writer = new PdfWriter(Paths.get(path, "smeta_zak_rassh.pdf").toString());

        PdfDocument pdfDoc = new PdfDocument(writer);
        pdfDoc.setDefaultPageSize(PageSize.A4.rotate());

        Document doc = new Document(pdfDoc);

        PdfFont font = PdfFontFactory.createFont(FONT, PdfEncodings.IDENTITY_H);
        doc.setFont(font);

        ImageData logoData = ImageDataFactory.create(Paths.get(System.getProperty("user.dir"), "/src/img/logo.png").toString());
        Image logoImg = new Image(logoData);
        logoImg.setWidth(75);

        this.addHeaderPage(doc, "Сметный расчет для закупок", logoImg);

        //////////////////////////
        // Adding header information
        Table tableMeta = new Table(new float[]{150f, 100f});
        tableMeta.addCell("Длина дома, м.п.").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMeta.addCell(df.format(this.getCellValue(evaluator, sheet, "J2")));
        tableMeta.addCell("Ширина дома, м.п.").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMeta.addCell(df.format(this.getCellValue(evaluator, sheet, "J3")));
        tableMeta.addCell("Этажность дома").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMeta.addCell(df.format(this.getCellValue(evaluator, sheet, "J4")));
        tableMeta.addCell("S строения общая, м2").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMeta.addCell(df.format(this.getCellValue(evaluator, sheet, "J5")));
        tableMeta.addCell("S строения чистая, м2").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMeta.addCell(df.format(this.getCellValue(evaluator, sheet, "J6")));
        doc.add(tableMeta);

        doc.add(new Paragraph("\n"));

        ArrayList<ArrayList<String>> cleared = new ArrayList<>();
        rows.forEach(row -> {
            if (row.size() != 2
                    && row.size() != 3
                    && row.size() != 7
                    && row.size() != 8) {
                return;
            }
            if (row.size() == 7 || row.size() == 8) {
                // Clear all zeroed rows
                try {
                    double price = Double.parseDouble(row.get(5));
                    if (price < 1) {
                        return;
                    }
                } catch (NumberFormatException e) {
                    return;
                }
            }
            cleared.add(row);
        });


        // А теперь сука удаляем ненужные хедеры блять
        ArrayList<ArrayList<String>> toDelete = new ArrayList<>();

        for (int i = 0; i < cleared.size(); i++) {
            ArrayList<String> row = cleared.get(i);

            // Если это большой тайтл, то проверяем следующие 3
            if (row.size() == 2 || row.size() == 3) {

                if (!Objects.equals(row.get(0), "Работы") && !Objects.equals(row.get(0), "Материалы")) {

                    int headersSize = 0;
                    ArrayList<String> currentRow;

                    boolean safeToDelete = true;

                    // Пытаемся найти блять другой тайтл в пределах 4 элементов
                    for (int j = 1; j < 4; j++) {
                        try {
                            currentRow = cleared.get(i + j);
                            if ((currentRow.size() == 2 || currentRow.size() == 3)
                                    && !Objects.equals(currentRow.get(0), "Работы")
                                    && !Objects.equals(currentRow.get(0), "Материалы")) {
                                headersSize = j;
                                break;
                            } else if (currentRow.size() == 7 || currentRow.size() == 8) {
                                safeToDelete = false;
                            }
                        } catch (IndexOutOfBoundsException ignored) {

                        }
                    }

                    if (headersSize > 0 && safeToDelete) {
                        for (int j = 0; j < headersSize; j++) {
                            ArrayList<String> rowToDelete = cleared.get(i + j);
                            toDelete.add(rowToDelete);
                        }
                    }
                }
            }
        }
        cleared.removeAll(toDelete);

        // Если последний элемент это тайтл - то удаляем нахуй
        if (cleared.get(cleared.size() - 1).size() == 2 || cleared.get(cleared.size() - 1).size() == 3) {
            cleared.remove(cleared.size() - 1);
        }

        AtomicReference<ZakSmetaStates> state = new AtomicReference<>(ZakSmetaStates.TITLE);
        AtomicReference<Table> bufTable = new AtomicReference<>(new Table(new float[]{100f, 50f, 50f, 50f, 50f}));
        addTableHeaders5(bufTable.get());

        cleared.forEach(row -> {

            if (row.size() == 2 || row.size() == 3) {

                // If the prev state was materials or works - we close buf table
                if (state.get() == ZakSmetaStates.DATA) {
                    doc.add(bufTable.get());
                }

                if (!Objects.equals(row.get(0), "Работы") && !Objects.equals(row.get(0), "Материалы")) {
                    state.set(ZakSmetaStates.TITLE);
                    doc.add(new Paragraph("\n"));
                    doc.add(new Paragraph(row.get(0)).setFontSize(18).setBold());
                } else {
                    state.set(ZakSmetaStates.SUBTITLE);
                    doc.add(new Paragraph(row.get(0)).setFontSize(14).setBold());
                }

                // Works and materials
            } else if (row.size() == 7 || row.size() == 8) {

                // Checking for validity
                if (row.get(0) == null) {
                    return;
                }

                float[] cols = {100f, 50f, 50f, 50f, 50f};

                if (state.get() == ZakSmetaStates.TITLE || state.get() == ZakSmetaStates.SUBTITLE) {
                    Table t = new Table(cols);
                    addTableHeaders5(t);
                    bufTable.set(t);
                }
                bufTable.get().addCell(new Paragraph(row.get(0)).setFontColor(GRAY_COLOR));
                bufTable.get().addCell(new Paragraph(row.get(1)).setFontColor(GRAY_COLOR));
                bufTable.get().addCell(new Paragraph(row.get(2)).setFontColor(GRAY_COLOR));
                bufTable.get().addCell(new Paragraph(row.get(3)).setFontColor(GRAY_COLOR));
                bufTable.get().addCell(new Paragraph(row.get(5)).setFontColor(GRAY_COLOR));

                state.set(ZakSmetaStates.DATA);
            }
        });

        doc.add(bufTable.get());
        doc.add(new AreaBreak());

        //////////////////////////
        // Adding footer information
        Table tableMetaFooter = new Table(new float[]{150f, 100f});
        tableMetaFooter.addCell("Всего материалов, рублей").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "J2490")));
        tableMetaFooter.addCell("Всего работы, рублей").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "J2491")));
        tableMetaFooter.addCell("Всего транспортные расходы, рублей").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "J2492")));
        tableMetaFooter.addCell("Всего дополнительные расходные материалы").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "J2493")));
        tableMetaFooter.addCell("Всего, командировочные расходы, рублей").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "J2494")));
        tableMetaFooter.addCell("Всего работ и материалов, рублей").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "J2495")));
        doc.add(tableMetaFooter);

        doc.add(new Paragraph("\n"));
        doc.add(new Paragraph("С перечнем работ и материалов, ознакомлен, с итоговой стоимостью согласен.").setBold());
        doc.add(new Paragraph("\n"));
        doc.add(new Paragraph("Подпись Заказчика ______________"));
        doc.add(new Paragraph("\n"));
        doc.add(new Paragraph("Подпись Подрядчика ______________"));
        doc.add(new Paragraph("\n"));

        doc.add(logoImg.setMarginLeft(30));

        doc.close();
        writer.close();
    }

    public void smetaInternal(String path, FormulaEvaluator evaluator, Sheet sheet, ArrayList<ArrayList<String>> rows) throws IOException {

        // |--------------------------------------------------
        // | 3 - Названия групп, в том числе материалы, работы
        // | 4 - Итого работы
        // | 5 - итого материалы
        // | 6
        // | 7 - Работы
        // | 8 - Материалы

        PdfWriter writer = new PdfWriter(Paths.get(path, "smeta_internal.pdf").toString());

        PdfDocument pdfDoc = new PdfDocument(writer);
        pdfDoc.setDefaultPageSize(PageSize.A4.rotate());

        Document doc = new Document(pdfDoc);

        PdfFont font = PdfFontFactory.createFont(FONT, PdfEncodings.IDENTITY_H);
        doc.setFont(font);

        ImageData logoData = ImageDataFactory.create(Paths.get(System.getProperty("user.dir"), "/src/img/logo.png").toString());
        Image logoImg = new Image(logoData);
        logoImg.setWidth(75);

        this.addHeaderPage(doc, "Сметный расчет (внутренний)", logoImg);

        DecimalFormat df = new DecimalFormat();

        //////////////////////////
        // Добавляем хэдеры блять
        Table tableMeta = new Table(new float[]{150f, 100f});
        tableMeta.addCell("Длина дома, м.п.").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMeta.addCell(df.format(this.getCellValue(evaluator, sheet, "H2")));
        tableMeta.addCell("Ширина дома, м.п.").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMeta.addCell(df.format(this.getCellValue(evaluator, sheet, "H3")));
        tableMeta.addCell("Этажность дома").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMeta.addCell(df.format(this.getCellValue(evaluator, sheet, "H4")));
        tableMeta.addCell("S строения общая, м2").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMeta.addCell(df.format(this.getCellValue(evaluator, sheet, "H5")));
        tableMeta.addCell("S строения чистая, м2").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMeta.addCell(df.format(this.getCellValue(evaluator, sheet, "H6")));
        doc.add(tableMeta);

        doc.add(new Paragraph("\n"));

        // Удаляем всю ебаную хуйню нулевую блять
        ArrayList<ArrayList<String>> cleared = new ArrayList<>();
        rows.forEach(row -> {
            if (CLEAR_ZEROS) {
                if (row.size() != 3 && row.size() != 7 && row.size() != 8) {
                    return;
                }

                if (row.size() == 7 || row.size() == 8) {
                    try {
                        double price = Double.parseDouble(row.get(6));

                        // Бля пацаны извините так получилось
                        if (price < 10 || price == 2000) {
                            return;
                        }

                    } catch (NumberFormatException e) {
                        return;
                    }
                }
            }
            cleared.add(row);
        });

        // А теперь сука удаляем ненужные хедеры блять
        ArrayList<ArrayList<String>> toDelete = new ArrayList<>();

        for (int i = 0; i < cleared.size(); i++) {
            ArrayList<String> row = cleared.get(i);

            // Если это большой тайтл, то проверяем следующие 4
            if (row.size() == 3) {

                if (!Objects.equals(row.get(0), "Работы") && !Objects.equals(row.get(0), "Материалы") && !Objects.equals(row.get(0), "Техника и дополнительные расходы")) {

                    int headersSize = 0;
                    ArrayList<String> currentRow;

                    boolean safeToDelete = true;

                    // Пытаемся найти блять другой тайтл в пределах 3 элементов
                    for (int j = 1; j < 5; j++) {
                        try {
                            currentRow = cleared.get(i + j);
                            if (currentRow.size() == 3
                                    && !Objects.equals(currentRow.get(0), "Работы")
                                    && !Objects.equals(currentRow.get(0), "Материалы")
                                    && !Objects.equals(currentRow.get(0), "Техника и дополнительные расходы")
                                    && !Objects.equals(currentRow.get(0), "Командировочные расходы")) {
                                headersSize = j;
                            } else if (currentRow.size() == 7 || currentRow.size() == 8) {
                                safeToDelete = false;
                            }
                        } catch (IndexOutOfBoundsException ignored) {

                        }
                    }

                    if (headersSize > 0 && safeToDelete) {
                        for (int j = 0; j < headersSize; j++) {
                            ArrayList<String> rowToDelete = cleared.get(i + j);
                            toDelete.add(rowToDelete);
                        }
                    }
                }
            }
        }

        cleared.removeAll(toDelete);
        if (cleared.get(cleared.size() - 1).size() == 2 || cleared.get(cleared.size() - 1).size() == 3) {
            cleared.remove(cleared.size() - 1);
        }

        AtomicReference<InternalSmetaStates> state = new AtomicReference<>(InternalSmetaStates.TITLE);
        AtomicReference<Table> bufTable = new AtomicReference<>(new Table(new float[]{50f, 20f, 20f, 20f, 20f, 20f, 20f}));

        cleared.forEach(row -> {

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

                row.forEach(val -> bufTable.get().addCell(new Paragraph(val).setFontColor(GRAY_COLOR)));

                state.set(InternalSmetaStates.WORKS);

                // Materials
            } else if (row.size() == 8) {

                // Checking for validity
                if (row.get(0) == null) {
                    return;
                }

                // If the first col is can be parsed into double, we emerge
                try {
                    Double.parseDouble(row.get(0));
                    logger.error("could not parse table row name: " + row.get(0));
                } catch (NumberFormatException e) {
                    float[] materialsColWidth = {50f, 20f, 20f, 20f, 20f, 20f, 20f, 20f};
                    if (state.get() == InternalSmetaStates.SUBTITLE || state.get() == InternalSmetaStates.TITLE) {
                        Table t = new Table(materialsColWidth);
                        addTableHeaders8(t);
                        bufTable.set(t);
                    }

                    row.forEach(val -> bufTable.get().addCell(new Paragraph(val).setFontColor(GRAY_COLOR)));
                    state.set(InternalSmetaStates.MATERIALS);
                }
            }
        });

        // Добавляем незакрытую таблицу
        doc.add(bufTable.get());
        doc.add(new AreaBreak());

        //////////////////////////
        // Добавляем футер нахуй
        Table tableMetaFooter = new Table(new float[]{150f, 100f});
        tableMetaFooter.addCell("Всего материалов, рублей").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "H2498")));
        tableMetaFooter.addCell("Всего работы, рублей").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "H2499")));
        tableMetaFooter.addCell("Всего транспортные расходы, рублей").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "H2500")));
        tableMetaFooter.addCell("Всего дополнительные расходные материалы").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "H2501")));
//        tableMetaFooter.addCell("Всего работ и материалов, рублей").setBold().setTextAlignment(TextAlignment.RIGHT);
//        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "H2502")));
        tableMetaFooter.addCell("Всего работ и материалов, рублей").setBold().setTextAlignment(TextAlignment.RIGHT);
        tableMetaFooter.addCell(df.format(this.getCellValue(evaluator, sheet, "H2503")));
        doc.add(tableMetaFooter);

        doc.add(new Paragraph("\n"));
        doc.add(new Paragraph("С перечнем работ и материалов, ознакомлен, с итоговой стоимостью согласен.").setBold());
        doc.add(new Paragraph("\n"));
        doc.add(new Paragraph("Подпись Заказчика ______________"));
        doc.add(new Paragraph("\n"));
        doc.add(new Paragraph("Подпись Подрядчика ______________"));
        doc.add(new Paragraph("\n"));

        doc.add(logoImg.setMarginLeft(30));

        doc.close();
        writer.close();
    }

    private void addHeaderPage(Document doc, String title, Image logo) throws MalformedURLException {
        /// Вставляем лого нахуй блять
        doc.add(logo);
        doc.add(new Paragraph("\n"));
        doc.add(new Paragraph("\n"));

        doc.add(new Paragraph(title).setFontSize(18).setBold().setTextAlignment(TextAlignment.LEFT));
        doc.add(new Paragraph("Документ подготовлен сайтом https://rbc.ramdom.work").setFontSize(14).setFontColor(Color.GRAY));
        doc.add(new AreaBreak());
    }

    private Cell getCell(Sheet sheet, String id) {
        CellReference cr = new CellReference(id);
        Row row = sheet.getRow(cr.getRow());

        return row.getCell(cr.getCol());
    }

    private Double getCellValue(FormulaEvaluator evaluator, Sheet sheet, String id) {
        Cell cell = this.getCell(sheet, id);

//        evaluator.clearAllCachedResultValues();
        return evaluator.evaluate(cell).getNumberValue();
    }

    public void addTableHeaders4(Table table) {
        table.addHeaderCell(new Paragraph("Наименование работ и затрат").setBold());
        table.addHeaderCell(new Paragraph("ед. изм").setBold());
        table.addHeaderCell(new Paragraph("кол-во").setBold());
        table.addHeaderCell(new Paragraph("Цена ед. изм.").setBold());
    }


    public void addTableHeaders5(Table table) {
        table.addHeaderCell(new Paragraph("Наименование работ и затрат").setBold());
        table.addHeaderCell(new Paragraph("ед. изм").setBold());
        table.addHeaderCell(new Paragraph("кол-во").setBold());
        table.addHeaderCell(new Paragraph("Цена ед. изм.").setBold());
        table.addHeaderCell(new Paragraph("Общая цена").setBold());
    }


    public void addTableHeaders8(Table table) {
        table.addHeaderCell(new Paragraph("Наименование работ и затрат").setBold());
        table.addHeaderCell(new Paragraph("ед. изм").setBold());
        table.addHeaderCell(new Paragraph("кол-во").setBold());
        table.addHeaderCell(new Paragraph("Цена ед.изм.").setBold());
        table.addHeaderCell(new Paragraph("Цена с накруткой").setBold());
        table.addHeaderCell(new Paragraph("Общая цена").setBold());
        table.addHeaderCell(new Paragraph("Цена по прайсу").setBold());
        table.addHeaderCell(new Paragraph("Для закупа материала").setBold());
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
