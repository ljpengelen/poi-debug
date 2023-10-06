package nl.cofx.poi;

import nl.cofx.poi.ticket.AssetPlanning;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.time.LocalDate;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.Mockito.spy;
import static org.mockito.Mockito.verify;

class SheetWriterTest {

    private static final LocalDate START_DATE = LocalDate.of(2023, 1, 1);
    private static final LocalDate END_DATE = LocalDate.of(2023, 3, 1);
    private static final String ASSET_NAME = "assetName";
    private static final String ASSET_ID = "assetId";

    private final SheetWriter sheetWriter = new SheetWriter();

    @Test
    void createsHeader() throws IOException {
        try (var workbook = new SXSSFWorkbook()) {
            var sheet = workbook.createSheet();
            sheetWriter.createHeader(sheet, START_DATE, END_DATE);

            var monthRow = sheet.getRow(0);
            assertThat(monthRow.getCell(1).getStringCellValue()).isEqualTo("December");
            assertThat(monthRow.getCell(2).getStringCellValue()).isEqualTo("January");
            assertThat(monthRow.getCell(6).getStringCellValue()).isEqualTo("February");

            var weekRow = sheet.getRow(1);
            assertThat(weekRow.getCell(1).getStringCellValue()).isEqualTo("52");
            assertThat(weekRow.getCell(2).getStringCellValue()).isEqualTo("1");
            assertThat(weekRow.getCell(3).getStringCellValue()).isEqualTo("2");
            assertThat(weekRow.getCell(4).getStringCellValue()).isEqualTo("3");
            assertThat(weekRow.getCell(5).getStringCellValue()).isEqualTo("4");
            assertThat(weekRow.getCell(6).getStringCellValue()).isEqualTo("5");
            assertThat(weekRow.getCell(7).getStringCellValue()).isEqualTo("6");
            assertThat(weekRow.getCell(8).getStringCellValue()).isEqualTo("7");
            assertThat(weekRow.getCell(9).getStringCellValue()).isEqualTo("8");

            assertThat(sheet.getMergedRegions()).containsExactlyInAnyOrder(
                    new CellRangeAddress(0, 0, 2, 5),
                    new CellRangeAddress(0, 0, 6, 9));
        }
    }

    @Test
    void initializesSheet() throws IOException {
        try (var workbook = new SXSSFWorkbook()) {
            var sheet = spy(workbook.createSheet());

            sheetWriter.initialize(sheet, 0);

            verify(sheet).trackColumnForAutoSizing(0);

            var defaultStyle = workbook.getCellStyleAt(1);
            assertThat(defaultStyle.getBorderTop()).isEqualTo(BorderStyle.THIN);
            assertThat(defaultStyle.getBorderRight()).isEqualTo(BorderStyle.THIN);
            assertThat(defaultStyle.getBorderBottom()).isEqualTo(BorderStyle.THIN);
            assertThat(defaultStyle.getBorderLeft()).isEqualTo(BorderStyle.THIN);

            var plannedStyle = workbook.getCellStyleAt(2);
            assertThat(((XSSFColor) plannedStyle.getFillForegroundColorColor()).getRGB()).contains(185, 224, 243);
            var ongoingStyle = workbook.getCellStyleAt(3);
            assertThat(((XSSFColor) ongoingStyle.getFillForegroundColorColor()).getRGB()).contains(0, 82, 136);
            var completedStyle = workbook.getCellStyleAt(4);
            assertThat(((XSSFColor) completedStyle.getFillForegroundColorColor()).getRGB()).contains(190, 205, 0);
        }
    }

    @Test
    void writesRow() throws IOException {
        try (var workbook = new SXSSFWorkbook()) {
            var sheet = workbook.createSheet();
            sheetWriter.initialize(sheet, 0);

            sheetWriter.writeRow(sheet, AssetPlanning.builder()
                    .assetId(ASSET_ID)
                    .assetName(ASSET_NAME)
                    .statePerWeek(Map.of(LocalDate.of(2022, 12, 26), "PLANNED"))
                    .build(), 0, START_DATE, END_DATE);

            var row = sheet.getRow(2);
            assertThat(row.getCell(0).getStringCellValue()).isEqualTo(ASSET_ID + " - " + ASSET_NAME);
            var fillForegroundColorColor = row.getCell(1).getCellStyle().getFillForegroundColorColor();
            assertThat(fillForegroundColorColor).isInstanceOf(XSSFColor.class);
            assertThat(((XSSFColor) fillForegroundColorColor).getRGB()).contains(185, 224, 243);
        }
    }

    @Test
    void writesRowContainingAssetId_givenAssetWithoutName() throws IOException {
        try (var workbook = new SXSSFWorkbook()) {
            var sheet = workbook.createSheet();
            sheetWriter.initialize(sheet, 0);

            sheetWriter.writeRow(sheet, AssetPlanning.builder()
                    .assetId(ASSET_ID)
                    .statePerWeek(Map.of(LocalDate.of(2022, 12, 26), "PLANNED"))
                    .build(), 0, START_DATE, END_DATE);

            var row = sheet.getRow(2);
            assertThat(row.getCell(0).getStringCellValue()).isEqualTo(ASSET_ID);
            var fillForegroundColorColor = row.getCell(1).getCellStyle().getFillForegroundColorColor();
            assertThat(fillForegroundColorColor).isInstanceOf(XSSFColor.class);
            assertThat(((XSSFColor) fillForegroundColorColor).getRGB()).contains(185, 224, 243);
        }
    }

    @Test
    void createsLegend() throws IOException {
        try (var workbook = new SXSSFWorkbook()) {
            var sheet = workbook.createSheet();
            sheetWriter.initialize(sheet, 1);
            sheetWriter.createLegend(sheet);

            var titleRow = sheet.getRow(0);
            assertThat(titleRow.getCell(0).getStringCellValue()).isEqualTo("Legend");

            var plannedTicketsRow = sheet.getRow(2);
            assertThat(plannedTicketsRow.getCell(1).getStringCellValue()).isEqualTo("Planned ticket(s)");
            var ptFillForegroundColorColor = plannedTicketsRow.getCell(0).getCellStyle().getFillForegroundColorColor();
            assertThat(ptFillForegroundColorColor).isInstanceOf(XSSFColor.class);
            assertThat(((XSSFColor) ptFillForegroundColorColor).getRGB()).contains(185, 224, 243);

            var ongoingTicketsRow = sheet.getRow(3);
            assertThat(ongoingTicketsRow.getCell(1).getStringCellValue()).isEqualTo("Ongoing ticket(s)");
            var otFillForegroundColorColor = ongoingTicketsRow.getCell(0).getCellStyle().getFillForegroundColorColor();
            assertThat(otFillForegroundColorColor).isInstanceOf(XSSFColor.class);
            assertThat(((XSSFColor) otFillForegroundColorColor).getRGB()).contains(0, 82, 136);

            var endedTicketsRow = sheet.getRow(4);
            assertThat(endedTicketsRow.getCell(1).getStringCellValue()).isEqualTo("Ended ticket(s)");
            var etFillForegroundColorColor = endedTicketsRow.getCell(0).getCellStyle().getFillForegroundColorColor();
            assertThat(etFillForegroundColorColor).isInstanceOf(XSSFColor.class);
            assertThat(((XSSFColor) etFillForegroundColorColor).getRGB()).contains(190, 205, 0);
        }
    }

    @Test
    void finalizesSheet() throws IOException {
        try (var workbook = new SXSSFWorkbook()) {
            var sheet = spy(workbook.createSheet());
            sheetWriter.initialize(sheet, 0);

            sheetWriter.finalizeSheet(sheet, 0);

            verify(sheet).autoSizeColumn(0);
        }
    }
}
