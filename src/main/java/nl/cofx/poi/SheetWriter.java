package nl.cofx.poi;

import nl.cofx.poi.ticket.AssetPlanning;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.time.LocalDate;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class SheetWriter {

    public void createHeader(Sheet sheet, LocalDate startDate, LocalDate endDate) {
        var months = sheet.createRow(0);
        var weeks = sheet.createRow(1);

        var defaultStyle = sheet.getWorkbook().getCellStyleAt(1);

        Integer previousMonthCell = null;
        int cell = 1;
        String previousMonth = null;
        var currentDate = startDate;

        do {
            var monthCell = months.createCell(cell);
            monthCell.setCellStyle(defaultStyle);
            var weekCell = weeks.createCell(cell);
            weekCell.setCellStyle(defaultStyle);

            var currentMonth = DateUtil.monthName(currentDate);
            if (!Objects.equals(currentMonth, previousMonth)) {
                monthCell.setCellValue(currentMonth);

                if (previousMonthCell != null && canMerge(previousMonthCell, cell - 1)) {
                    sheet.addMergedRegion(new CellRangeAddress(0, 0, previousMonthCell, cell - 1));
                }
                previousMonthCell = cell;
            }
            previousMonth = currentMonth;

            var weekNumber = DateUtil.weekNumber(currentDate);
            weekCell.setCellValue(String.valueOf(weekNumber));

            currentDate = currentDate.plusWeeks(1);
            ++cell;
        } while (currentDate.isBefore(endDate));

        if (canMerge(previousMonthCell, cell - 1)) {
            sheet.addMergedRegion(new CellRangeAddress(0, 0, previousMonthCell, cell - 1));
        }
    }

    private boolean canMerge(Integer firstCell, int secondCell) {
        return secondCell - firstCell > 0;
    }

    public void writeRow(Sheet sheet, AssetPlanning assetPlanning, int rowNumber, LocalDate startDate, LocalDate endDate) {
        var row = sheet.createRow(rowNumber + 2);

        var defaultStyle = sheet.getWorkbook().getCellStyleAt(1);

        var assetNameCell = row.createCell(0);
        assetNameCell.setCellStyle(defaultStyle);
        assetNameCell.setCellValue(Stream.of(assetPlanning.getAssetId(), assetPlanning.getAssetName())
                .filter(Objects::nonNull)
                .filter(s -> !s.isBlank())
                .collect(Collectors.joining(" - ")));

        var plannedStyle = sheet.getWorkbook().getCellStyleAt(2);
        var ongoingStyle = sheet.getWorkbook().getCellStyleAt(3);
        var completedStyle = sheet.getWorkbook().getCellStyleAt(4);

        int cell = 1;
        var currentDate = DateUtil.firstDayOfWeek(startDate);
        while (currentDate.isBefore(endDate)) {
            var stateForWeek = assetPlanning.getStatePerWeek().get(currentDate);
            var stateCell = row.createCell(cell);
            if ("PLANNED".equals(stateForWeek)) {
                stateCell.setCellStyle(plannedStyle);
            } else if ("ONGOING".equals(stateForWeek)) {
                stateCell.setCellStyle(ongoingStyle);
            } else if ("COMPLETED".equals(stateForWeek)) {
                stateCell.setCellStyle(completedStyle);
            } else {
                stateCell.setCellStyle(defaultStyle);
            }

            currentDate = currentDate.plusWeeks(1);
            ++cell;
        }
    }

    private Color plannedColor() {
        var bytes = new byte[3];
        bytes[0] = (byte) 185;
        bytes[1] = (byte) 224;
        bytes[2] = (byte) 243;

        return new XSSFColor(bytes);
    }

    private Color ongoingColor() {
        var bytes = new byte[3];
        bytes[1] = (byte) 82;
        bytes[2] = (byte) 136;

        return new XSSFColor(bytes);
    }

    private Color completedColor() {
        var bytes = new byte[3];
        bytes[0] = (byte) 190;
        bytes[1] = (byte) 205;

        return new XSSFColor(bytes);
    }

    public void finalizeSheet(Sheet sheet, int autoSizeColumn) {
        sheet.autoSizeColumn(autoSizeColumn);
    }

    public void initialize(SXSSFSheet sheet, int autoSizeColumn) {
        sheet.trackColumnForAutoSizing(autoSizeColumn);

        var defaultStyle = sheet.getWorkbook().createCellStyle();
        setBorders(defaultStyle);

        var plannedStyle = sheet.getWorkbook().createCellStyle();
        setBorders(plannedStyle);
        setForegroundColor(plannedStyle, plannedColor());

        var ongoingStyle = sheet.getWorkbook().createCellStyle();
        setBorders(ongoingStyle);
        setForegroundColor(ongoingStyle, ongoingColor());

        var completedStyle = sheet.getWorkbook().createCellStyle();
        setBorders(completedStyle);
        setForegroundColor(completedStyle, completedColor());

        var defaultBoldStyle = sheet.getWorkbook().createCellStyle();
        setBorders(defaultBoldStyle);
        setFont(defaultBoldStyle, sheet.getWorkbook());
    }

    private void setFont(CellStyle defaultBoldStyle, SXSSFWorkbook workbook) {
        var boldFont = workbook.createFont();
        boldFont.setBold(true);
        defaultBoldStyle.setFont(boldFont);
    }

    private static void setForegroundColor(CellStyle cellStyle, Color color) {
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(color);
    }

    private static void setBorders(CellStyle cellStyle) {
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
    }

    public void createLegend(SXSSFSheet sheet) {

        var defaultStyle = sheet.getWorkbook().getCellStyleAt(1);
        var plannedStyle = sheet.getWorkbook().getCellStyleAt(2);
        var ongoingStyle = sheet.getWorkbook().getCellStyleAt(3);
        var completedStyle = sheet.getWorkbook().getCellStyleAt(4);
        var boldStyle = sheet.getWorkbook().getCellStyleAt(5);

        var titleRow = sheet.createRow(0);
        var titleCell = titleRow.createCell(0);
        titleCell.setCellStyle(boldStyle);
        titleCell.setCellValue("Legend");


        createLegendEntry(sheet, 2, plannedStyle, "Planned ticket(s)", defaultStyle);
        createLegendEntry(sheet, 3, ongoingStyle, "Ongoing ticket(s)", defaultStyle);
        createLegendEntry(sheet, 4, completedStyle, "Ended ticket(s)", defaultStyle);
    }

    private void createLegendEntry(Sheet sheet, int row, CellStyle colorCellStyle, String value, CellStyle nameCellStyle) {
        var entryRow = sheet.createRow(row);
        var colorCell = entryRow.createCell(0);
        colorCell.setCellStyle(colorCellStyle);
        var nameCell = entryRow.createCell(1);
        nameCell.setCellStyle(nameCellStyle);
        nameCell.setCellValue(value);
    }
}
