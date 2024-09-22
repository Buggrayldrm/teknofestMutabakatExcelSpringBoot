package org.example.excelproject.service;

import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.excelproject.entities.DataEntity;
import org.example.excelproject.repository.ExcelRepository;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.util.List;

@Service
@RequiredArgsConstructor
public class ExcelService {

    private final ExcelRepository excelRepository;
    private XSSFWorkbook workbook;
    private Sheet sheet;

    private CellStyle createHeaderStyle() {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        font.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private CellStyle createDataCellStyle() {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private CellStyle createNumericCellStyle() {
        CellStyle style = createDataCellStyle();
        DataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat("#,##0"));
        return style;
    }

    private CellStyle createCurrencyCellStyle() {
        CellStyle style = createDataCellStyle();
        DataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat("₺#,##0.00"));
        return style;
    }

    private void writeHeader() {
        Row headerRow = sheet.createRow(0);
        String[] headers = {"TC", "City", "Gün", "Öğlen Taşeron", "Öğlen Taşeron Birim Fiyat", "Öğlen Personel", "Öğlen Personel Birim Fiyat", "Akşam Taşeron", "Akşam Taşeron Birim Fiyat", "Akşam Personel", "Akşam Personel Birim Fiyat", "Gece Yemek", "Gece Yemek Birim Fiyat", "Lunchbox", "Lunchbox Birim Fiyat"};
        CellStyle headerStyle = createHeaderStyle();

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }
    }

    private void writeData(List<DataEntity> data) {
        int rowCount = 1;
        CellStyle dataStyle = createDataCellStyle();
        CellStyle numericStyle = createNumericCellStyle();
        CellStyle currencyStyle = createCurrencyCellStyle();

        for (DataEntity entity : data) {
            Row row = sheet.createRow(rowCount++);
            row.createCell(0).setCellValue(entity.getTc());
            row.createCell(1).setCellValue("Adana");
            row.createCell(2).setCellValue(entity.getDate());
            row.createCell(3).setCellValue(entity.getLunchSubcontractor());
            row.createCell(4).setCellValue(entity.getLunchSubcontractorPrice());
            row.createCell(5).setCellValue(entity.getLunchStaff());
            row.createCell(6).setCellValue(entity.getLunchStaffPrice());
            row.createCell(7).setCellValue(entity.getDinnerSubcontractor());
            row.createCell(8).setCellValue(entity.getDinnerSubcontractorPrice());
            row.createCell(9).setCellValue(entity.getDinnerStaff());
            row.createCell(10).setCellValue(entity.getDinnerStaffPrice());
            row.createCell(11).setCellValue(entity.getNightFood());
            row.createCell(12).setCellValue(entity.getNightFoodPrice());
            row.createCell(13).setCellValue(entity.getLunchbox());
            row.createCell(14).setCellValue(entity.getLunchboxPrice());

            for (int i = 0; i < 15; i++) {
                Cell cell = row.getCell(i);
                if (i % 2 == 1) { // Fiyat hücreleri
                    cell.setCellStyle(currencyStyle);
                } else if (i >= 3) { // Sayısal hücreler
                    cell.setCellStyle(numericStyle);
                } else { // Metin hücreler
                    cell.setCellStyle(dataStyle);
                }
            }
        }
    }

    private void writeTotals(List<DataEntity> data) {
        int lastRow = sheet.getLastRowNum() + 1;

        // Her sipariş türü için toplam sayılar
        Row totalsRow = sheet.createRow(lastRow++);
        CellStyle headerStyle = createHeaderStyle();

        Cell tcCell = totalsRow.createCell(0);
        tcCell.setCellValue("Sipariş Toplamı:");
        tcCell.setCellStyle(headerStyle);

        int totalLunchSubcontractor = 0;
        int totalLunchStaff = 0;
        int totalDinnerSubcontractor = 0;
        int totalDinnerStaff = 0;
        int totalNightFood = 0;
        int totalLunchbox = 0;

        for (DataEntity entity : data) {
            totalLunchSubcontractor += entity.getLunchSubcontractor();
            totalLunchStaff += entity.getLunchStaff();
            totalDinnerSubcontractor += entity.getDinnerSubcontractor();
            totalDinnerStaff += entity.getDinnerStaff();
            totalNightFood += entity.getNightFood();
            totalLunchbox += entity.getLunchbox();
        }

        // Toplamları her sütunun altına yazdırırken createCell() ile hücreleri oluştur
        totalsRow.createCell(3).setCellValue(totalLunchSubcontractor);
        totalsRow.createCell(5).setCellValue(totalLunchStaff);
        totalsRow.createCell(7).setCellValue(totalDinnerSubcontractor);
        totalsRow.createCell(9).setCellValue(totalDinnerStaff);
        totalsRow.createCell(11).setCellValue(totalNightFood);
        totalsRow.createCell(13).setCellValue(totalLunchbox);

        // Hücre stillerini uygularken her hücreyi önce createCell() ile oluşturduğumuzdan emin olalım
        for (int i = 3; i <= 13; i += 2) {
            Cell cell = totalsRow.getCell(i);
            if (cell == null) {
                cell = totalsRow.createCell(i); // Eğer hücre yoksa oluştur
            }
            cell.setCellStyle(createNumericCellStyle());
        }

        // Fiyatlar için KDV'li ve KDV'siz toplamı ayrı satırda yazalım
        Row priceTotalsRow = sheet.createRow(lastRow);
        Cell totalLabelCell = priceTotalsRow.createCell(0);
        totalLabelCell.setCellValue("KDV'siz Toplam:");
        totalLabelCell.setCellStyle(headerStyle);

        double totalWithoutVAT = 0;
        for (DataEntity entity : data) {
            totalWithoutVAT += entity.getLunchSubcontractor() * entity.getLunchSubcontractorPrice();
            totalWithoutVAT += entity.getLunchStaff() * entity.getLunchStaffPrice();
            totalWithoutVAT += entity.getDinnerSubcontractor() * entity.getDinnerSubcontractorPrice();
            totalWithoutVAT += entity.getDinnerStaff() * entity.getDinnerStaffPrice();
            totalWithoutVAT += entity.getNightFood() * entity.getNightFoodPrice();
            totalWithoutVAT += entity.getLunchbox() * entity.getLunchboxPrice();
        }

        double totalWithVAT = totalWithoutVAT * 1.10; // %10 KDV

        Cell totalWithoutVATCell = priceTotalsRow.createCell(1);
        totalWithoutVATCell.setCellValue(totalWithoutVAT);
        totalWithoutVATCell.setCellStyle(createCurrencyCellStyle());

        Row vatRow = sheet.createRow(lastRow + 1);
        Cell totalWithVATLabelCell = vatRow.createCell(0);
        totalWithVATLabelCell.setCellValue("%10 KDV'li Toplam:");
        totalWithVATLabelCell.setCellStyle(headerStyle);

        Cell totalWithVATCell = vatRow.createCell(1);
        totalWithVATCell.setCellValue(totalWithVAT);
        totalWithVATCell.setCellStyle(createCurrencyCellStyle());
    }



    private void createTotalRow(Row row, String label, int total) {
        Cell labelCell = row.createCell(0);
        labelCell.setCellValue(label);
        labelCell.setCellStyle(createHeaderStyle());

        Cell totalCell = row.createCell(1);
        totalCell.setCellValue(total);
        totalCell.setCellStyle(createNumericCellStyle());
    }



    public void export(HttpServletResponse response) throws IOException {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Yemek Tablosu");

        // Sütun genişliklerini ayarla
        for (int i = 0; i < 15; i++) {
            sheet.setColumnWidth(i, 24 * 256);
        }

        // Başlık, veri ve toplamları yaz
        writeHeader();
        List<DataEntity> data = excelRepository.findAll();
        writeData(data);
        writeTotals(data);

        // HTTP yanıtını ayarla ve dosyayı gönder
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=yemek_tablosu.xlsx");
        response.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
        response.setHeader("Pragma", "no-cache");
        response.setDateHeader("Expires", 0);

        workbook.write(response.getOutputStream());
        workbook.close();
    }
}
