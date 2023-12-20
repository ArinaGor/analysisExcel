package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Locale;
import java.util.Scanner;

public class Main {
    public static void main(String[] args) throws IOException {
        System.out.println("Пожалуйста, введите путь до таблицы успеваемости студентов без кавычек");
        Scanner scan = new Scanner(System.in);
        String way = scan.nextLine();
        if (!(way.endsWith(".xlsx") | way.endsWith(".xls"))) {
            do {
                System.out.println("Введен неверный путь. Проверьте, что указан путь до файла Excel и попробуйте ещё раз");
                way = scan.nextLine();
            }
            while (!(way.endsWith(".xlsx") | way.endsWith(".xls")));
        }

        File input = new File(way);
        Workbook workbook = WorkbookFactory.create(input);
        Sheet sheet = workbook.getSheetAt(0);

        int startRow = 0;
        Row rowCheck = sheet.getRow(0);
        Cell cellCheck = rowCheck.getCell(1);
        try {
            String cellValue = cellCheck.getStringCellValue().toLowerCase(Locale.ROOT);
            if (cellValue.contains("оценка")){
                startRow = 1;
            }
        }
        catch (IllegalStateException e){
            startRow = 0;
        }

        int excellentCount = 0;
        int goodCount = 0;
        int averageCount = 0;
        int studentCount = 0;
        int lastRow = sheet.getLastRowNum();
        double averageScore;
        double totalScore = 0.0;

        for (int i = startRow; i <= lastRow; i++) {

            Row row = sheet.getRow(i);
            Cell scoreCell = row.getCell(1);
            int score;
            try {
                score = (int) scoreCell.getNumericCellValue();
                // обработка статистики
                try {
                    if (score == 5) {
                        excellentCount++;
                        totalScore += score;
                        studentCount++;
                    } else if (score == 4) {
                        goodCount++;
                        totalScore += score;
                        studentCount++;
                    } else if (score == 3) {
                        averageCount++;
                        totalScore += score;
                        studentCount++;
                    } else if (score > 5) {
                        System.out.println("В Вашей таблице успеваемости присутствует оценка выше пяти. Пожалуйста, убедитесь в правильности введенных данных");
                    } else if (score < 3) {
                        System.out.println("В Вашей таблице успеваемости присутствует оценка ниже трех. Пожалуйста, убедитесь в правильности введенных данных");
                    }
                } catch (NumberFormatException ignored) {
                }
            } catch (IllegalStateException ignored) {
            }
        }

        averageScore = totalScore / studentCount;

        // Создание нового Excel файла для результатов
        Workbook resultWorkbook = new XSSFWorkbook();
        Sheet resultSheet = resultWorkbook.createSheet("Результаты группы");

        // Форматирование заголовков столбцов
        Font headerExampleFont = resultWorkbook.createFont();
        headerExampleFont.setBold(true);
        CellStyle headerExampleCellStyle = resultWorkbook.createCellStyle();
        headerExampleCellStyle.setFont(headerExampleFont);

        // Отличники
        Row headerExcellentRow = resultSheet.createRow(0);
        Cell headerExcellentCell = headerExcellentRow.createCell(0);
        headerExcellentCell.setCellValue("Количество отличников");
        headerExcellentCell.setCellStyle(headerExampleCellStyle);

        Row resultExcellentRow = resultSheet.createRow(1);
        Cell resultExcellentCell = resultExcellentRow.createCell(0);
        resultExcellentCell.setCellValue(excellentCount);

        // Хорошисты
        Row headerGoodRow = resultSheet.getRow(0);
        Cell headerGoodCell = headerGoodRow.createCell(1);
        headerGoodCell.setCellValue("Количество хорошистов");
        headerGoodCell.setCellStyle(headerExampleCellStyle);

        Row resultGoodRow = resultSheet.getRow(1);
        Cell resultGoodCell = resultGoodRow.createCell(1);
        resultGoodCell.setCellValue(goodCount);

        // Троечники
        Row headerAverageRow = resultSheet.getRow(0);
        Cell headerAverageCell = headerAverageRow.createCell(2);
        headerAverageCell.setCellValue("Количество троечников");
        headerAverageCell.setCellStyle(headerExampleCellStyle);

        Row resultAverageRow = resultSheet.getRow(1);
        Cell resultAverageCell = resultAverageRow.createCell(2);
        resultAverageCell.setCellValue(averageCount);

        // Средний балл
        Row headerAverageScoreRow = resultSheet.getRow(0);
        Cell headerAverageScoreCell = headerAverageScoreRow.createCell(3);
        headerAverageScoreCell.setCellValue("Средний балл учеников");
        headerAverageScoreCell.setCellStyle(headerExampleCellStyle);

        Row resultAverageScoreRow = resultSheet.getRow(1);
        Cell resultAverageScoreCell = resultAverageScoreRow.createCell(3);
        resultAverageScoreCell.setCellValue(averageScore);

        // Запись статистики в новый Excel файл
        Row headerExcellentFioRow = resultSheet.getRow(0);
        Cell headerExcellentFioCell = headerExcellentFioRow.createCell(4);
        headerExcellentFioCell.setCellValue("Данные получивших оценку отлично");
        headerExcellentFioCell.setCellStyle(headerExampleCellStyle);

        Row headerGoodFioRow = resultSheet.getRow(0);
        Cell headerGoodFioCell = headerGoodFioRow.createCell(5);
        headerGoodFioCell.setCellValue("Данные получивших оценку 4");
        headerGoodFioCell.setCellStyle(headerExampleCellStyle);

        Row headerAverageFioRow = resultSheet.getRow(0);
        Cell headerAverageFioCell = headerAverageFioRow.createCell(6);
        headerAverageFioCell.setCellValue("Данные получивших оценку 3");
        headerAverageFioCell.setCellStyle(headerExampleCellStyle);

        Row headerFailedFioRow = resultSheet.getRow(0);
        Cell headerFailedFioCell = headerFailedFioRow.createCell(7);
        headerFailedFioCell.setCellValue("Данные недопущенных");
        headerFailedFioCell.setCellStyle(headerExampleCellStyle);

        // Создание графика
        XSSFDrawing drawing = (XSSFDrawing) resultSheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 10, 8, 26);
        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText("График успеваемости учащихся");
        chart.setTitleOverlay(false);

        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP_RIGHT);

        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle("Полученные оценки");
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle("Количество студентов");

        XDDFDataSource<String> numberOfStudents = XDDFDataSourcesFactory.fromStringCellRange((XSSFSheet) resultSheet,
                new CellRangeAddress(0, 0, 0, 2));
        XDDFNumericalDataSource<Double> Students = XDDFDataSourcesFactory.fromNumericCellRange((XSSFSheet) resultSheet,
                new CellRangeAddress(1, 1, 0, 2));
        XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);

        XDDFLineChartData.Series series1 = (XDDFLineChartData.Series) data.addSeries(numberOfStudents, Students);
        series1.setTitle("Статистика", null);
        series1.setSmooth(false);
        series1.setMarkerStyle(MarkerStyle.STAR);

        chart.plot(data);

        int scoreResult;
        int excellentRowNum = 0;
        int goodRowNum = 0;
        int averageRowNum = 0;
        int failedRowNum = 0;

        for (int i = startRow; i <= lastRow; i++) {
            Row rowResult = sheet.getRow(startRow);
            String fio = rowResult.getCell(0).getStringCellValue();
            try {
                scoreResult = (int) rowResult.getCell(1).getNumericCellValue();
                if (scoreResult == 2) {
                    scoreResult = 0;
                }
            } catch (IllegalStateException e) {
                scoreResult = 2;
            }

            if (scoreResult == 5) {
                excellentRowNum++;
                Row excellentRow = resultSheet.getRow(excellentRowNum);

                if (excellentRow == null) {
                    excellentRow = resultSheet.createRow(excellentRowNum);
                }
                Cell excellentCell = excellentRow.createCell(4);
                excellentCell.setCellValue(fio);
            }

            if (scoreResult == 4) {
                goodRowNum++;
                Row goodRow = resultSheet.getRow(goodRowNum);

                if (goodRow == null) {
                    goodRow = resultSheet.createRow(goodRowNum);
                }
                Cell goodCell = goodRow.createCell(5);
                goodCell.setCellValue(fio);
            }

            if (scoreResult == 3) {
                averageRowNum++;
                Row averageRow = resultSheet.getRow(averageRowNum);

                if (averageRow == null) {
                    averageRow = resultSheet.createRow(averageRowNum);
                }
                Cell averageCell = averageRow.createCell(6);
                averageCell.setCellValue(fio);
            }

            if (scoreResult == 2) {
                failedRowNum++;
                Row failedRow = resultSheet.getRow(failedRowNum);

                if (failedRow == null) {
                    failedRow = resultSheet.createRow(failedRowNum);
                }

                Cell failedCell = failedRow.createCell(7);
                failedCell.setCellValue(fio);
            }
            startRow++;
        }

        int columnCount = resultSheet.getRow(0).getLastCellNum();
        for (int k = 0; k < columnCount; k++){
            resultSheet.autoSizeColumn(k);
        }

        // Закрытие файлов
        resultWorkbook.write(new FileOutputStream("результаты.xlsx"));
        resultWorkbook.close();
        System.out.println("Анализ успешно выполнен. Результат сохранен в файл 'результаты.xlsx'.");
    }
}