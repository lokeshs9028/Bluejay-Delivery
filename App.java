package com.example;

// Import DateUtil
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class App {

    public static void main(String[] args) {
        String filePath = "C:/Users/lokes/BlueJay-Delivery/demo/src/main/java/com/example/Assignment_Timecard.xlsx"; // Replace
                                                                                                                     // with
                                                                                                                     // the
                                                                                                                     // actual
                                                                                                                     // file
                                                                                                                     // path
        SimpleDateFormat dateFormat = new SimpleDateFormat("MM-dd-yyyy hh:mm a");

        try (FileInputStream fis = new FileInputStream(filePath);
                Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming the data is in the first sheet

            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                String positionID = getStringCellValue(row.getCell(0));
                String positionStatus = getStringCellValue(row.getCell(1));
                String timeIn = getStringCellValue(row.getCell(2));
                String timeOut = getStringCellValue(row.getCell(3));
                String timecardHours = getStringCellValue(row.getCell(4));
                String employeeName = getStringCellValue(row.getCell(7));

                if (!timeIn.isEmpty() && !timeOut.isEmpty()) {
                    Date startTime = getDateValue(row.getCell(2)); // Use getDateValue
                    Date endTime = getDateValue(row.getCell(3)); // Use getDateValue

                    long timeDifference = endTime.getTime() - startTime.getTime();
                    long hoursBetweenShifts = timeDifference / (60 * 60 * 1000);

                    if (timecardHours.contains(":")) {
                        String[] timeParts = timecardHours.split(":");
                        if (timeParts.length == 3) {
                            int totalHours = Integer.parseInt(timeParts[0]);
                            int totalMinutes = Integer.parseInt(timeParts[1]);
                            int totalSeconds = Integer.parseInt(timeParts[2]);
                            int totalWorkedHours = totalHours + (totalMinutes / 60) + (totalSeconds / 3600);

                            if (totalWorkedHours >= 7) {
                                System.out.println("Employee Name: " + employeeName);
                                System.out.println("Position ID: " + positionID);
                            }
                        }
                    }

                    if (hoursBetweenShifts < 10 && hoursBetweenShifts > 1) {
                        System.out.println("Employee Name: " + employeeName);
                        System.out.println("Position ID: " + positionID);
                    }

                    if (timeDifference > (14 * 60 * 60 * 1000)) {
                        System.out.println("Employee Name: " + employeeName);
                        System.out.println("Position ID: " + positionID);
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Helper method to safely retrieve cell value as a string
    private static String getStringCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue().trim();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            // If the cell contains a numeric value, convert it to a string
            return String.valueOf(cell.getNumericCellValue()).trim();
        } else {
            // Handle other cell types if needed
            return "";
        }
    }

    // Helper method to convert numeric date values to a java.util.Date object
    private static Date getDateValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            return cell.getDateCellValue();
        } else {
            // Handle other cell types or invalid dates if needed
            return null;
        }
    }
}
