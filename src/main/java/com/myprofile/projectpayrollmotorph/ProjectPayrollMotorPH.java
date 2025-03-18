package com.myprofile.projectpayrollmotorph;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.time.*;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.*;

public class ProjectPayrollMotorPH {
    public static void main(String[] args) {
        String filePath = "src/MotorPH Employee Data.xlsx";
        try (Scanner scanner = new Scanner(System.in)) {
            System.out.print("Enter Employee Number: ");
            int empNumberInput = scanner.nextInt();
            readEmployeeDetails(filePath, empNumberInput);
        }
    }

    public static void readEmployeeDetails(String filePath, int empNumberInput) {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet1 = workbook.getSheet("Employee Details");
            Sheet sheet2 = workbook.getSheet("Attendance Record");

            if (sheet1 == null || sheet2 == null) {
                System.out.println("One or more sheets not found.");
                return;
            }

            boolean found = false;
            Row employeeRow = null;

            for (Row row : sheet1) {
                if (row.getRowNum() == 0) continue;
                Cell empNumberCell = row.getCell(0);
                if (empNumberCell == null || empNumberCell.getCellType() != CellType.NUMERIC) continue;
                int empNumber = (int) empNumberCell.getNumericCellValue();

                if (empNumber == empNumberInput) {
                    found = true;
                    employeeRow = row;
                    break;
                }
            }

            if (!found) {
                System.out.println("Employee not found.");
                return;
            }

            displayEmployeeDetails(employeeRow);

            TreeMap<LocalDate, Double> weeklyHours = new TreeMap<>();
            DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("MM/dd/yyyy");
            DateTimeFormatter[] possibleFormats = {
                DateTimeFormatter.ofPattern("HH:mm"),
                DateTimeFormatter.ofPattern("H:mm"),
                DateTimeFormatter.ofPattern("hh:mm a")
            };

            System.out.println("\n--- Daily Time In and Time Out ---");
            for (Row row : sheet2) {
                if (row.getRowNum() == 0) continue;
                Cell empNumberCell = row.getCell(0);
                if (empNumberCell == null || empNumberCell.getCellType() != CellType.NUMERIC) continue;
                int empNumber = (int) empNumberCell.getNumericCellValue();
                if (empNumber != empNumberInput) continue;

                DataFormatter formatter = new DataFormatter();
                String dateStr = formatter.formatCellValue(row.getCell(3));
                String timeInStr = formatter.formatCellValue(row.getCell(4));
                String timeOutStr = formatter.formatCellValue(row.getCell(5));

                LocalDate date;
                try {
                    date = LocalDate.parse(dateStr, dateFormatter);
                } catch (DateTimeParseException e) {
                    continue;
                }

                LocalTime timeIn = null, timeOut = null;
                for (DateTimeFormatter format : possibleFormats) {
                    try {
                        timeIn = LocalTime.parse(timeInStr, format);
                        timeOut = LocalTime.parse(timeOutStr, format);
                        break;
                    } catch (DateTimeParseException ignored) {}
                }

                if (timeIn == null || timeOut == null) continue;
                double hoursWorked = (double) Duration.between(timeIn, timeOut).toMinutes() / 60;
                LocalDate weekStart = date.with(DayOfWeek.MONDAY);
                weeklyHours.put(weekStart, weeklyHours.getOrDefault(weekStart, 0.0) + hoursWorked);
                
                System.out.println(date + " | Time In: " + timeIn + " | Time Out: " + timeOut + " | Hours Worked: " + String.format("%.2f", hoursWorked));
            }

            System.out.println("\n--- Total Weekly Working Hours ---");
            for (Map.Entry<LocalDate, Double> entry : weeklyHours.entrySet()) {
                LocalDate weekStart = entry.getKey();
                LocalDate weekEnd = weekStart.plusDays(4);
                System.out.println(weekStart + " - " + weekEnd + ": " + String.format("%.2f", entry.getValue()) + " hours");
            }
            System.out.println("------------------------------------------");

        } catch (IOException e) {
            System.out.println("Error reading file: " + e.getMessage());
        }
    }

    private static void displayEmployeeDetails(Row row) {
    System.out.println("\n--- Employee Details ---");
    System.out.println("Employee #: " + (int) row.getCell(0).getNumericCellValue());
    System.out.println("Name: " + row.getCell(2).getStringCellValue() + " " + row.getCell(1).getStringCellValue());
    System.out.println("Position: " + row.getCell(11).toString());
    System.out.println("Birthday: " + row.getCell(3).toString());
    System.out.println("Address: " + row.getCell(4).toString());
    System.out.println("Phone Number: " + row.getCell(5).toString());

    // Read Monthly Salary
    Cell salaryCell = row.getCell(13);
    double hourlyRate = 0.0;
    if (salaryCell != null && salaryCell.getCellType() == CellType.NUMERIC) {
        double monthlySalary = salaryCell.getNumericCellValue();
        hourlyRate = monthlySalary / 168;
    }

    System.out.println("Monthly Salary: " + (salaryCell != null ? salaryCell.toString() : "N/A"));
    System.out.println("Computed Hourly Rate: " + String.format("%.2f", hourlyRate));

    printWholeNumber(row.getCell(6), "SSS Number");
    printWholeNumber(row.getCell(7), "PhilHealth Number");
    printWholeNumber(row.getCell(8), "TIN Number");
    printWholeNumber(row.getCell(9), "PagIbig Number");

    System.out.println("------------------------------------------");
    printAllowances(row);
}

    private static void printWholeNumber(Cell cell, String label) {
        if (cell != null) {
            System.out.println(label + ": " + (cell.getCellType() == CellType.NUMERIC ? (long) cell.getNumericCellValue() : cell.toString()));
        }
    }

    private static void printAllowances(Row row) {
        System.out.println("Allowances:");
        System.out.println("------------------------------------------");
        System.out.println("Rice Subsidy: " + row.getCell(14).toString());
        System.out.println("Phone Allowance: " + row.getCell(15).toString());
        System.out.println("Clothing Allowance: " + row.getCell(16).toString());
        System.out.println("------------------------------------------");
    }
}