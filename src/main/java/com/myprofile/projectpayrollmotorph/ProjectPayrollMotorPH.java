package com.myprofile.projectpayrollmotorph;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.time.*;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.temporal.ChronoUnit;
import java.util.*;

public class ProjectPayrollMotorPH {
    public static void main(String[] args) {
        String filePath = "src/MotorPH Employee Data.xlsx";
        try (Scanner scanner = new Scanner(System.in)) {
            System.out.print("Enter Employee Number: ");
            int empNumberInput = scanner.nextInt();
            scanner.nextLine();

            System.out.print("Enter Start Date (MM/dd/yyyy): ");
            String startDateStr = scanner.nextLine();
            System.out.print("Enter End Date (MM/dd/yyyy): ");
            String endDateStr = scanner.nextLine();

            DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("MM/dd/yyyy");
            LocalDate startDate = LocalDate.parse(startDateStr, dateFormatter);
            LocalDate endDate = LocalDate.parse(endDateStr, dateFormatter);

            readEmployeeDetails(filePath, empNumberInput, startDate, endDate);
        } catch (DateTimeParseException e) {
            System.out.println("Invalid date format. Please use MM/dd/yyyy.");
        }
    }

    public static void readEmployeeDetails(String filePath, int empNumberInput, LocalDate startDate, LocalDate endDate) {
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
            TreeMap<LocalDate, Double> weeklyOvertime = new TreeMap<>();
            DateTimeFormatter dateFormatter = DateTimeFormatter.ofPattern("MM/dd/yyyy");
            DateTimeFormatter[] possibleFormats = {
                DateTimeFormatter.ofPattern("HH:mm"),
                DateTimeFormatter.ofPattern("H:mm"),
                DateTimeFormatter.ofPattern("hh:mm a")
            };

            System.out.println("\n--- Daily Time In and Time Out (Filtered) ---");
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
                    if (date.isBefore(startDate) || date.isAfter(endDate)) continue;
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

                 // Standard times
                LocalTime standardTimeIn = LocalTime.of(8, 0);   // 8:00 AM
                LocalTime gracePeriod = LocalTime.of(8, 10);    // Latest clock-in for payable OT
                LocalTime strictLimit = LocalTime.of(8, 11);    // 8:11 AM threshold
                LocalTime standardTimeOut = LocalTime.of(17, 0); // 5:00 PM
                LocalTime overtimeCutoff = LocalTime.of(17, 1); // 5:01 PM

                double hoursWorked, netHoursWorked, overtimeHours = 0.0, lateMinutes = 0.0;

                // Deduct late minutes if time in is 8:11 AM or later
            if (!timeIn.isBefore(strictLimit)) {
                lateMinutes = (double) Duration.between(standardTimeIn, timeIn).toMinutes();
                timeOut = standardTimeOut; // Force standard time out at 5:00 PM for late employees
            }

            // Cap time out at 5:00 PM if the employee was late (clocked in 8:11 AM or later)
            if (timeIn.isAfter(gracePeriod) && timeOut.isAfter(standardTimeOut)) {
                timeOut = standardTimeOut;
            }

            // Calculate total hours worked
            hoursWorked = (double) Duration.between(timeIn, timeOut).toMinutes() / 60;
            netHoursWorked = Math.max(hoursWorked - 1.0, 0); // Deduct 1-hour break

            // Overtime only applies if the employee clocked in on or before 8:10 AM
            if (!timeIn.isAfter(gracePeriod) && timeOut.isAfter(standardTimeOut)) {
                overtimeHours = (double) Duration.between(standardTimeOut, timeOut).toMinutes() / 60;
            }

            // Store weekly records
            LocalDate weekStart = date.with(DayOfWeek.MONDAY);
            weeklyHours.put(weekStart, weeklyHours.getOrDefault(weekStart, 0.0) + netHoursWorked);
            weeklyOvertime.put(weekStart, weeklyOvertime.getOrDefault(weekStart, 0.0) + overtimeHours);

            System.out.println(date + " | Time In: " + timeIn + " | Time Out: " + timeOut +
                " | Hours Worked: " + String.format("%.2f", hoursWorked) +
                " | Net Hours Worked: " + String.format("%.2f", netHoursWorked) +
                " | Payable Overtime: " + String.format("%.2f", overtimeHours) + " hours" +
                " | Late Deduction: " + String.format("%.2f", lateMinutes) + " minutes");
            }
            
            System.out.println("\n--- Total Weekly Working Hours ---");
            for (Map.Entry<LocalDate, Double> entry : weeklyHours.entrySet()) {
                LocalDate weekStart = entry.getKey();
                LocalDate weekEnd = weekStart.plusDays(4);
                System.out.println(weekStart + " - " + weekEnd + ": " + String.format("%.2f", entry.getValue()) + " hours");
            }
            System.out.println("------------------------------------------");

            // Salary calculation
        Cell salaryCell = employeeRow.getCell(13);
        double hourlyRate = (salaryCell != null && salaryCell.getCellType() == CellType.NUMERIC) ? salaryCell.getNumericCellValue() / 168 : 0.0;
        double overtimeRate = hourlyRate * 1.25;

        System.out.println("\n--- Total Weekly Salary ---");

            double totalMonthlyNetSalary = 0.0; 

            for (Map.Entry<LocalDate, Double> entry : weeklyHours.entrySet()) {
                LocalDate weekStart = entry.getKey();
                LocalDate weekEnd = weekStart.plusDays(4);
                double totalHours = entry.getValue();
                double overtimeHours = weeklyOvertime.getOrDefault(weekStart, 0.0);

                // Calculate Gross Salary
                double grossSalary = (totalHours * hourlyRate) + (overtimeHours * overtimeRate);

                // Retrieve Total Allowance from Excel (Columns 14, 15, 16)
                double totalAllowance = 0.0;
                for (int i = 14; i <= 16; i++) {
                    Cell allowanceCell = employeeRow.getCell(i);
                    if (allowanceCell != null && allowanceCell.getCellType() == CellType.NUMERIC) {
                        totalAllowance += allowanceCell.getNumericCellValue();
                    }
                }

                // Determine the number of weeks in the month
                YearMonth currentMonth = YearMonth.from(weekStart);
                int numberOfWeeks = (int) ChronoUnit.WEEKS.between(currentMonth.atDay(1), currentMonth.atEndOfMonth());

                // Divide Allowance by Number of Weeks
                double weeklyAllowance = totalAllowance / numberOfWeeks;

                // Calculate Net Weekly Salary (Without Deductions Yet)
                double netWeeklySalary = grossSalary + weeklyAllowance;
                totalMonthlyNetSalary += netWeeklySalary; // Accumulate weekly net salaries

                // Display Weekly Salary Breakdown
                System.out.println(weekStart + " - " + weekEnd + 
                    "\n  Gross Salary    : " + String.format("%.2f", grossSalary) + " PHP" +
                    "\n  Weekly Allowance: " + String.format("%.2f", weeklyAllowance) + " PHP" +
                    "\n  Net Weekly Pay  : " + String.format("%.2f", netWeeklySalary) + " PHP" +
                    "\n------------------------------------------------------");
            }

            // Calculate Monthly Deductions (Excluding Tax)
            double sss = calculateSSS(totalMonthlyNetSalary);  
            double philHealth = calculatePhilHealth(totalMonthlyNetSalary);
            double pagIbig = 100.00;  // Fixed Pag-IBIG Contribution
            double totalDeductionsBeforeTax = sss + philHealth + pagIbig;

            double finalMonthlyNetSalaryBeforeTax = totalMonthlyNetSalary - totalDeductionsBeforeTax;

            // Display Total Monthly Net Pay and Deductions
            System.out.println("\n--- End of Month Summary ---");
            System.out.println("  Total Net Weekly Pay : " + String.format("%.2f", totalMonthlyNetSalary) + " PHP");
            System.out.println("\n  ---------------- Deductions (Before Tax) ----------------");
            System.out.println("  SSS Contribution : " + String.format("%.2f", sss) + " PHP");
            System.out.println("  PhilHealth       : " + String.format("%.2f", philHealth) + " PHP");
            System.out.println("  Pag-IBIG         : " + String.format("%.2f", pagIbig) + " PHP");
            System.out.println("  --------------------------------------------");
            System.out.println("  Total Deductions Before Tax : " + String.format("%.2f", totalDeductionsBeforeTax) + " PHP");

            // Calculate and Display Tax Separately
            double tax = calculateTax(finalMonthlyNetSalaryBeforeTax);
            double finalMonthlyNetSalary = finalMonthlyNetSalaryBeforeTax - tax;

            System.out.println("\n  Income Tax       : " + String.format("%.2f", tax) + " PHP");
            System.out.println("  --------------------------------------------");
            System.out.println("  Final Monthly Net Salary: " + String.format("%.2f", finalMonthlyNetSalary) + " PHP");
            System.out.println("------------------------------------------------------");


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

    private static double calculateSSS(double totalMonthlyNetSalary) {
        if (totalMonthlyNetSalary < 3250) return 135.00;
        else if (totalMonthlyNetSalary < 3750) return 157.50;
        else if (totalMonthlyNetSalary < 4250) return 180.00;
        else if (totalMonthlyNetSalary < 4750) return 202.50;
        else if (totalMonthlyNetSalary < 5250) return 225.00;
        else if (totalMonthlyNetSalary < 5750) return 247.50;
        else if (totalMonthlyNetSalary < 6250) return 270.00;
        else if (totalMonthlyNetSalary < 6750) return 292.50;
        else if (totalMonthlyNetSalary < 7250) return 315.00;
        else if (totalMonthlyNetSalary < 7750) return 337.50;
        else if (totalMonthlyNetSalary < 8250) return 360.00;
        else if (totalMonthlyNetSalary < 8750) return 382.50;
        else if (totalMonthlyNetSalary < 9250) return 405.00;
        else if (totalMonthlyNetSalary < 9750) return 427.50;
        else if (totalMonthlyNetSalary < 10250) return 450.00;
        else if (totalMonthlyNetSalary < 10750) return 472.50;
        else if (totalMonthlyNetSalary < 11250) return 495.00;
        else if (totalMonthlyNetSalary < 11750) return 517.50;
        else if (totalMonthlyNetSalary < 12250) return 540.00;
        else if (totalMonthlyNetSalary < 12750) return 562.50;
        else if (totalMonthlyNetSalary < 13250) return 585.00;
        else if (totalMonthlyNetSalary < 13750) return 607.50;
        else if (totalMonthlyNetSalary < 14250) return 630.00;
        else if (totalMonthlyNetSalary < 14750) return 652.50;
        else if (totalMonthlyNetSalary < 15250) return 675.00;
        else if (totalMonthlyNetSalary < 15750) return 697.50;
        else if (totalMonthlyNetSalary < 16250) return 720.00;
        else if (totalMonthlyNetSalary < 16750) return 742.50;
        else if (totalMonthlyNetSalary < 17250) return 765.00;
        else if (totalMonthlyNetSalary < 17750) return 787.50;
        else if (totalMonthlyNetSalary < 18250) return 810.00;
        else if (totalMonthlyNetSalary < 18750) return 832.50;
        else if (totalMonthlyNetSalary < 19250) return 855.00;
        else if (totalMonthlyNetSalary < 19750) return 877.50;
        else if (totalMonthlyNetSalary < 20250) return 900.00;
        else if (totalMonthlyNetSalary < 20750) return 922.50;
        else if (totalMonthlyNetSalary < 21250) return 945.00;
        else if (totalMonthlyNetSalary < 21750) return 967.50;
        else if (totalMonthlyNetSalary < 22250) return 990.00;
        else if (totalMonthlyNetSalary < 22750) return 1012.50;
        else if (totalMonthlyNetSalary < 23250) return 1035.00;
        else if (totalMonthlyNetSalary < 23750) return 1057.50;
        else if (totalMonthlyNetSalary < 24250) return 1080.00;
        else if (totalMonthlyNetSalary < 24750) return 1102.50;
        else return 1125.00;
    }

    private static double calculatePhilHealth(double totalMonthlyNetSalary) {
    if (totalMonthlyNetSalary <= 10000) {
        return 300.00; // Minimum PhilHealth contribution
    } else if (totalMonthlyNetSalary <= 59999.99) {
        return totalMonthlyNetSalary * 0.03 / 2; // 3% of salary, split between employer & employee
    } else {
        return 1800.00; // Maximum PhilHealth contribution
    } 
}
    
    private static double calculateTax(double totalMonthlyNetSalary){
    if (totalMonthlyNetSalary <= 20832) {
        return 0;
    } else if (totalMonthlyNetSalary <= 33333) {
        return (totalMonthlyNetSalary - 20832) * 0.20;
    } else if (totalMonthlyNetSalary <= 66667) {
        return 2500 + (totalMonthlyNetSalary - 33333) * 0.25;
    } else if (totalMonthlyNetSalary <= 166667) {
        return 10833 + (totalMonthlyNetSalary - 66667) * 0.30;
    } else {
        return 40833 + (totalMonthlyNetSalary - 166667) * 0.32;
    }
    }
    
}