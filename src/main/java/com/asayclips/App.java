package com.asayclips;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * $ java -jar target/payroll-tool-1.0-SNAPSHOT.jar 04/08/2018 UT201 payroll.xls
 */
public class App 
{
    private static String _userHome = System.getProperty("user.home");

    public static void main( String[] args )
    {
        if (args.length != 3)
        {
            System.err.println("Usage: {begin MM/DD/YYYY} {store# e.g. UT104} {output}");
            System.exit(1);
        }
        String firstDay = args[0];
        String storeNumber = args[1];
        String payrollFile = args[2];

        App app = new App();
        app.generatePayroll(firstDay, storeNumber, payrollFile);
    }

    class Stylist
    {
        String name;
        double firstWeekOtherHours = 0.0;
        double firstWeekTotalHours = 0.0;
        double fullPeriodOtherHours;
        double fullPeriodTotalHours;
        double backBar;
        int serviceClients;
        double totalRetail;
        double totalService;
        double takeHomePerClient;
        double cutsPerHour;
        double tips;
        String bonusLevel = null;
        boolean foundMatch = false;

        public String toString()
        {
            return String.format(
                    "Name: %s, total hours: %5.2f, %5.2f, other: %5.2f, %5.2f, backbar: %2.0f, rpc: %d, retail: %5.2f, service: %5.2f, tips: %5.2f",
                    name, firstWeekTotalHours, fullPeriodTotalHours, firstWeekOtherHours, fullPeriodOtherHours,
                    backBar, serviceClients, totalRetail, totalService, tips);
        }
    }

    public void generatePayroll(String firstDay, String storeNumber, String payrollFile)
    {
        try
        {
            String lastDay = findOffsetDate(firstDay, 13);
            File firstWeekStylistAnalysis = findFirstWeekStylistAnalysis(firstDay, storeNumber);
            File fullPeriodStylistAnalysis = findStylistAnalysis(firstDay, lastDay, storeNumber);
            File tipsFile = findTips(firstDay, lastDay, storeNumber);

            Map<String, Stylist> stylists = readData(firstWeekStylistAnalysis, fullPeriodStylistAnalysis, tipsFile);

//            for (Stylist stylist : stylists.values())
//                System.out.println(stylist);

            String templateFile = String.format("%s/Documents/Payroll Templates/%s.xls", _userHome, storeNumber);
            buildPayrollFile(templateFile, payrollFile, stylists, lastDay);
        }
        catch (Exception e)
        {
            System.out.println("ERROR : " + e.toString());
            e.printStackTrace();
        }
    }

    private Map<String, Stylist> readData(File firstWeekStylistAnalysis, File fullPeriodStylistAnalysis, File tipsFile) throws Exception
    {
        Map<String, Stylist> stylists = readFirstWeekStylistAnalysis(firstWeekStylistAnalysis);
        stylists = readFullPeriodStylistAnalysis(fullPeriodStylistAnalysis, stylists);
        stylists = readTips(tipsFile, stylists);

        return stylists;
    }

    private File findFirstWeekStylistAnalysis(String firstDay, String storeNumber) throws Exception
    {
        return findStylistAnalysis(firstDay, findOffsetDate(firstDay, 6), storeNumber);
    }

    private File findStylistAnalysis(String firstDay, String lastDay, String storeNumber) throws Exception
    {
        File file = null;
        File dir = new File(_userHome + "/Downloads/");
        for (File f : dir.listFiles())
        {
            if (f.getName().startsWith("Stylist_Analysis")
                    && f.getName().endsWith(".xls")
                    && isTheRightStylistAnalysisFile(f, firstDay, lastDay, storeNumber))
            {
                System.out.printf("Found stylist analysis report for store: %s, (%s - %s) %s\n",
                        storeNumber, firstDay, lastDay, f.getAbsolutePath());
                file = f;
            }
        }
        if (file == null)
            System.out.printf("Download stylist analysis report for store: %s, (%s - %s)\n",
                    storeNumber, firstDay, lastDay);

        return file;
    }

    private boolean isTheRightStylistAnalysisFile(File file, String firstDay, String lastDay, String storeNumber) throws Exception
    {
        FileInputStream inputStream = new FileInputStream(file);
        try
        {
            HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
            HSSFSheet sheet = workbook.getSheet("Worksheet");

            HSSFRow row = sheet.getRow(1);
            if (!String.format("%s - %s", firstDay, lastDay).equals(row.getCell(1).getStringCellValue()))
                return false;

            return storeNumber.equals(row.getCell(2).getStringCellValue());
        }
        finally
        {
            inputStream.close();
        }
    }

    private static SimpleDateFormat _dateFormat = new SimpleDateFormat("MM/dd/yyyy");
    private static String findOffsetDate(String firstDay, int days) throws ParseException
    {
        Date d = new Date(_dateFormat.parse(firstDay).getTime() + days*24*60*60*1000);
        return _dateFormat.format(d);
    }

    private File findTips(String firstDay, String lastDay, String storeNumber) throws Exception
    {
        File file = null;
        File dir = new File(_userHome + "/Downloads/");
        for (File f : dir.listFiles())
        {
            if (f.getName().startsWith("Tips_By_Employee_Report")
                    && f.getName().endsWith(".xls")
                    && isTheRightTipsFile(f, firstDay, lastDay, storeNumber))
            {
                file = f;
                System.out.printf("Found tips by employee report for store: %s (%s - %s) %s\n",
                        storeNumber, firstDay, lastDay, f.getAbsolutePath());
            }
        }
        if (file == null)
            System.out.printf("Download tips by employee report for store: %s (%s - %s)\n",
                    storeNumber, firstDay, lastDay);

        return file;
    }

    private boolean isTheRightTipsFile(File file, String firstDay, String lastDay, String storeNumber) throws Exception
    {
        FileInputStream inputStream = new FileInputStream(file);
        try
        {
            HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
            HSSFSheet sheet = workbook.getSheet("Worksheet");

            HSSFRow row = sheet.getRow(1);
            if (!String.format("%s - %s", firstDay, lastDay).equals(row.getCell(1).getStringCellValue()))
                return false;

            row = sheet.getRow(0);
            return storeNumber.equals(row.getCell(1).getStringCellValue());
        }
        finally
        {
            inputStream.close();
        }
    }

    private void buildPayrollFile(String templateFilename, String payrollFilename, Map<String,Stylist> stylists, String lastDay) throws Exception
    {
        FileInputStream inputStream = new FileInputStream(new File(templateFilename));
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);

        setPayrollDate(workbook.getSheet("Information"), lastDay);

        HSSFSheet sheet = workbook.getSheet("Template");
        processManager(sheet, stylists);
        processStylists(sheet, stylists);
        processCoordinators(sheet, stylists);
        processHouseSales(sheet, stylists.get("business house"));
        inputStream.close();

        workbook.setForceFormulaRecalculation(true);
        FileOutputStream outputStream =new FileOutputStream(new File(payrollFilename));
        workbook.write(outputStream);
        outputStream.close();

        verifyEmployees(stylists);
    }

    private void setPayrollDate(HSSFSheet sheet, String date) throws ParseException
    {
        HSSFRow row = sheet.getRow(5);
        row.getCell(4).setCellValue(_dateFormat.parse(date));
    }

    private void processManager(HSSFSheet sheet, Map<String,Stylist> stylists)
    {
        HSSFRow row = sheet.getRow(8);
        String name = getNameFromRow(row, 0);
        Stylist manager = stylists.get(name.toLowerCase());
        if (manager == null)
        {
            System.out.printf("No match for manager (%s) found in stylist analysis reports.\n", name);
            return;
        }

        manager.foundMatch = true;

        row.getCell(1).setCellValue(manager.backBar);
        row.getCell(2).setCellValue(manager.tips);
        row.getCell(4).setCellValue(manager.fullPeriodOtherHours);
        row.getCell(6).setCellValue(manager.fullPeriodTotalHours);
        row.getCell(9).setCellValue(manager.serviceClients);
        row.getCell(10).setCellValue(manager.totalService);
        row.getCell(12).setCellValue(manager.totalRetail);

        row = sheet.getRow(9);
        row.getCell(1).setCellValue(stylists.get("total salon").backBar);
    }

    private void processStylists(HSSFSheet sheet, Map<String,Stylist> stylists)
    {
        for (int i=12; i<88; i+=3)
        {
            HSSFRow row = sheet.getRow(i);
            String name = getNameFromRow(row, 0);
            Stylist stylist = stylists.get(name.toLowerCase());
            if (stylist != null)
            {
                row.getCell(1).setCellValue(stylist.backBar);
                row.getCell(4).setCellValue(stylist.firstWeekOtherHours);
                row.getCell(6).setCellValue(stylist.firstWeekTotalHours);
                row.getCell(9).setCellValue(stylist.serviceClients);

                row = sheet.getRow(i + 1);
                row.getCell(4).setCellValue(stylist.fullPeriodOtherHours - stylist.firstWeekOtherHours);
                row.getCell(6).setCellValue(stylist.fullPeriodTotalHours - stylist.firstWeekTotalHours);

                row = sheet.getRow(i + 2);
                row.getCell(2).setCellValue(stylist.tips);
                row.getCell(10).setCellValue(stylist.totalService);
                row.getCell(12).setCellValue(stylist.totalRetail);
                if (stylist.bonusLevel != null)
                    row.getCell(1).setCellValue(stylist.bonusLevel);

                stylist.foundMatch = true;
            }
            else if (name != null && !name.equals("") && !name.contains("Stylist"))
                System.out.printf("No match for stylist (%s) found in stylist analysis report.\n", name);
        }
    }

    private void verifyEmployees(Map<String, Stylist> employees)
    {
        for (Stylist employee : employees.values())
        {
            if (!employee.foundMatch && !"Business House".equals(employee.name) && !"Total Salon".equals(employee.name))
                System.out.printf("No match for employee (%s) found in payroll template.\n", employee.name);
        }
    }

    private void processCoordinators(HSSFSheet sheet, Map<String,Stylist> stylists)
    {
        for (int i=91; i<104; i+=3)
        {
            HSSFRow row = sheet.getRow(i);
            String name = getNameFromRow(row, 0);
            Stylist receptionist = stylists.get(name.toLowerCase());
            if (receptionist != null)
            {
                row.getCell(4).setCellValue(receptionist.firstWeekTotalHours);

                row = sheet.getRow(i + 1);
                row.getCell(4).setCellValue(receptionist.fullPeriodTotalHours - receptionist.firstWeekTotalHours);

                receptionist.foundMatch = true;
            }
            else if (name != null && !name.equals("") && !name.contains("Coordinator"))
                System.out.printf("No match for coordinator (%s) found in stylist analysis report.\n", name);
        }
    }

    private String getNameFromRow(HSSFRow row, int cellNumber)
    {
        HSSFCell cell = row.getCell(cellNumber);
        if (cell == null)
            return "";

        return cleanupName(cell.getStringCellValue());
    }

    private void processHouseSales(HSSFSheet sheet, Stylist house)
    {
        HSSFRow row = sheet.getRow(106);
        row.getCell(9).setCellValue(house.serviceClients);
        row.getCell(10).setCellValue(house.totalService);
        row.getCell(12).setCellValue(house.totalRetail);
    }

    private Map<String, Stylist> readFirstWeekStylistAnalysis(File file) throws Exception
    {
        FileInputStream inputStream = new FileInputStream(file);
        try
        {
            Map<String, Stylist> stylists = new HashMap<String, Stylist>();
            HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
            HSSFSheet sheet = workbook.getSheet("Worksheet");

            int rowIndx = 4;
            HSSFRow row = sheet.getRow(rowIndx);
            while (true)
            {
                String firstname = row.getCell(1).getStringCellValue();
                if ("Total".equals(firstname))
                    break;

                Stylist stylist = new Stylist();
                String name = firstname.trim() + " " + row.getCell(2).getStringCellValue().trim();

                stylist.name = name;
                stylist.firstWeekTotalHours = row.getCell(10).getNumericCellValue();
                stylist.firstWeekOtherHours = row.getCell(12).getNumericCellValue();
                stylists.put(name.toLowerCase(), readFullPeriodRow(row, stylist));
                rowIndx++;
                row = sheet.getRow(rowIndx);
            }
            return stylists;
        }
        finally
        {
            inputStream.close();
        }
    }

    private Map<String, Stylist> readFullPeriodStylistAnalysis(File filename, Map<String, Stylist> stylists) throws Exception
    {
        FileInputStream inputStream = new FileInputStream(filename);
        try
        {
            HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
            HSSFSheet sheet = workbook.getSheet("Worksheet");

            int rowIndx = 3;
            HSSFRow row = sheet.getRow(rowIndx);
            while (true)
            {
                String firstname = row.getCell(1).getStringCellValue();
                String name = String.format("%s %s",
                        firstname.trim(), row.getCell(2).getStringCellValue().trim());
                Stylist stylist = stylists.get(name.toLowerCase());
                if (stylist == null)
                {
                    stylist = new Stylist();
                    stylist.name = name;
                    stylists.put(name.toLowerCase(), stylist);
                }
                readFullPeriodRow(row, stylist);
                getStylistBonusLevel(stylist);
                if ("Total".equals(firstname))
                    break;

                rowIndx++;
                row = sheet.getRow(rowIndx);
            }
            return stylists;
        }
        finally
        {
            inputStream.close();
        }
    }

    private Map<String, Stylist> readTips(File filename, Map<String, Stylist> stylists) throws Exception
    {
        FileInputStream inputStream = new FileInputStream(filename);
        try
        {
            HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
            HSSFSheet sheet = workbook.getSheet("Worksheet");

            int rowIndx = 7;
            HSSFRow row = sheet.getRow(rowIndx);
            while (row != null)
            {
                String name = getNameFromRow(row, 0);
                Stylist stylist = stylists.get(name.toLowerCase());
                if (stylist == null)
                {
                    stylist = new Stylist();
                    stylist.name = name;
                    stylists.put(name.toLowerCase(), stylist);
                }
                stylist.tips = row.getCell(3).getNumericCellValue();
                rowIndx++;
                row = sheet.getRow(rowIndx);
            }
            return stylists;
        }
        finally
        {
            inputStream.close();
        }
    }

    private static String cleanupName(String name)
    {
        return name.replaceAll("  ", " ").trim();
    }

    private Stylist readFullPeriodRow(HSSFRow row, Stylist stylist) {
        stylist.fullPeriodTotalHours = row.getCell(10).getNumericCellValue();
        stylist.fullPeriodOtherHours = row.getCell(12).getNumericCellValue();
        stylist.backBar = Math.round(row.getCell(21).getNumericCellValue()) / 100.0;
        stylist.serviceClients = (int)row.getCell(3).getNumericCellValue();
        stylist.totalRetail = row.getCell(8).getNumericCellValue();
        stylist.totalService = row.getCell(7).getNumericCellValue();
        stylist.takeHomePerClient = row.getCell(13).getNumericCellValue();
        stylist.cutsPerHour = row.getCell(16).getNumericCellValue();
        return stylist;
    }

    /**
     *          PaidBB  Take Home  CPTH
     *  Star      35%   $1.50     1.8
     *  All Star  40%   $1.75     2.0
     *  MVP       45%   $2.00     2.2
     *  Platinum  65%   $3.00     2.2
     *
     * @param stylist
     * @return
     */
    private static void getStylistBonusLevel(Stylist stylist)
    {
//        double takeHomePerClient = stylist.totalRetail / (double)stylist.serviceClients;
//        double cutsPerHour = (double)stylist.serviceClients / stylist.fullPeriodTotalHours;

        if (stylist.backBar >= 0.65 && stylist.takeHomePerClient >= 3.0 && stylist.cutsPerHour >= 2.2)
            stylist.bonusLevel = "Platinum";

        else if (stylist.backBar >= 0.45 && stylist.takeHomePerClient >= 2.0 && stylist.cutsPerHour >= 2.2)
            stylist.bonusLevel = "MVP";

        else if (stylist.backBar >= 0.4 && stylist.takeHomePerClient >= 1.75 && stylist.cutsPerHour >= 2.0)
            stylist.bonusLevel = "All Star";

        else if (stylist.backBar >= 0.35 && stylist.takeHomePerClient >= 1.5 && stylist.cutsPerHour >= 1.8)
            stylist.bonusLevel =  "Star";
    }
}
