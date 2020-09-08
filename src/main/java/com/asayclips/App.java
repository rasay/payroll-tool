package com.asayclips;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

import javafx.scene.control.TextArea;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * $ java -jar target/payroll-tool-1.0-SNAPSHOT.jar 04/08/2018 UT201 payroll.xls
 */
public class App 
{
    private static String _userHome = System.getProperty("user.home");
    private TextArea _messageArea;

    public App(TextArea messageArea)
    {
        _messageArea = messageArea;
    }

    public static void main( String[] args )
    {
        if (args.length != 2)
        {
            System.err.println("Usage: {begin MM/DD/YYYY} {store# e.g. UT104}");
            System.exit(1);
        }
        String firstDay = args[0];
        String storeNumber = args[1];

        App app = new App(null);

        app.generatePayroll(firstDay, storeNumber);
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
        boolean foundMatch = false;

        public boolean isPerson()
        {
            return (!"Business House".equals(name) && !"Total Salon".equals(name));
        }

        public String toString()
        {
            return String.format(
                    "Name: %s, total hours: %5.2f, %5.2f, other: %5.2f, %5.2f, backbar: %2.0f, rpc: %d, retail: %5.2f, service: %5.2f, tips: %5.2f",
                    name, firstWeekTotalHours, fullPeriodTotalHours, firstWeekOtherHours, fullPeriodOtherHours,
                    backBar, serviceClients, totalRetail, totalService, tips);
        }
    }

    public void generatePayroll(String endDate, String storeNumber)
    {
        try
        {
            displayMessage("### Processing: " + storeNumber + ", " + endDate);
            // ensure correct format for storeNumber and endDate
            endDate = findOffsetDate(endDate, 0);
            storeNumber = storeNumber.trim().toUpperCase();

            String beginDate = findOffsetDate(endDate, -13);

            String outputFilename = String.format("%s/Documents/Payroll/%s_%s.xls",
                    _userHome, storeNumber, endDate.replaceAll("/", "-"));

            File firstWeekStylistAnalysis = findFirstWeekStylistAnalysis(beginDate, storeNumber);
            File fullPeriodStylistAnalysis = findStylistAnalysis(beginDate, endDate, storeNumber);
            File tipsFile = findTips(beginDate, endDate, storeNumber);

            if (firstWeekStylistAnalysis == null || fullPeriodStylistAnalysis == null || tipsFile == null)
                displayMessage("### Failure!!!");
            else
            {
                Map<String, Stylist> stylists = readData(firstWeekStylistAnalysis, fullPeriodStylistAnalysis, tipsFile);
                String template2File = String.format("%s/Documents/Payroll Templates/%s.xls", _userHome, storeNumber);
                populateTemplate(template2File, outputFilename, stylists, endDate);
                displayMessage("### Success!!!");
            }
        }
        catch (Exception e)
        {
            displayMessage("ERROR : " + e.toString());
            e.printStackTrace();
        }
    }

    private void displayMessage(String message)
    {
        if (_messageArea == null)
            System.out.println(message);
        else
            _messageArea.appendText(message + "\n");
    }

    private Map<String, Stylist> readData(File firstWeekStylistAnalysis, File fullPeriodStylistAnalysis, File tipsFile) throws Exception
    {
        Map<String, Stylist> stylists = readFirstWeekStylistAnalysis(firstWeekStylistAnalysis);
        stylists = readFullPeriodStylistAnalysis(fullPeriodStylistAnalysis, stylists);
        stylists = readTips(tipsFile, stylists);

        return stylists;
    }

    private File findFirstWeekStylistAnalysis(String beginDate, String storeNumber) throws Exception
    {
        return findStylistAnalysis(beginDate, findOffsetDate(beginDate, 6), storeNumber);
    }

    private File findStylistAnalysis(String beginDate, String endDate, String storeNumber) throws Exception
    {
        File file = null;
        File dir = new File(_userHome + "/Downloads/");
        for (File f : dir.listFiles())
        {
            if (f.getName().startsWith("Stylist_Analysis")
                    && f.getName().endsWith(".xls")
                    && isTheRightStylistAnalysisFile(f, beginDate, endDate, storeNumber))
            {
                displayMessage(String.format("Found stylist analysis report for store: %s, (%s - %s) %s",
                        storeNumber, beginDate, endDate, f.getAbsolutePath()));
                file = f;
                break;
            }
        }
        if (file == null)
            displayMessage(String.format("Download stylist analysis report for store: %s, (%s - %s)",
                    storeNumber, beginDate, endDate));

        return file;
    }

    private boolean isTheRightStylistAnalysisFile(File file, String beginDate, String endDate, String storeNumber) throws Exception
    {
        FileInputStream inputStream = new FileInputStream(file);
        try
        {
            HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
            HSSFSheet sheet = workbook.getSheet("Worksheet");

            HSSFRow row = sheet.getRow(1);
            if (!String.format("%s - %s", beginDate, endDate).equals(row.getCell(1).getStringCellValue()))
                return false;

            return storeNumber.equals(row.getCell(2).getStringCellValue());
        }
        finally
        {
            inputStream.close();
        }
    }

    private static final long MILISECONDS_PER_HOUR = 1000 * 60 * 60;
    private static final long MILISECONDS_PER_DAY = 24 * MILISECONDS_PER_HOUR;
    private static SimpleDateFormat _dateFormat = new SimpleDateFormat("MM/dd/yyyy");
    public static String findOffsetDate(String baseDate, int days) throws ParseException
    {
        Date d = new Date(_dateFormat.parse(baseDate).getTime() + days*MILISECONDS_PER_DAY + MILISECONDS_PER_HOUR);
        return _dateFormat.format(d);
    }

    private File findTips(String beginDate, String endDate, String storeNumber) throws Exception
    {
        File file = null;
        File dir = new File(_userHome + "/Downloads/");
        for (File f : dir.listFiles())
        {
            if (f.getName().startsWith("Tips_By_Employee_Report")
                    && f.getName().endsWith(".xls")
                    && isTheRightTipsFile(f, beginDate, endDate, storeNumber))
            {
                file = f;
                displayMessage(String.format("Found tips by employee report for store: %s (%s - %s) %s",
                        storeNumber, beginDate, endDate, f.getAbsolutePath()));
            }
        }
        if (file == null)
            displayMessage(String.format("Download tips by employee report for store: %s (%s - %s)",
                    storeNumber, beginDate, endDate));

        return file;
    }

    private boolean isTheRightTipsFile(File file, String beginDate, String endDate, String storeNumber) throws Exception
    {
        FileInputStream inputStream = new FileInputStream(file);
        try
        {
            HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
            HSSFSheet sheet = workbook.getSheet("Worksheet");

            HSSFRow row = sheet.getRow(1);
            if (!String.format("%s - %s", beginDate, endDate).equals(row.getCell(1).getStringCellValue()))
                return false;

            row = sheet.getRow(0);
            return storeNumber.equals(row.getCell(1).getStringCellValue());
        }
        finally
        {
            inputStream.close();
        }
    }

    private void populateTemplate(String templateFilename, String payrollFilename, Map<String,Stylist> stylists, String endDate) throws Exception
    {
        FileInputStream inputStream = new FileInputStream(new File(templateFilename));
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);

        setPayrollDate(workbook.getSheet("Information"), endDate);

        HSSFSheet sheet = workbook.getSheet("Template");
        processManager(sheet, stylists);
        processStylists(sheet, stylists);
        processCoordinators(sheet, stylists);
        processUnfoundStylists(sheet, stylists);
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
        String name = ExcelUtils.getNameFromRow(row, 0);
        Stylist manager = stylists.get(name.toLowerCase());
        if (manager == null)
        {
            displayMessage(String.format("No match for manager (%s) found in stylist analysis reports.", name));
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
            String name = ExcelUtils.getNameFromRow(row, 0);
            Stylist stylist = stylists.get(name.toLowerCase());
            if (stylist != null)
            {
                processStylist(sheet, row, stylist);
                stylist.foundMatch = true;
            }
            else if (isEmployeeName(name))
                displayMessage(String.format("No match for stylist (%s) found in stylist analysis report.", name));
        }
    }

    private void processStylist(HSSFSheet sheet, HSSFRow row, Stylist stylist)
    {
        row.getCell(1).setCellValue(stylist.backBar);
        row.getCell(4).setCellValue(stylist.firstWeekOtherHours);
        row.getCell(6).setCellValue(stylist.firstWeekTotalHours);
        row.getCell(9).setCellValue(stylist.serviceClients);

        row = sheet.getRow(row.getRowNum() + 1);
        row.getCell(4).setCellValue(stylist.fullPeriodOtherHours - stylist.firstWeekOtherHours);
        row.getCell(6).setCellValue(stylist.fullPeriodTotalHours - stylist.firstWeekTotalHours);

        row = sheet.getRow(row.getRowNum() + 1);
        row.getCell(2).setCellValue(stylist.tips);
        row.getCell(10).setCellValue(stylist.totalService);
        row.getCell(12).setCellValue(stylist.totalRetail);
    }

    private void processUnfoundStylists(HSSFSheet sheet, Map<String,Stylist> stylists)
    {
        for (Stylist stylist : stylists.values())
        {
            if (!stylist.foundMatch && stylist.isPerson())
            {
                HSSFRow row = findEmptyStylistRow(sheet);
                row.createCell(0).setCellValue(stylist.name);
                processStylist(sheet, row, stylist);
            }
        }
    }

    private HSSFRow findEmptyStylistRow(HSSFSheet sheet)
    {
        for (int i=12; i<88; i+=3)
        {
            HSSFRow row = sheet.getRow(i);
            String name = ExcelUtils.getNameFromRow(row, 0);
            displayMessage(String.format("Is name '%s'", name));
            if (!isEmployeeName(name))
                return row;
        }
        return null; // shouldn't ever happen
    }

    private boolean isEmployeeName(String name)
    {
        return (!StringUtils.isBlank(name) && !name.contains("Stylist"));
    }

    private void verifyEmployees(Map<String, Stylist> employees)
    {
        for (Stylist employee : employees.values())
        {
            if (!employee.foundMatch && employee.isPerson())
                displayMessage(String.format("No match for employee (%s) found in payroll template.", employee.name));
        }
    }

    private void processCoordinators(HSSFSheet sheet, Map<String,Stylist> stylists)
    {
        for (int i=91; i<104; i+=3)
        {
            HSSFRow row = sheet.getRow(i);
            String name = ExcelUtils.getNameFromRow(row, 0);
            Stylist receptionist = stylists.get(name.toLowerCase());
            if (receptionist != null)
            {
                row.getCell(4).setCellValue(receptionist.firstWeekTotalHours);

                row = sheet.getRow(i + 1);
                row.getCell(4).setCellValue(receptionist.fullPeriodTotalHours - receptionist.firstWeekTotalHours);

                receptionist.foundMatch = true;
            }
            else if (name != null && !name.equals("") && !name.contains("Coordinator"))
                displayMessage(String.format("No match for coordinator (%s) found in stylist analysis report.", name));
        }
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
            StylistAnalysisCells cells = new StylistAnalysisCells();
            cells.init(sheet);

            int rowIndx = 4;
            HSSFRow row = sheet.getRow(rowIndx);
            while (true)
            {
                String firstname = row.getCell(cells.firstName).getStringCellValue();
                if ("Total".equals(firstname))
                    break;

                Stylist stylist = new Stylist();
                String name = firstname.trim() + " " + row.getCell(cells.lastName).getStringCellValue().trim();

                stylist.name = name;
                stylist.firstWeekTotalHours = row.getCell(cells.totalHours).getNumericCellValue();
                stylist.firstWeekOtherHours = row.getCell(cells.otherHours).getNumericCellValue();
                stylists.put(name.toLowerCase(), readFullPeriodRow(cells, row, stylist));
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
            StylistAnalysisCells cells = new StylistAnalysisCells();
            cells.init(sheet);

            int rowIndx = 3;
            HSSFRow row = sheet.getRow(rowIndx);
            while (true)
            {
                String firstname = row.getCell(cells.firstName).getStringCellValue();
                String name = String.format("%s %s",
                        firstname.trim(), row.getCell(cells.lastName).getStringCellValue().trim());
                Stylist stylist = stylists.get(name.toLowerCase());
                if (stylist == null)
                {
                    stylist = new Stylist();
                    stylist.name = name;
                    stylists.put(name.toLowerCase(), stylist);
                }
                readFullPeriodRow(cells, row, stylist);
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
                String name = ExcelUtils.getNameFromRow(row, 0);
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

    private Stylist readFullPeriodRow(StylistAnalysisCells cells, HSSFRow row, Stylist stylist) {
        stylist.fullPeriodTotalHours = row.getCell(cells.totalHours).getNumericCellValue();
        stylist.fullPeriodOtherHours = row.getCell(cells.otherHours).getNumericCellValue();
        stylist.backBar = Math.round(row.getCell(cells.backBar).getNumericCellValue()) / 100.0;
        stylist.serviceClients = (int)row.getCell(cells.serviceClients).getNumericCellValue();
        stylist.totalRetail = row.getCell(cells.retailSales).getNumericCellValue();
        stylist.totalService = row.getCell(cells.serviceSales).getNumericCellValue();
        stylist.takeHomePerClient = row.getCell(cells.takeHomePerClient).getNumericCellValue();
        stylist.cutsPerHour = row.getCell(cells.cutsPerHour).getNumericCellValue();
        return stylist;
    }
}
