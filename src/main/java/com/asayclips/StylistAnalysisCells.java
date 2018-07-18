package com.asayclips;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;

public class StylistAnalysisCells
{
    public int firstName;
    public int lastName;
    public int backBar;
    public int otherHours;
    public int totalHours;
    public int serviceClients;
    public int serviceSales;
    public int retailSales;
    public int takeHomePerClient;
    public int cutsPerHour;

    public StylistAnalysisCells() {}

    public void init(HSSFSheet sheet)
    {
        HSSFRow row = sheet.getRow(2);

        for (int i=0; i<30; i++)
        {
            String cellName = ExcelUtils.getNameFromRow(row, i);
            if ("First".equals(cellName))
                firstName = i;
            else if ("Last".equals(cellName))
                lastName = i;
            else if ("% Paid BB".equals(cellName))
                backBar = i;
            else if ("Non-Store Hours".equals(cellName))
                otherHours = i;
            else if ("Total Hours".equals(cellName))
                totalHours = i;
            else if ("Service Clients".equals(cellName))
                serviceClients = i;
            else if ("Service Sales".equals(cellName))
                serviceSales = i;
            else if ("Retail Sales".equals(cellName))
                retailSales = i;
            else if ("Retail $/ Client".equals(cellName))
                takeHomePerClient =i;
            else if ("Clients / Floor Hour".equals(cellName))
                cutsPerHour = i;
        }
    }
}
