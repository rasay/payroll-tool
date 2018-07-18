package com.asayclips;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;

public class ExcelUtils
{
    public static String getNameFromRow(HSSFRow row, int cellNumber)
    {
        HSSFCell cell = row.getCell(cellNumber);
        if (cell == null)
            return "";

        return cleanupName(cell.getStringCellValue());
    }

    private static String cleanupName(String name)
    {
        return name.replaceAll("  ", " ").trim();
    }
}
