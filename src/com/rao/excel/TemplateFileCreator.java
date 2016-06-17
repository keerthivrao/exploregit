package com.rao.excel;

import java.util.HashMap;
import java.util.List;

public class TemplateFileCreator
{
    /**
     * @param args
     */
    public static void main(String[] args) throws Exception
    {
        ExcelReader excelReader = new ExcelReader();
        List<HashMap<String, String>> excelData = excelReader.read("c:/TEMP/in.xls");
        
        ExcelWriter excelWriter = new ExcelWriter();
        
        for(int index = 0; index < excelData.size(); index++)
        {
            excelWriter.createTemplate(excelData.get(index));
        }
        
    }
}
