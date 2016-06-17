package com.rao.excel;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ExcelReader
{
    public List<HashMap<String, String>> read(String inputFile) throws IOException
    {
        List<HashMap<String, String>> excelData = new ArrayList<HashMap<String, String>>();
        File inputWorkbook = new File(inputFile);
        Workbook w;
        try
        {
            w = Workbook.getWorkbook(inputWorkbook);
            // Get the first sheet
            Sheet sheet = w.getSheet(0);
            // Loop over first 10 column and lines
            for(int i = 0; i < sheet.getRows(); i++)
            {
                HashMap<String, String> entityMap = new HashMap<String, String>();
                Cell[] cols = sheet.getRow(i);
                for(int j = 0; j < cols.length; j++)
                {
                    Cell cell = cols[j];
                    switch(j)
                    {
                        case 0:
                            entityMap.put("ENTITY", cell.getContents());
                            break;
                        case 1:
                            entityMap.put("FILE", "c:/test/out/"+cell.getContents());
                            break;
                        case 2:
                            entityMap.put("QUERY", cell.getContents());
                            break;
                        default:
                            break;
                    }
                }
                excelData.add(entityMap);
            }
        }
        catch(BiffException e)
        {
            e.printStackTrace();
        }
        return excelData;
    }    
}