package com.rao.excel;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ExcelWriter
{
    //private WritableCellFormat timesBoldUnderline;

    private WritableCellFormat times;

    private List<String>       headerLabels = new ArrayList<String>();

    public void createTemplate(HashMap<String, String> entityParams) throws IOException, WriteException
    {
        System.out.println(entityParams);
        if(entityParams.get("ENTITY") == null || entityParams.get("ENTITY").trim().equals(""))
        {
            return;
        }
        File file = new File(entityParams.get("FILE"));
        System.out.println("PATH() "+file.getAbsolutePath());
        System.out.println("Directory() "+file.isDirectory());
        System.out.println("File() "+file.isFile());
        WorkbookSettings wbSettings = new WorkbookSettings();
        wbSettings.setLocale(new Locale("en", "EN"));
        WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
        workbook.createSheet("Sheet1", 0);
        WritableSheet excelSheet = workbook.getSheet(0);
        createLabel(excelSheet, entityParams);
        createContent(excelSheet);
        workbook.write();
        workbook.close();
    }

    private void createLabel(WritableSheet sheet, HashMap<String, String> entityParams) throws WriteException
    {
        // Lets create a times font
        WritableFont times10pt = new WritableFont(WritableFont.TIMES, 10);
        // Define the cell format
        times = new WritableCellFormat(times10pt);
        // Lets automatically wrap the cells
        times.setWrap(true);
        // Create create a bold font with unterlines
        //WritableFont times10ptBoldUnderline = new WritableFont(WritableFont.TIMES, 10, WritableFont.BOLD, false, UnderlineStyle.SINGLE);
        //timesBoldUnderline = new WritableCellFormat(times10ptBoldUnderline);
        // Lets automatically wrap the cells
        //timesBoldUnderline.setWrap(true);
        CellView cv = new CellView();
        cv.setFormat(times);
        //cv.setFormat(timesBoldUnderline);
        cv.setAutosize(true);
        getHeaders(entityParams);
        for(int index = 0; index < headerLabels.size(); index++)
        {
            addCaption(sheet, index, 0, headerLabels.get(index));
        }
    }

    private void getHeaders(HashMap<String, String> entityParams)
    {
        Connection connection = null;
        Statement statement = null;
        ResultSet resultSet = null;
        try
        {
            connection = getConnection();
            statement = connection.createStatement();
            resultSet = statement.executeQuery(entityParams.get("QUERY"));
            ResultSetMetaData metaData = resultSet.getMetaData();
            int cols = metaData.getColumnCount();
            for(int index = 1; index <= cols; index++)
            {
                String colName = metaData.getColumnLabel(index);
                if(!headerLabels.contains(colName) && !( "statusid".equalsIgnoreCase(colName) || "statusrfcid".equalsIgnoreCase(colName) ))
                {
                    headerLabels.add(colName);
                }
            }
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
        finally
        {
            try
            {
                resultSet.close();
                statement.close();
                connection.close();
            }
            catch(Exception e)
            {
                e.printStackTrace();
            }
        }
    }

    private void createContent(WritableSheet sheet) throws WriteException, RowsExceededException
    {
        for(int index = 0; index < headerLabels.size(); index++)
        {
            addLabel(sheet, index, 1, "%%=Table." + headerLabels.get(index));
        }
    }

    private void addCaption(WritableSheet sheet, int column, int row, String s) throws RowsExceededException, WriteException
    {
        Label label;
        label = new Label(column, row, s, times);
        sheet.addCell(label);
    }

    private void addNumber(WritableSheet sheet, int column, int row, Integer integer) throws WriteException, RowsExceededException
    {
        Number number;
        number = new Number(column, row, integer, times);
        sheet.addCell(number);
    }

    private void addLabel(WritableSheet sheet, int column, int row, String s) throws WriteException, RowsExceededException
    {
        Label label;
        label = new Label(column, row, s, times);
        sheet.addCell(label);
    }

    private Connection getConnection()
    {
        String dbUrl = "jdbc:sqlserver://localhost:1433;databaseName=IOLSSOT_Playground;integratedSecurity=false;";
        String dbdriver = "com.microsoft.sqlserver.jdbc.SQLServerDriver";
        Connection connection = null;
        try
        {
            Class.forName(dbdriver);
            connection = DriverManager.getConnection(dbUrl, "ssot_usr", "SSOT123!");
        }
        catch(Exception e)
        {
        }
        return connection;
    }
}
