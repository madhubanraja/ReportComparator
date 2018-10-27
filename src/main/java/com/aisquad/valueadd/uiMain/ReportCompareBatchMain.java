package com.aisquad.valueadd.uiMain;

import com.aisquad.valueadd.reportcompare.pdf.CompareResult;
import com.aisquad.valueadd.reportcompare.pdf.PdfComparator;
import org.apache.poi.ss.usermodel.*;

import java.awt.*;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Iterator;

public class ReportCompareBatchMain {
    static String strHtmlReportStart = "<html>\n" +"<body bgcolor=\"lightgrey\">\n" +"\t<center>\n" +"\t\t<b><u><font size=\"6\" align='center'>Report Comparator</font> <br><font size=\"5\" align='center'>Summary Report</font></u></b></br></br></br>\n" +"<table border=\"5\">\n" +"<tr><th><font size=\"5\">S.No</font></th><th>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font size=\"5\">Source</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</th><th>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font size=\"5\">Target</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</th><th>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font size=\"5\">ComparisonResult</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</th><th>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font size=\"5\">ComparisonOutput</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</th></tr>\n";
    static String strHtmlSerialNumber ="<tr><td><center>#Value#).</center></td>";
    static String strHtmlSourceFilePath ="\t<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=\"#Value#\" target=\"sourcetab\">#FileName#</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>\n";
    static String strHtmlTargetFilePath ="<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=\"#Value#\" target=\"targettab\">#FileName#</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>";
    static String strHtmlResultPASS ="<td><center><bold><font color=\"green\">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;PASS&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></bold></center></td>";
    static String strHtmlResultFAIL ="<td><center><bold><font color=\"red\">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;FAIL&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></bold></center></td>";
    static String strHtmlCompareisonResultFilePath ="<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=\"#Value#\" target=\"comparisonresulttab\">#FileName#</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>";
    static String strHtmlCompareisonResultFilePathEmpty ="<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;N/A&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>";
    static String strHtmlReportEnd ="</table></center></body></html>";
    static String strUserDirectory=System.getProperty("user.dir");
    static String strSourceFilePath="";
    static String strTargetFilePath="";
    static String strOutputFilePath="";
    static String strConsolidatedSummaryFileName =".\\ReportComparatorSummary.html";
    static String strOutputPath=strUserDirectory+"\\comparisonOutputPDFs\\";
    static String BATCH_XLSX_INPUT_FILE_NAME =  strUserDirectory+"\\BatchModeInputFile.xlsx";
    static String BATCH_HTML_SUMMARY_FILE_NAME =  strUserDirectory+"\\ReportComparatorSummaryFile.html";
    static int intSourceFilePathIndex=0;
    static int intSourceFileNameIndex=1;
    static int intTargetFilePathIndex=2;
    static int intTargetFileNameIndex=3;
    static boolean isDiffFound=true;
    static String strSummaryText="";
    static int intSno=1;
    static String strOutPutFileName="";

    public static void main(String args[]) throws Exception
    {
        try
        {
            File summaryHtmlFile = new File(BATCH_HTML_SUMMARY_FILE_NAME);
            FileWriter fileWriter = new FileWriter(summaryHtmlFile,false);
            fileWriter.write(strHtmlReportStart);
            String timeStamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(Calendar.getInstance().getTime());
            //Get workbook object
            Workbook workbook = WorkbookFactory.create(new File(BATCH_XLSX_INPUT_FILE_NAME));
            //read sheet object
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter dataFormatter = new DataFormatter();
            Iterator<Row> rowIterator = sheet.rowIterator();
            if(rowIterator.hasNext())
                rowIterator.next();
            while (rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                // Now let's iterate over the columns of the current row
                Iterator<Cell> cellIterator = row.cellIterator();
                //iterating all rows and columns
                strSourceFilePath=dataFormatter.formatCellValue(row.getCell(intSourceFilePathIndex))+dataFormatter.formatCellValue(row.getCell(intSourceFileNameIndex));
                strTargetFilePath=dataFormatter.formatCellValue(row.getCell(intTargetFilePathIndex))+dataFormatter.formatCellValue(row.getCell(intTargetFileNameIndex));
                System.out.println(strSourceFilePath);
                System.out.println(strTargetFilePath);
                final CompareResult result = new PdfComparator<>(strSourceFilePath,strTargetFilePath).compare();
                {
                    File output = new File(strOutputPath);
                    if(!output.exists())
                    {
                        output.mkdir();
                    }
                }
                if (result.hasDifferenceInExclusion())
                {
                    System.out.println("Only Differences in excluded areas found!");
                }
                else if (result.isNotEqual())
                {
                    System.out.println("Differences found!");


                    timeStamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(Calendar.getInstance().getTime());
                    strOutPutFileName=dataFormatter.formatCellValue(
                            row.getCell(intSourceFileNameIndex)).substring(0,dataFormatter.formatCellValue(row.getCell(intSourceFileNameIndex)).indexOf("."))+"_"+
                            dataFormatter.formatCellValue(row.getCell(intTargetFileNameIndex)).substring(0,dataFormatter.formatCellValue(row.getCell(intTargetFileNameIndex)).indexOf("."))+"_"+
                            timeStamp;
                    if (result.writeTo(strOutputPath+strOutPutFileName))
                    {
                        System.out.println("Successfully write into file");
                    } else
                        {
//                        System.out.println("Error While writing into file.");
                        }

                    isDiffFound = true;
                }
                else
                    { //         if (result.isEqual())
                    System.out.println("No Differences found!");
                    isDiffFound = false;
                    strOutPutFileName="";
                    }
                System.out.println();

                fileWriter.write(strHtmlSerialNumber.replace("#Value#",(intSno++)+""));
                fileWriter.write(strHtmlSourceFilePath.replace("#Value#",(strSourceFilePath)).replace("#FileName#",dataFormatter.formatCellValue(row.getCell(intSourceFileNameIndex))));
                fileWriter.write(strHtmlTargetFilePath.replace("#Value#",(strTargetFilePath)).replace("#FileName#",dataFormatter.formatCellValue(row.getCell(intTargetFileNameIndex))));
                if(isDiffFound) {
                    fileWriter.write(strHtmlResultFAIL);
                    fileWriter.write(strHtmlCompareisonResultFilePath.replace("#Value#",(strOutputPath+strOutPutFileName+".pdf")).replace("#FileName#",strOutPutFileName+".pdf"));
                }else {
                    fileWriter.write(strHtmlResultPASS);
                    fileWriter.write(strHtmlCompareisonResultFilePathEmpty);
                }

            }//end of batch while


            fileWriter.write(strHtmlReportEnd);

            fileWriter.flush();
            fileWriter.close();

            File htmlFile = new File(BATCH_HTML_SUMMARY_FILE_NAME);
            Desktop.getDesktop().browse(htmlFile.toURI());
        }//end of try
        catch (IOException e)
        {
            e.printStackTrace();
        }catch(Exception e)
        {
            e.printStackTrace();
        }
    }
}
         /*
// Creating a Workbook from an Excel file (.xls or .xlsx)


        // Retrieving the number of sheets in the Workbook
        */

        /*
           =============================================================
           Iterating over all the sheets in the workbook (Multiple ways)
           =============================================================
        */

// 1. You can obtain a sheetIterator and iterate over it
        /*Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        System.out.println("Retrieving Sheets using Iterator");
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            System.out.println("=> " + sheet.getSheetName());
        }

        // 2. Or you can use a for-each loop
        System.out.println("Retrieving Sheets using for-each loop");
        for(Sheet sheet: workbook) {
            System.out.println("=> " + sheet.getSheetName());
        }

        // 3. Or you can use a Java 8 forEach with lambda
        System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
        workbook.forEach(sheet -> {
            System.out.println("=> " + sheet.getSheetName());
        });*/

        /*
           ==================================================================
           Iterating over all the rows and columns in a Sheet (Multiple ways)
           ==================================================================
        */

// Getting the Sheet at index zero
//        lSheet sheet = workbook.getSheetAt(0);

// Create a DataFormatter to format and get each cell's value as String
//        DataFormatter dataFormatter = new DataFormatter();

// 1. You can obtain a rowIterator and columnIterator and iterate over them
        /*System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Now let's iterate over the columns of the current row
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }*/

// 2. Or you can use a for-each loop to iterate over the rows and columns
        /*System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
        for (Row row: sheet) {
            for(Cell cell: row) {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }*/

// 3. Or you can use Java 8 forEach loop with lambda
/*        System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
        sheet.forEach(row -> {
            row.forEach(cell -> {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            });
            System.out.println();
        });

        // Closing the workbook
        workbook.close();
    }*/
        /*try {

//            new PdfComparator("C:\\Raja\\IdeaProjects\\pdfcompare\\input\\A.pdf", "C:\\Raja\\IdeaProjects\\pdfcompare\\input\\A.pdf").compare().writeTo("C:\\Raja\\IdeaProjects\\pdfcompare\\output\\diffOutput.pdf");
            final CompareResult result = new PdfComparator<>(
                    "C:\\Raja\\IdeaProjects\\AiSquadReportCompare\\input\\அப்பா_src.pdf",
                    "C:\\Raja\\IdeaProjects\\AiSquadReportCompare\\input\\அப்பா_tgt.pdf").compare();
            if (result.isNotEqual()) {
                System.out.println("Differences found!");
                if(result.writeTo("C:\\Raja\\IdeaProjects\\AiSquadReportCompare\\output\\diffOutput.pdf"))
                {
                    System.out.println("Successfully write into file");
                }else
                {
                    System.out.println("Error While writing into file.");
                }
            }
            if (result.isEqual()) {
                System.out.println("No Differences found!");
            }
            if (result.hasDifferenceInExclusion()) {
                System.out.println("Only Differences in excluded areas found!");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }*/


