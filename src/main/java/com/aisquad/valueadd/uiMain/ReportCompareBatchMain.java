package com.aisquad.valueadd.uiMain;

import com.aisquad.valueadd.reportcompare.pdf.CompareResult;
import com.aisquad.valueadd.reportcompare.pdf.PdfComparator;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Iterator;

public class ReportCompareBatchMain {
    public static void main(String args[]) throws Exception {

        try{
            String strSourcePath="C:\\Users\\madhu\\OneDrive\\Documents\\GitHub\\ReportComparator\\input\\";
            String strTargetPath="C:\\Users\\madhu\\OneDrive\\Documents\\GitHub\\ReportComparator\\input\\";
            String strOutputPath="C:\\Users\\madhu\\OneDrive\\Documents\\GitHub\\ReportComparator\\output\\";
            String SAMPLE_XLSX_FILE_NAME =  "C:\\Users\\madhu\\OneDrive\\Documents\\GitHub\\ReportComparator\\input\\BatchModeInput.xlsx";
            String timeStamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(Calendar.getInstance().getTime());
            //Get workbook object
            Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_NAME));
            //read sheet object
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter dataFormatter = new DataFormatter();
            Iterator<Row> rowIterator = sheet.rowIterator();
            if(rowIterator.hasNext())
                rowIterator.next();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                // Now let's iterate over the columns of the current row
                Iterator<Cell> cellIterator = row.cellIterator();
                //iterating all rows and columns
//            while (cellIterator.hasNext()) {
//                Cell cell = cellIterator.next();
//                String cellValue = dataFormatter.formatCellValue(cell);
//                System.out.print(cellValue + " ");
//            }

                final CompareResult result = new PdfComparator<>
                        (
                                strSourcePath+dataFormatter.formatCellValue(row.getCell(0)),
                                strTargetPath+dataFormatter.formatCellValue(row.getCell(1))
                        ).compare();
                if (result.hasDifferenceInExclusion()) {
                    System.out.println("Only Differences in excluded areas found!");
                } else if (result.isNotEqual()) {
                    System.out.println("Differences found!");
                    timeStamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(Calendar.getInstance().getTime());
                    if (result.writeTo(strOutputPath+
                            dataFormatter.formatCellValue(row.getCell(0)).substring(0,dataFormatter.formatCellValue(row.getCell(0)).indexOf("."))+"_"+
                            dataFormatter.formatCellValue(row.getCell(1)).substring(0,dataFormatter.formatCellValue(row.getCell(1)).indexOf("."))+"_"+
                            timeStamp+".pdf")) {
                        System.out.println("Successfully write into file");
                    } else {
                        System.out.println("Error While writing into file.");
                    }
                } else { //         if (result.isEqual()) {
                    System.out.println("No Differences found!");
                }


                System.out.println();

            }
        } catch (IOException e) {
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


