package com.aisquad.valueadd.uiMain;

import com.aisquad.valueadd.reportcompare.pdf.CompareResult;
import com.aisquad.valueadd.reportcompare.pdf.PdfComparator;

import java.io.IOException;

public class ReportComparePdfMain
{
    public static void main(String args[])
    {
        try {
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
        }
    }
}
