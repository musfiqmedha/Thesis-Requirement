package thesis_requirements;

import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class Thesis_requirements 
{
    public static void main(String[] args) throws Exception 
    {
        
        File f = new File("C:\\Users\\Medha\\Desktop\\Project 250 - Thesis Requirement\\Sample Data.xlsx");//Select the path
        
        FileInputStream fis = new FileInputStream(f);
        
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        
        XSSFSheet sh1 = wb.getSheetAt(0);
        
        int rownum = sh1.getLastRowNum();
        
        System.out.println(rownum);
        
        int flag1, flag2, j, m;
        double max_cgpa = 0.00;
        m = 2;
        j = 1;
        String AsT = null;
        
        //Creating flag for all teachers
        int EH = 0, MHN = 0, MMR = 0, BPC = 0, MJ = 0, SNM = 0, MSI = 0, MER = 0, AT = 0, SS = 0, FR = 0, FC = 0, HAC = 0, MSC = 0, MM = 0, AAM = 0, MJI = 0, MRS = 0, MSR = 0, MZI = 0;
        
        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\Medha\\Desktop\\Project 250 - Thesis Requirement\\Assigned List.xlsx");
	XSSFWorkbook workbook = new XSSFWorkbook();
	XSSFSheet worksheet = workbook.createSheet("Newsheet");
        
        DataFormatter objDefaultFormat = new DataFormatter();
        FormulaEvaluator objFormulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
        
        XSSFRow row1 = worksheet.createRow(0);
        
        XSSFCell cellA1 = row1.createCell(0);
	cellA1.setCellValue("Team No.");
        XSSFCell cellB1 = row1.createCell(1);
	cellB1.setCellValue("CGPA of member 1");
        XSSFCell cellC1 = row1.createCell(2);
	cellC1.setCellValue("CGPA of member 2");
        XSSFCell cellD1 = row1.createCell(3);
	cellD1.setCellValue("Max CGPA");
        XSSFCell cellE1 = row1.createCell(4);
	cellE1.setCellValue("Assigned Teacher");
        
         
        for(int i = 1; i <= rownum; i++)
        {
            //Member 1 data input
            
            String data1 = sh1.getRow(i).getCell(1).getStringCellValue(); //Name
            int data2 = (int) sh1.getRow(i).getCell(2).getNumericCellValue(); //Reg
            String data3 = sh1.getRow(i).getCell(3).getStringCellValue();  //Mail
            double data4 = sh1.getRow(i).getCell(4).getNumericCellValue(); //Mobile
            double data5 = sh1.getRow(i).getCell(5).getNumericCellValue(); //Total_credit
            double data6 = sh1.getRow(i).getCell(6).getNumericCellValue(); //CGPA
            
            String data7 = sh1.getRow(i).getCell(7).getStringCellValue(); //Course
            String data8 = sh1.getRow(i).getCell(8).getStringCellValue(); //Course
            String data9 = sh1.getRow(i).getCell(9).getStringCellValue(); //Course
            String data10 = sh1.getRow(i).getCell(10).getStringCellValue(); //Course
            String data11 = sh1.getRow(i).getCell(11).getStringCellValue(); //Course
            String data12 = sh1.getRow(i).getCell(12).getStringCellValue(); //Course
            String data13 = sh1.getRow(i).getCell(13).getStringCellValue(); //Course
            String data14 = sh1.getRow(i).getCell(14).getStringCellValue(); //Course
            String data15 = sh1.getRow(i).getCell(15).getStringCellValue(); //Course
            String data16 = sh1.getRow(i).getCell(16).getStringCellValue(); //Course
            String data17 = sh1.getRow(i).getCell(17).getStringCellValue(); //Course
            String data18 = sh1.getRow(i).getCell(18).getStringCellValue(); //Course
            String data19 = sh1.getRow(i).getCell(19).getStringCellValue(); //Course
            String data20 = sh1.getRow(i).getCell(20).getStringCellValue(); //Course
            String data21 = sh1.getRow(i).getCell(21).getStringCellValue(); //Course
            String data22 = sh1.getRow(i).getCell(22).getStringCellValue(); //Course
            String data23 = sh1.getRow(i).getCell(23).getStringCellValue(); //Course
            String data24 = sh1.getRow(i).getCell(24).getStringCellValue(); //Course
            String data25 = sh1.getRow(i).getCell(25).getStringCellValue(); //Course
            String data26 = sh1.getRow(i).getCell(26).getStringCellValue(); //Course
            String data27 = sh1.getRow(i).getCell(27).getStringCellValue(); //Course
            String data28 = sh1.getRow(i).getCell(28).getStringCellValue(); //Course
            String data29 = sh1.getRow(i).getCell(29).getStringCellValue(); //Course
            String data30 = sh1.getRow(i).getCell(30).getStringCellValue(); //Course
            String data31 = sh1.getRow(i).getCell(31).getStringCellValue(); //Course
            
            
            //Member 2 data input
            
            String data32 = sh1.getRow(i).getCell(32).getStringCellValue(); //Name
            int data33 = (int) sh1.getRow(i).getCell(33).getNumericCellValue(); //Reg
            String data34 = sh1.getRow(i).getCell(34).getStringCellValue();  //Mail
            double data35 = sh1.getRow(i).getCell(35).getNumericCellValue(); //Mobile
            double data36 = sh1.getRow(i).getCell(36).getNumericCellValue(); //Total_credit
            double data37 = sh1.getRow(i).getCell(37).getNumericCellValue(); //CGPA
            
            String data38 = sh1.getRow(i).getCell(38).getStringCellValue(); //Course
            String data39 = sh1.getRow(i).getCell(39).getStringCellValue(); //Course
            String data40 = sh1.getRow(i).getCell(40).getStringCellValue(); //Course
            String data41 = sh1.getRow(i).getCell(41).getStringCellValue(); //Course
            String data42 = sh1.getRow(i).getCell(42).getStringCellValue(); //Course
            String data43 = sh1.getRow(i).getCell(43).getStringCellValue(); //Course
            String data44 = sh1.getRow(i).getCell(44).getStringCellValue(); //Course
            String data45 = sh1.getRow(i).getCell(45).getStringCellValue(); //Course
            String data46 = sh1.getRow(i).getCell(46).getStringCellValue(); //Course
            String data47 = sh1.getRow(i).getCell(47).getStringCellValue(); //Course
            String data48 = sh1.getRow(i).getCell(48).getStringCellValue(); //Course
            String data49 = sh1.getRow(i).getCell(49).getStringCellValue(); //Course
            String data50 = sh1.getRow(i).getCell(50).getStringCellValue(); //Course
            String data51 = sh1.getRow(i).getCell(51).getStringCellValue(); //Course
            String data52 = sh1.getRow(i).getCell(52).getStringCellValue(); //Course
            String data53 = sh1.getRow(i).getCell(53).getStringCellValue(); //Course
            String data54 = sh1.getRow(i).getCell(54).getStringCellValue(); //Course
            String data55 = sh1.getRow(i).getCell(55).getStringCellValue(); //Course
            String data56 = sh1.getRow(i).getCell(56).getStringCellValue(); //Course
            String data57 = sh1.getRow(i).getCell(57).getStringCellValue(); //Course
            String data58 = sh1.getRow(i).getCell(58).getStringCellValue(); //Course
            String data59 = sh1.getRow(i).getCell(59).getStringCellValue(); //Course
            String data60 = sh1.getRow(i).getCell(60).getStringCellValue(); //Course
            String data61 = sh1.getRow(i).getCell(61).getStringCellValue(); //Course
            String data62 = sh1.getRow(i).getCell(62).getStringCellValue(); //Course
            
            String data63 = sh1.getRow(i).getCell(63).getStringCellValue(); //Teacher 1
            String data64 = sh1.getRow(i).getCell(64).getStringCellValue(); //Teacher 2
            String data65 = sh1.getRow(i).getCell(65).getStringCellValue(); //Teacher 3
            
           
            flag1 = 0;
            flag2 = 0;
            
            if ((data6 >= 3.00) && 
                    (!data7.equals("F")) && (!data8.equals("F")) && (!data9.equals("F")) && (!data10.equals("F")) && (!data11.equals("F")) && (!data12.equals("F")) && (!data13.equals("F")) && (!data14.equals("F")) && (!data15.equals("F")) && (!data16.equals("F")) && (!data17.equals("F")) && (!data18.equals("F")) &&(!data19.equals("F")) && (!data20.equals("F")) && (!data21.equals("F")) && (!data22.equals("F")) && (!data23.equals("F")) && (!data24.equals("F")) && (!data25.equals("F")) && (!data26.equals("F")) && (!data27.equals("F")) && (!data28.equals("F")) && (!data29.equals("F")) && (!data30.equals("F")) && (!data31.equals("F"))
                ) 
                flag1 = 1;
                
            if (data37 >= 3.00 && 
                    (!data38.equals("F")) && (!data39.equals("F")) && (!data40.equals("F")) && (!data41.equals("F")) && (!data42.equals("F")) && (!data43.equals("F")) && (!data44.equals("F")) && (!data45.equals("F")) && (!data46.equals("F")) && (!data47.equals("F")) && (!data48.equals("F")) && (!data49.equals("F")) &&(!data50.equals("F")) && (!data51.equals("F")) && (!data52.equals("F")) && (!data53.equals("F")) && (!data54.equals("F")) && (!data55.equals("F")) && (!data56.equals("F")) && (!data57.equals("F")) && (!data58.equals("F")) && (!data59.equals("F")) && (!data60.equals("F")) && (!data61.equals("F")) && (!data62.equals("F"))
                ) 
                flag2 = 1;
            
            if(flag1 == 1 && flag2 == 1)
            {
                System.out.println("Team " + i + " is eligible for thesis.");
             
                XSSFRow row = worksheet.createRow(j);
                
                XSSFCell cellA = row.createCell(0);
                cellA.setCellValue(i);
                XSSFCell cellB = row.createCell(1);
                cellB.setCellValue(data6);
                XSSFCell cellC = row.createCell(2);
                cellC.setCellValue(data37);
                
                if(data6 > data37)
                    max_cgpa = data6;
                else
                    max_cgpa = data37;
                
                XSSFCell cellD = row.createCell(3);
                cellD.setCellValue(max_cgpa);
                
                //Assigning Teacher first time
                
                AsT = null;
                if(AsT == null)
                {
                    if(data63.equals("EH") && (EH < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        EH++;
                    }
                    else if(data63.equals("MHN") && (MHN < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MHN++;
                    }
                    else if(data63.equals("MMR") && (MMR < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MMR++;
                    }
                    else if(data63.equals("BPC") && (BPC < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        BPC++;
                    }
                    else if(data63.equals("MJ") && (MJ < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MJ++;
                    }
                    else if(data63.equals("SNM") && (SNM < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        SNM++;
                    }
                    else if(data63.equals("MSI") && (MSI < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MSI++;
                    }
                    else if(data63.equals("MER") && (MER < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MER++;
                    }
                    else if(data63.equals("AT") && (AT < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        AT++;
                    }
                    else if(data63.equals("SS") && (SS < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        SS++;
                    }
                    else if(data63.equals("FR") && (FR < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        FR++;
                    }
                    else if(data63.equals("FC") && (FC < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        FC++;
                    }
                    else if(data63.equals("HAC") && (HAC < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        HAC++;
                    }
                    else if(data63.equals("MSC") && (MSC < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MSC++;
                    }
                    else if(data63.equals("MM") && (MM < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MM++;
                    }
                    else if(data63.equals("AAM") && (AAM < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        AAM++;
                    }
                    else if(data63.equals("MJI") && (MJI < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MJI++;
                    }
                    else if(data63.equals("MRS") && (MRS < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MRS++;
                    }
                    else if(data63.equals("MSR") && (MSR < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MSR++;
                    }
                    else if(data63.equals("MZI") && (MZI < 2))
                    {
                        AsT = data63;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MZI++;
                    }
                }
                //Assigning Teacher second time
                if(AsT == null)
                {
                    if(data64.equals("EH") && (EH < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        EH++;
                    }
                    else if(data64.equals("MHN") && (MHN < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MHN++;
                    }
                    else if(data64.equals("MMR") && (MMR < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MMR++;
                    }
                    else if(data64.equals("BPC") && (BPC < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        BPC++;
                    }
                    else if(data64.equals("MJ") && (MJ < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MJ++;
                    }
                    else if(data64.equals("SNM") && (SNM < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        SNM++;
                    }
                    else if(data64.equals("MSI") && (MSI < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MSI++;
                    }
                    else if(data64.equals("MER") && (MER < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MER++;
                    }
                    else if(data64.equals("AT") && (AT < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        AT++;
                    }
                    else if(data64.equals("SS") && (SS < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        SS++;
                    }
                    else if(data64.equals("FR") && (FR < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        FR++;
                    }
                    else if(data64.equals("FC") && (FC < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        FC++;
                    }
                    else if(data64.equals("HAC") && (HAC < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        HAC++;
                    }
                    else if(data64.equals("MSC") && (MSC < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MSC++;
                    }
                    else if(data64.equals("MM") && (MM < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MM++;
                    }
                    else if(data64.equals("AAM") && (AAM < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        AAM++;
                    }
                    else if(data64.equals("MJI") && (MJI < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MJI++;
                    }
                    else if(data64.equals("MRS") && (MRS < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MRS++;
                    }
                    else if(data64.equals("MSR") && (MSR < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MSR++;
                    }
                    else if(data64.equals("MZI") && (MZI < 2))
                    {
                        AsT = data64;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MZI++;
                    }
                }
                //Assigning Teacher third time
                if(AsT == null)
                {
                    if(data65.equals("EH") && (EH < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        EH++;
                    }
                    else if(data65.equals("MHN") && (MHN < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MHN++;
                    }
                    else if(data65.equals("MMR") && (MMR < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MMR++;
                    }
                    else if(data65.equals("BPC") && (BPC < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        BPC++;
                    }
                    else if(data65.equals("MJ") && (MJ < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MJ++;
                    }
                    else if(data65.equals("SNM") && (SNM < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        SNM++;
                    }
                    else if(data65.equals("MSI") && (MSI < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MSI++;
                    }
                    else if(data65.equals("MER") && (MER < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MER++;
                    }
                    else if(data65.equals("AT") && (AT < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        AT++;
                    }
                    else if(data65.equals("SS") && (SS < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        SS++;
                    }
                    else if(data65.equals("FR") && (FR < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        FR++;
                    }
                    else if(data65.equals("FC") && (FC < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        FC++;
                    }
                    else if(data65.equals("HAC") && (HAC < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        HAC++;
                    }
                    else if(data65.equals("MSC") && (MSC < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MSC++;
                    }
                    else if(data65.equals("MM") && (MM < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MM++;
                    }
                    else if(data65.equals("AAM") && (AAM < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        AAM++;
                    }
                    else if(data65.equals("MJI") && (MJI < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MJI++;
                    }
                    else if(data65.equals("MRS") && (MRS < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MRS++;
                    }
                    else if(data65.equals("MSR") && (MSR < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MSR++;
                    }
                    else if(data65.equals("MZI") && (MZI < 2))
                    {
                        AsT = data65;
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MZI++;
                    }
                }
                if(AsT == null)
                {
                    if(EH < 2)
                    {
                        AsT = "EH";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        EH++;
                    }
                    else if(MHN < 2)
                    {
                        AsT = "MHN";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MHN++;
                    }
                    else if(MMR < 2)
                    {
                        AsT = "MMR";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MMR++;
                    }
                    else if(BPC < 2)
                    {
                        AsT = "BPC";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        BPC++;
                    }
                    else if(MJ < 2)
                    {
                        AsT = "MJ";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MJ++;
                    }
                    else if(SNM < 2)
                    {
                        AsT = "SNM";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        SNM++;
                    }
                    else if(MSI < 2)
                    {
                        AsT = "MSI";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MSI++;
                    }
                    else if(MER < 2)
                    {
                        AsT = "MER";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MER++;
                    }
                    else if(AT < 2)
                    {
                        AsT = "AT";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        AT++;
                    }
                    else if(SS < 2)
                    {
                        AsT = "SS";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        SS++;
                    }
                    else if(FR < 2)
                    {
                        AsT = "FR";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        FR++;
                    }
                    else if(FC < 2)
                    {
                        AsT = "FC";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        FC++;
                    }
                    else if(HAC < 2)
                    {
                        AsT = "HAC";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        HAC++;
                    }
                    else if(MSC < 2)
                    {
                        AsT = "MSC";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MSC++;
                    }
                    else if(MM < 2)
                    {
                        AsT = "MM";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MM++;
                    }
                    else if(AAM < 2)
                    {
                        AsT = "AAM";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        AAM++;
                    }
                    else if(MJI < 2)
                    {
                        AsT = "MJI";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MJI++;
                    }
                    else if(MRS < 2)
                    {
                        AsT = "MRS";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MRS++;
                    }
                    else if(MSR < 2)
                    {
                        AsT = "MSR";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MSR++;
                    }
                    else if(MZI < 2)
                    {
                        AsT = "MZI";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                        MZI++;
                    }
                    else
                    {
                        AsT = "ERROR";
                        XSSFCell cellE = row.createCell(4);
                        cellE.setCellValue(AsT);
                    }
                }
                j++;   
            }
            else if(flag1 == 1 && flag2 == 0)
            {
                System.out.println("Team " + i + " is not eligible for thesis. Because member 2 Reg. No. " + data33 + " hasn't fulfilled the requirements.");
            }
            else if(flag1 == 0 && flag2 == 1)
            {
                System.out.println("Team " + i + " is not eligible for thesis. Because member 1 Reg. No. " + data2 + " hasn't fulfilled the requirements.");
            }
            else if(flag1 == 0 && flag2 == 0)
            {
                System.out.println("Team " + i + " is not eligible for thesis. Because none of them (Reg. No. " + data2 + " and " + data33 + ") fulfilled the requirements.");
            }
        }
        System.out.println("\n");
        System.out.println("Total eligible team : " + (j - 1) + "\n");
        workbook.write(fileOut);
	fileOut.flush();
	fileOut.close();
    } 
    catch (Exception e)
    {
        System.out.println(e.getMessage());
    }  
}