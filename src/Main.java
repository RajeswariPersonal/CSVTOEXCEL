import java.io.*;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import  org.apache.poi.hssf.usermodel.HSSFSheet;
import  org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;


//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {
    public static final String delimiter = ",";

        public static void main (String[]args) throws IOException {
            String strCsvFile = "C:\\Users\\User\\Documents\\test.csv";
            String strExcelFile = "C:\\Users\\User\\Documents\\sample.xls";

            ReadAndWriteToExcel();
        }
  public static void ReadAndWriteToExcel() throws IOException
  {
      ArrayList<ArrayList<String>> allRowAndColData = null;
      ArrayList<String> oneRowData = null;
      String fName = "C:\\Users\\User\\Documents\\test.csv";
      String currentLine;
      FileInputStream fis = new FileInputStream(fName);
      DataInputStream myInput = new DataInputStream(fis);

      allRowAndColData = new ArrayList<ArrayList<String>>();
      while ((currentLine = myInput.readLine()) != null) {
          oneRowData = new ArrayList<String>();
          String oneRowArray[] = currentLine.split(",");
          for (int j = 0; j < oneRowArray.length; j++) {
              oneRowData.add(oneRowArray[j]);
          }
          allRowAndColData.add(oneRowData);
          System.out.println();

      }

      try {
          HSSFWorkbook workBook = new HSSFWorkbook();
          HSSFSheet sheet = workBook.createSheet("sheet1");
          for (int i = 0; i < allRowAndColData.size(); i++) {
              ArrayList<String> ardata = (ArrayList<String>) allRowAndColData.get(i);
              HSSFRow row = sheet.createRow(0 + i);
              for (int k = 0; k < ardata.size(); k++) {
                  System.out.print(ardata.get(k));
                  HSSFCell cell = row.createCell(k);
                  cell.setCellValue(ardata.get(k).toString());
              }
              System.out.println();
          }
          FileOutputStream fileOutputStream =  new FileOutputStream("C:\\Users\\User\\Documents\\sample.xls");
          workBook.write(fileOutputStream);
          fileOutputStream.close();
          System.out.println("File Exported Successfully in Excel file");
      } catch (Exception ex) {
      }
  }




    static void readDatafromFile (String csvFile)
        {
            BufferedReader br = null;
            try {
                br = new BufferedReader(new FileReader(csvFile));
            } catch (FileNotFoundException e) {
                throw new RuntimeException(e);
            }
            try {

                String line = " ";
                String[] tempArr;
                while ((line = br.readLine()) != null) {
                    tempArr = line.split(delimiter);
                    for (String tempStr : tempArr) {
                        System.out.print(tempStr + " ");
                    }
                    System.out.println();
                }
            } catch (IOException ioe) {
                ioe.printStackTrace();
            } finally {
                try {
                    br.close();
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
            }

        }
    }
