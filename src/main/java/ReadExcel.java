import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadExcel {
    File file;
    FileInputStream fileInputStream;
    XSSFWorkbook xssfWorkbook;
    XSSFSheet xssfSheet;
    public void readExcel(String path, String sheet)
    {
        file = new File(path);
        Row row  = null;

        try{
            fileInputStream=new FileInputStream(file);
            xssfWorkbook = new XSSFWorkbook(fileInputStream);
            xssfSheet = xssfWorkbook.getSheet(sheet);
            int primeraFila = xssfSheet.getFirstRowNum();
            int ultimaFila = xssfSheet.getLastRowNum();
            row = xssfSheet.getRow(primeraFila);
            int primeraColumna = row.getFirstCellNum();
            int ultimaColumna = row.getLastCellNum();
            for(int i=primeraFila;i<=ultimaFila;i++)
            {
                row = xssfSheet.getRow(i);
                for(int j = primeraColumna; j<ultimaColumna;j++)
                {
                    System.out.println(row.getCell(j).toString());
                }
            }



        }catch (FileNotFoundException e)
        {
            e.printStackTrace();
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }

    }


    @Test
    public void testRead()
    {
        readExcel("/Users/sady.cabrera/Downloads/dataExcel.xlsx","usuarios");
    }


}
