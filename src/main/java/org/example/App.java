package org.example;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.*;
import java.nio.file.Path;
import java.util.Arrays;

public class App
{
    public static void main( String[] args ) throws Exception
    {
        //исходный файл и ридер в кодировке Windows-1251
        File folder = new File("d:/dev/2022.8/excel/out/");
        File in = new File( folder.getAbsolutePath()+"/file.txt");
        in.getParentFile().mkdirs();
        if (!in.exists()){
        System.out.println("Отсутствует файл (табуляция разделитель) "+in.getAbsolutePath());
        System.exit(0);
        }
        BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(in), "Windows-1251"));


        //конечьный файл и райтер
        File out = new File("d:/dev/2022.8/excel/out/file.xls");
        out.getParentFile().mkdirs();
        FileOutputStream writer = new FileOutputStream(out);

        //сознаем эксель-книгу
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("txt-->excel");
        int rowid = 0;
        int cellid = 0;
        Cell cell;
        Row row;
        row = sheet.createRow(rowid++);

        //читаем строку из файла
        String line = reader.readLine();


        while (line != null) {

            System.out.println(line);
            //преобразуем строку в массив, разделитель табуляция
            String [] tmpstr =  StringUtils.split(line, "\t");
            //копируем массив в массив - это просто для обучения
            String [] str3 = Arrays.copyOf(tmpstr, tmpstr.length) ;
            //переводим массив в строку и выводим на экран
            System.out.println(Arrays.toString(tmpstr));

            for (String str : tmpstr){
                cell = row.createCell(cellid++);
                cell.setCellValue(str);
            }

            line = reader.readLine();
            if (line != null) {
                row = sheet.createRow(rowid++);
                cellid = 0;
            }
        }

        reader.close();
        workbook.write(writer);
        writer.close();

        System.out.println("Результат программы "+out.getAbsolutePath() );
    }
}
