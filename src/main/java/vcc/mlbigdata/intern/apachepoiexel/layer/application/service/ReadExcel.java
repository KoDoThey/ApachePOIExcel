package vcc.mlbigdata.intern.apachepoiexel.layer.application.service;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import vcc.mlbigdata.intern.apachepoiexel.layer.application.domain.model.Person;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ReadExcel {
    public static void main(String[] args) throws IOException {
        final String fileName = "/home/minh/Downloads/Telegram Desktop/6_Danh_sách_12_000_Thành_viên_Vàng_Mobifone_tại_Hà_Nội.xls";
        final List<Person> people = readExcel(fileName);
        for (Person person : people) {
            System.out.println(person);
        }
    }

    public static List<Person> readExcel(String fileName) throws IOException {
        List<Person> personList = new ArrayList<>();
        InputStream inputStream = new FileInputStream(fileName);

        Workbook workbook = new HSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        DataFormatter dataFormatter = new DataFormatter();

        Iterator<Row> iteratorRow = sheet.iterator();
        Row firstRow = sheet.getRow(5);
        Cell firstCell = firstRow.getCell(0);
//        System.out.println(firstCell.getStringCellValue().getClass());
        List<Person> listOfPerson = new ArrayList<Person>();

        while (iteratorRow.hasNext()) {
            Row currentRow = iteratorRow.next();

            Person person = new Person();
            person.setNumber(Double.parseDouble(dataFormatter.formatCellValue(currentRow.getCell(0))));
            person.setName(currentRow.getCell(1).getStringCellValue());
            person.setTelephone(currentRow.getCell(2).getStringCellValue());
            person.setAddress(currentRow.getCell(3).getStringCellValue());
            person.setPermanentAddress(currentRow.getCell(4).getStringCellValue());
            person.setTel(currentRow.getCell(5).getStringCellValue());
            person.setIdNumber(currentRow.getCell(6).getStringCellValue());
            listOfPerson.add(person);
        }
        for (Person person : listOfPerson) {
            System.out.println(person);
        }
        workbook.close();
        return personList;
    }
}



