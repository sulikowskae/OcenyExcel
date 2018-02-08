package sample;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.FillPatternType;


import javafx.event.ActionEvent;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.control.cell.PropertyValueFactory;


public class DaneOsobowe implements HierarchicalController<MainController> {

    public TextField imie;
    public TextField nazwisko;
    public TextField pesel;
    public TextField indeks;
    public TableView<Student> tabelka;
    private MainController parentController;

    public void dodaj(ActionEvent actionEvent) {
        Student st = new Student();
        st.setName(imie.getText());
        st.setSurname(nazwisko.getText());
        st.setPesel(pesel.getText());
        st.setIdx(indeks.getText());
        tabelka.getItems().add(st);
    }

    public void setParentController(MainController parentController) {
        this.parentController = parentController;
        tabelka.getItems().addAll(parentController.getDataContainer().getStudents());
        tabelka.setItems(parentController.getDataContainer().getStudents());
    }

    public void usunZmiany() {
        tabelka.getItems().clear();
        tabelka.getItems().addAll(parentController.getDataContainer().getStudents());
    }

    public MainController getParentController() {
        return parentController;
    }

    public void initialize() {
        for (TableColumn<Student, ?> studentTableColumn : tabelka.getColumns()) {
            if ("imie".equals(studentTableColumn.getId())) {
                studentTableColumn.setCellValueFactory(new PropertyValueFactory<>("name"));
            } else if ("nazwisko".equals(studentTableColumn.getId())) {
                studentTableColumn.setCellValueFactory(new PropertyValueFactory<>("surname"));
            } else if ("pesel".equals(studentTableColumn.getId())) {
                studentTableColumn.setCellValueFactory(new PropertyValueFactory<>("pesel"));
            } else if ("indeks".equals(studentTableColumn.getId())) {
                studentTableColumn.setCellValueFactory(new PropertyValueFactory<>("idx"));
            }
        }

    }

    public void zapisz(ActionEvent actionEvent) {

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("Studenci");

        String[] headers = {"Imię", "Nazwisko", "Ocena", "Uzasadnienie", "Numer indeksu", "PESEL"};
        int len = headers.length;

        XSSFRow row0 = sheet.createRow(0);
        XSSFCellStyle styleHeader = wb.createCellStyle();
        XSSFFont fontHeader = wb.createFont();

        fontHeader.setBold(true);
        styleHeader.setFont(fontHeader);

        for (int i = 0; i < len; i++) {
            Cell cell = row0.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(styleHeader);
        }
        int row = 1;
        for (Student student : tabelka.getItems()) {

            XSSFRow r = sheet.createRow(row);
            XSSFCellStyle style = wb.createCellStyle();
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            if (student.getGrade() == null) {
                style.setFillForegroundColor(IndexedColors.RED1.getIndex());
            } else if(student.getGrade() < 3){
                style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            } else if(student.getGrade() >= 3){
                style.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
            }

                Cell cell0 = r.createCell(0);
                cell0.setCellValue(student.getName());
                cell0.setCellStyle(style);

                Cell cell1 = r.createCell(1);
                cell1.setCellValue(student.getSurname());
                cell1.setCellStyle(style);

                Cell cell2 = r.createCell(2);
                if (student.getGrade() != null) {
                    cell2.setCellValue(student.getGrade());
                }
                cell2.setCellStyle(style);

                Cell cell3 = r.createCell(3);
                cell3.setCellValue(student.getGradeDetailed());
                cell3.setCellStyle(style);

                Cell cell4 = r.createCell(4);
                cell4.setCellValue(student.getIdx());
                cell4.setCellStyle(style);

                Cell cell5 = r.createCell(5);
                cell5.setCellValue(student.getPesel());
                cell5.setCellStyle(style);

                row++;

        }

        try (FileOutputStream fos = new FileOutputStream("C:/Users/Admin/Desktop/data.xlsx")) {
            wb.write(fos);
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /** Uwaga na serializację: https://sekurak.pl/java-vs-deserializacja-niezaufanych-danych-i-zdalne-wykonanie-kodu-czesc-i/ */
    public void wczytaj(ActionEvent actionEvent) {
        ArrayList<Student> studentsList = new ArrayList<>();
        ArrayList<String> headersList = new ArrayList<>();

        try (FileInputStream ois = new FileInputStream("C:/Users/Admin/Desktop/data.xlsx")) {
            XSSFWorkbook wb = new XSSFWorkbook(ois);
            XSSFSheet sheet = wb.getSheet("Studenci");

            String[] headers = {"Imię", "Nazwisko", "Ocena", "Uzasadnienie", "Numer indeksu", "PESEL"};
            int len = headers.length;

            XSSFRow row0 = sheet.createRow(0);
            XSSFCellStyle styleHeader = wb.createCellStyle();
            XSSFFont fontHeader = wb.createFont();

            fontHeader.setBold(true);
            styleHeader.setFont(fontHeader);

            for (int i = 0; i < len; i++) {
                Cell cell = row0.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(styleHeader);
            }


            for (int i = 1; i <= sheet.getLastRowNum(); i++) {

                XSSFRow row = sheet.getRow(i);
                Student student = new Student();
                student.setName(row.getCell(0).getStringCellValue());
                student.setSurname(row.getCell(1).getStringCellValue());
                if (row.getCell(2) != null) {
                    student.setGrade(row.getCell(2).getNumericCellValue());
                }
                student.setGradeDetailed(row.getCell(3).getStringCellValue());
                student.setIdx(row.getCell(4).getStringCellValue());
                student.setPesel(row.getCell(5).getStringCellValue());
                studentsList.add(student);
            }
            tabelka.getItems().clear();
            tabelka.getItems().addAll(studentsList);
            ois.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}