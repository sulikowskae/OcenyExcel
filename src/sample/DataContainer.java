package sample;

import java.util.List;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;


public class DataContainer {

    protected ObservableList<Student> students;

    public ObservableList<Student> getStudents() {
        return students;
    }

    public void setStudents(List<Student> students) {
        this.students = FXCollections.observableArrayList(students);
    }

    public DataContainer() {
        students = FXCollections.observableArrayList();
    }
}