package timetablekrylov.timetablekrylovgr;

import javafx.scene.control.CheckBox;

import java.util.Comparator;

public class ComparatorCheckBox implements Comparator<CheckBox> {

    @Override
    public int compare(CheckBox checkBox1, CheckBox checkBox2) {
        return checkBox1.getText().toLowerCase().compareTo(checkBox2.getText().toLowerCase());
    }
}
