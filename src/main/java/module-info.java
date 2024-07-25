module timetablekrylov.timetablekrylovgr {
    requires javafx.controls;
    requires javafx.fxml;


    opens timetablekrylov.timetablekrylovgr to javafx.fxml;
    exports timetablekrylov.timetablekrylovgr;
}