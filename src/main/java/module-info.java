module timetablekrylov.timetablekrylovgr {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.jsoup;
    requires jdk.jsobject;
    requires java.desktop;
    requires org.apache.poi.poi;


    opens timetablekrylov.timetablekrylovgr to javafx.fxml;
    exports timetablekrylov.timetablekrylovgr;
}