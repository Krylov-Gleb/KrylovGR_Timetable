module timetablekrylov.timetablekrylovgr {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.jsoup;
    requires jdk.jsobject;
    requires java.desktop;


    opens timetablekrylov.timetablekrylovgr to javafx.fxml;
    exports timetablekrylov.timetablekrylovgr;
}