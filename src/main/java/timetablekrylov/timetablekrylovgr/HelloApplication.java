package timetablekrylov.timetablekrylovgr;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.scene.text.Font;
import javafx.stage.Stage;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

public class HelloApplication extends Application {

    @Override
    public void start(Stage stage) throws IOException {

        // Creating a universal separator
        String separator = File.separator;

        // CSS style options
        String ButtonFontSize = "-fx-font-size: 15px";

        String NameButtonGroup = "Группы";

        String NameButtonTeacher = "Преподаватели";

        String NameButtonClassroom = "Аудитории";

        // Scene groups
        // --------------------------------------------------------------------------------------

        // I get an array (Checkbox) of the Faculty of FITR (from ReadGroupFITR)
        ArrayList<CheckBox> ArrayCheckBoxGrFITR = ReadGroupFITR(separator);
        ArrayCheckBoxGrFITR = CheckBoxStyleChanges(ArrayCheckBoxGrFITR);

        // I get an array (Checkbox) of the Faculty of SPO (from ReadGroupSPO)
        ArrayList<CheckBox> ArrayCheckBoxGrSPO = ReadGroupSPO(separator);
        ArrayCheckBoxGrSPO = CheckBoxStyleChanges(ArrayCheckBoxGrSPO);

        ArrayList<CheckBox> ArrayCheckBoxTeacher = ReadTeacher(separator);
        ArrayCheckBoxTeacher = CheckBoxStyleChanges(ArrayCheckBoxTeacher);

        ArrayList<CheckBox> ArrayCheckBoxClassroom = ReadClassroom(separator);
        ArrayCheckBoxClassroom = CheckBoxStyleChanges(ArrayCheckBoxClassroom);

        // Creating a (head) text
        Label TitleGroup = SetStyleTitle("Расписание групп:");

        // Creating a button to go to groups
        Button buttonSwapGroup = CreatorButtonSwap(NameButtonGroup,ButtonFontSize);

        // Creating a button to go to teachers
        Button buttonSwapTeacher = CreatorButtonSwap(NameButtonTeacher,ButtonFontSize);

        // Creating a button to go to the audience
        Button buttonSwapClassroom = CreatorButtonSwap(NameButtonClassroom,ButtonFontSize);

        // Creating horizontal layout
        HBox ButtonHBoxGroup = new HBox(10);
        ButtonHBoxGroup.setAlignment(Pos.CENTER);

        // Adding buttons to it
        ButtonHBoxGroup.getChildren().addAll(buttonSwapGroup,buttonSwapTeacher,buttonSwapClassroom);

        // Creating a drop-down list
        ChoiceBox<String> ChoiceBoxFaculty = SetStyleChoiceBox(ButtonFontSize);
        //Adding values to the ChoiceBox
        ChoiceBoxFaculty.getItems().addAll("Информационные технологии и радиоэлектроника (ФИТР)","Машиностроительный (МСФ)","Гуманитарный (ГФ)","Определение среднего профессионального образования (СПО)");

        // Adding the hbox layout to control the choicebox
        HBox ChoiceBoxControl = new HBox();
        ChoiceBoxControl.setAlignment(Pos.CENTER);
        ChoiceBoxControl.getChildren().add(ChoiceBoxFaculty);
        ChoiceBoxControl.setPadding(new Insets(-20,-20,-20,-20));

        // --------------------------------------------------------------------------------------

        GridPane GridPaneGroupFITR = CreatorGridPaneCheckBox(ArrayCheckBoxGrFITR);

        VBox VBoxGridPaneGroupFITR = new VBox();
        VBoxGridPaneGroupFITR.getChildren().addAll(GridPaneGroupFITR);

        GridPane GridPaneGroupSPO = CreatorGridPaneCheckBox(ArrayCheckBoxGrSPO);

        VBox VBoxGridPaneGroupSPO = new VBox();
        VBoxGridPaneGroupSPO.getChildren().add(GridPaneGroupSPO);

        Label VBoxGridPaneGroupMSFText = new Label("");

        VBox VBoxGridPaneGroupMSF = new VBox();
        VBoxGridPaneGroupMSF.getChildren().add(VBoxGridPaneGroupMSFText);

        Label VBoxGridPaneGroupGFText = new Label("");

        VBox VBoxGridPaneGroupGF = new VBox();
        VBoxGridPaneGroupGF.getChildren().add(VBoxGridPaneGroupGFText);

        // -----------------------------------------------------------------------------------------------------

        // Creating a scroll pane
        ScrollPane GroupScrollPane = CreatorScrollPane();

        // Creating a button for creating a schedule
        Button buttonCreatorTimeTableOne = CreatorButtonCreatorTimetable("Создать индивидуальное расписание",ButtonFontSize);

        // Creating a button for creating a schedule
        Button buttonCreatorTimeTableTwo = CreatorButtonCreatorTimetable("Создать групповое расписание",ButtonFontSize);

        // Adding the hbox layout to control the button(buttonCreatorTimeTable)
        VBox ButtonTimeTableControl = new VBox(10);
        ButtonTimeTableControl.setAlignment(Pos.CENTER);
        ButtonTimeTableControl.setPadding(new Insets(-60,-60,-60,-60));
        ButtonTimeTableControl.getChildren().addAll(buttonCreatorTimeTableOne,buttonCreatorTimeTableTwo);

        // Creating a vbox to add all the elements scene
        VBox VBoxGroup  = new VBox(50);
        VBoxGroup.setAlignment(Pos.BASELINE_CENTER);
        VBoxGroup.getChildren().addAll(TitleGroup,ButtonHBoxGroup,ChoiceBoxControl,GroupScrollPane,ButtonTimeTableControl);

        // Creating a group scene
        Scene sceneGroup = new Scene(VBoxGroup,1300,900);

        // Scene teacher
        // --------------------------------------------------------------------------------------

        Label TitleTeacher = SetStyleTitle("Расписание преподавателей:");

        // Creating a button to go to groups
        Button buttonSwapGroupTeacher = CreatorButtonSwap(NameButtonGroup,ButtonFontSize);

        // Creating a button to go to teachers
        Button buttonSwapTeacherTeacher = CreatorButtonSwap(NameButtonTeacher,ButtonFontSize);

        // Creating a button to go to the audience
        Button buttonSwapClassroomTeacher = CreatorButtonSwap(NameButtonClassroom,ButtonFontSize);

        HBox ButtonTeacherHbox = new HBox(10);
        ButtonTeacherHbox.setAlignment(Pos.CENTER);
        ButtonTeacherHbox.getChildren().addAll(buttonSwapGroupTeacher,buttonSwapTeacherTeacher,buttonSwapClassroomTeacher);

        ScrollPane TeacherScrollPane = CreatorScrollPane();

        GridPane TeacherGridPane = CreatorGridPaneCheckBox(ArrayCheckBoxTeacher);

        TeacherScrollPane.setContent(TeacherGridPane);

        Button buttonCreatorTimeTableTeacherOne = CreatorButtonCreatorTimetable("Создать индивидуальное расписание",ButtonFontSize);

        Button buttonCreatorTimeTableTeacherTwo = CreatorButtonCreatorTimetable("Создать групповое расписание",ButtonFontSize);

        VBox ButtonTimeTableControlTeacher = new VBox(10);
        ButtonTimeTableControlTeacher.setPadding(new Insets(-60,-60,-60,-60));
        ButtonTimeTableControlTeacher.setAlignment(Pos.CENTER);
        ButtonTimeTableControlTeacher.getChildren().addAll(buttonCreatorTimeTableTeacherOne,buttonCreatorTimeTableTeacherTwo);

        VBox VBoxTeacher = new VBox(50);
        VBoxTeacher.setAlignment(Pos.BASELINE_CENTER);
        VBoxTeacher.getChildren().addAll(TitleTeacher,ButtonTeacherHbox,TeacherScrollPane,ButtonTimeTableControlTeacher);

        Scene sceneTeacher = new Scene(VBoxTeacher,1300,900);


        // Scene classroom
        // ---------------------------------------------------------------------------------------

        Label TitleClassroom = SetStyleTitle("Расписание аудиторий:");

        // Creating a button to go to groups
        Button buttonSwapGroupClassroom = CreatorButtonSwap(NameButtonGroup,ButtonFontSize);

        // Creating a button to go to teachers
        Button buttonSwapTeacherClassroom = CreatorButtonSwap(NameButtonTeacher,ButtonFontSize);

        // Creating a button to go to the audience
        Button buttonSwapClassroomClassroom = CreatorButtonSwap(NameButtonClassroom,ButtonFontSize);

        HBox buttonHBoxClassroom = new HBox(10);
        buttonHBoxClassroom.setAlignment(Pos.CENTER);
        buttonHBoxClassroom.getChildren().addAll(buttonSwapGroupClassroom,buttonSwapTeacherClassroom,buttonSwapClassroomClassroom);

        ScrollPane ClassroomScrollPane = CreatorScrollPane();

        GridPane ClassroomGridPane = CreatorGridPaneCheckBox(ArrayCheckBoxClassroom);
        ClassroomScrollPane.setContent(ClassroomGridPane);

        Button buttonCreatorTimeTableClassroomOne = CreatorButtonCreatorTimetable("Создать индивидуальное расписание",ButtonFontSize);

        Button buttonCreatorTimeTableClassroomTwo = CreatorButtonCreatorTimetable("Создать групповое расписание",ButtonFontSize);

        VBox ButtonTimeTableControlClassroom = new VBox(10);
        ButtonTimeTableControlClassroom.setPadding(new Insets(-60,-60,-60,-60));
        ButtonTimeTableControlClassroom.setAlignment(Pos.CENTER);
        ButtonTimeTableControlClassroom.getChildren().addAll(buttonCreatorTimeTableClassroomOne,buttonCreatorTimeTableClassroomTwo);


        VBox VBoxClassroom = new VBox(50);
        VBoxClassroom.setAlignment(Pos.BASELINE_CENTER);
        VBoxClassroom.getChildren().addAll(TitleClassroom,buttonHBoxClassroom,ClassroomScrollPane,ButtonTimeTableControlClassroom);

        Scene sceneClassroom = new Scene(VBoxClassroom,1300,900);


        // Button
        // ---------------------------------------------------------------------------------------

        buttonSwapTeacher.setOnAction(Event -> {
            stage.setScene(sceneTeacher);
        });

        buttonSwapGroup.setOnAction(Event -> {
            stage.setScene(sceneGroup);
        });

        buttonSwapClassroom.setOnAction(Event -> {
            stage.setScene(sceneClassroom);
        });

        buttonSwapTeacherTeacher.setOnAction(Event -> {
            stage.setScene(sceneTeacher);
        });

        buttonSwapGroupTeacher.setOnAction(Event -> {
            stage.setScene(sceneGroup);
        });

        buttonSwapClassroomTeacher.setOnAction(Event -> {
            stage.setScene(sceneClassroom);
        });

        buttonSwapGroupClassroom.setOnAction(Event -> {
            stage.setScene(sceneGroup);
        });

        buttonSwapTeacherClassroom.setOnAction(Event -> {
            stage.setScene(sceneTeacher);
        });

        buttonSwapClassroomClassroom.setOnAction(Event -> {
            stage.setScene(sceneClassroom);
        });

        ChoiceBoxFaculty.getSelectionModel().selectedItemProperty().addListener((V,OldView,NewView) -> {
            if(NewView.equals("Информационные технологии и радиоэлектроника (ФИТР)")){
                GroupScrollPane.setContent(VBoxGridPaneGroupFITR);
            }
            if(NewView.equals("Определение среднего профессионального образования (СПО)")){
                GroupScrollPane.setContent(VBoxGridPaneGroupSPO);
            }
            if(NewView.equals("Машиностроительный (МСФ)")){
                GroupScrollPane.setContent(VBoxGridPaneGroupMSF);
            }
            if(NewView.equals("Гуманитарный (ГФ)")){
                GroupScrollPane.setContent(VBoxGridPaneGroupGFText);
            }
        });


        // Stage
        // ---------------------------------------------------------------------------------------

        stage.setScene(sceneGroup);
        stage.show();

    }

    public static void main(String[] args) {
        launch();
    }

    public ArrayList<CheckBox> ReadGroupFITR(String separator) throws FileNotFoundException {

        ArrayList<CheckBox> CheckBoxGroup = new ArrayList<>();

        String path = "D:" + separator + "Javal" + separator + "TimetableKrylovGR" + separator + "(FITR) Group.txt";

        File file = new File(path);

        Scanner scanner = new Scanner(file);

        while(scanner.hasNextLine()){
            CheckBoxGroup.add(new CheckBox(scanner.nextLine()));
        }

        return CheckBoxGroup;

    }

    public ArrayList<CheckBox> ReadGroupSPO(String separator) throws FileNotFoundException {

        ArrayList<CheckBox> ArrayCheckBoxGroup = new ArrayList<>();

        String patch = "D:" + separator + "Javal" + separator + "TimetableKrylovGR" + separator + "(SPO) Group.txt";

        File file = new File(patch);

        Scanner scanner = new Scanner(file);

        while(scanner.hasNextLine()){
            ArrayCheckBoxGroup.add(new CheckBox(scanner.nextLine()));
        }

        return ArrayCheckBoxGroup;

    }

    public ArrayList<CheckBox> ReadTeacher(String separator) throws FileNotFoundException {

        ArrayList<CheckBox> ArrayTeacher = new ArrayList<>();

        String path = "D:" + separator + "Javal" + separator + "TimetableKrylovGR" + separator + "Teacher.txt";

        File file = new File(path);

        Scanner scanner = new Scanner(file);

        while(scanner.hasNextLine()){
            ArrayTeacher.add(new CheckBox(scanner.nextLine()));
        }

        return ArrayTeacher;
    }

    public ArrayList<CheckBox> ReadClassroom(String separator) throws FileNotFoundException {

        ArrayList<CheckBox> ArrayClassroom = new ArrayList<>();

        String patch = "D:" + separator + "Javal" + separator + "TimetableKrylovGR" + separator + "Classroom.txt";

        File file = new File(patch);

        Scanner scanner = new Scanner(file);

        while(scanner.hasNextLine()){
            ArrayClassroom.add(new CheckBox(scanner.nextLine()));
        }

        return ArrayClassroom;
    }

    public Button CreatorButtonSwap(String nameButton, String ButtonStyle){

        Button button =  new Button(nameButton);

        button.setStyle(ButtonStyle);

        button.setMinWidth(150);
        button.setMinHeight(40);
        button.setMaxHeight(40);
        button.setMaxWidth(150);

        return button;
    }

    public Button CreatorButtonCreatorTimetable(String nameButton, String ButtonStyle){

        Button button = new Button(nameButton);
        button.setStyle(ButtonStyle);

        // Setting the size of the Button
        button.setMinWidth(300);
        button.setMinHeight(40);
        button.setMaxHeight(40);
        button.setMaxWidth(300);

        return button;
    }

    public ArrayList<CheckBox> CheckBoxStyleChanges(ArrayList<CheckBox> Array){

        for(int i = 0; i < Array.size(); i++){
            Array.get(i).setFont(new Font("Time New Roman",15));
        }

        return Array;
    }

    public Label SetStyleTitle(String NameLabel){

        Label label = new Label(NameLabel);
        label.setFont(new Font("Arial",40));
        return label;

    }

    public ChoiceBox<String> SetStyleChoiceBox(String ButtonStyle){

        ChoiceBox<String> choiceBox = new ChoiceBox<>();
        choiceBox.setStyle(ButtonStyle);

        // Setting the size of the ChoiceBox
        choiceBox.setMaxSize(468,40);
        choiceBox.setMinSize(468,40);

        return choiceBox;
    }

    public GridPane CreatorGridPaneCheckBox(ArrayList<CheckBox> Array){

        GridPane gridPane = new GridPane();
        gridPane.setPadding(new Insets(10,10,10,10));
        gridPane.setVgap(10);
        gridPane.setHgap(10);

        int countStrInColumnZero = 0;

        for(int i = 0; i < Array.size()/2; i++){
            GridPane.setConstraints(Array.get(i),0,countStrInColumnZero);
            countStrInColumnZero++;

        }

        int countStrInColumnOne = 0;

        for(int i = Array.size()/2; i < Array.size(); i++){
            GridPane.setConstraints(Array.get(i),1,countStrInColumnOne);
            countStrInColumnOne++;
        }

        for(int i = 0; i < Array.size(); i++){
            gridPane.getChildren().add(Array.get(i));
        }

        return gridPane;
    }

    public ScrollPane CreatorScrollPane(){

        ScrollPane scrollPane = new ScrollPane();

        // Setting the size of the ScrollPane
        scrollPane.setMinHeight(300);
        scrollPane.setMinWidth(450);
        scrollPane.setMaxSize(470,300);

        return scrollPane;
    }


}