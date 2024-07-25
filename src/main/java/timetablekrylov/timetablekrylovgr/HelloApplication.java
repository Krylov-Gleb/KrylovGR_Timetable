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

        // Scene groups
        // --------------------------------------------------------------------------------------

        // I get an array (Checkbox) of the Faculty of FITR (from ReadGroupFITR)
        ArrayList<CheckBox> ArrayCheckBoxGrFITR = ReadGroupFITR(separator);

        // I sort through the objects and set the style
        for(int i = 0; i < ArrayCheckBoxGrFITR.size(); i++){
            ArrayCheckBoxGrFITR.get(i).setFont(new Font("Time New Roman",15));
        }

        // I get an array (Checkbox) of the Faculty of SPO (from ReadGroupSPO)
        ArrayList<CheckBox> ArrayCheckBoxGrSPO = ReadGroupSPO(separator);

        // I sort through the objects and set the style
        for(int i = 0; i < ArrayCheckBoxGrSPO.size(); i++){
            ArrayCheckBoxGrSPO.get(i).setFont(new Font("Time New Roman",15));
        }

        // Creating a (head) text
        Label TitleGroup = new Label("Расписание групп:");
        TitleGroup.setFont(new Font("Arial",40));

        // Creating a button to go to groups
        Button buttonSwapGroup = new Button("Группы");
        buttonSwapGroup.setStyle(ButtonFontSize);

        // Setting the size of the button
        buttonSwapGroup.setMinWidth(150);
        buttonSwapGroup.setMinHeight(40);
        buttonSwapGroup.setMaxHeight(40);
        buttonSwapGroup.setMaxWidth(150);

        // Creating a button to go to teachers
        Button buttonSwapTeacher = new Button("Преподаватели");
        buttonSwapTeacher.setStyle(ButtonFontSize);

        // Setting the size of the button
        buttonSwapTeacher.setMinHeight(40);
        buttonSwapTeacher.setMinWidth(150);
        buttonSwapTeacher.setMaxWidth(150);
        buttonSwapTeacher.setMaxHeight(40);

        // Creating a button to go to the audience
        Button buttonSwapClassroom = new Button("Аудитории");
        buttonSwapClassroom.setStyle(ButtonFontSize);

        // Setting the size of the button
        buttonSwapClassroom.setMinWidth(150);
        buttonSwapClassroom.setMinHeight(40);
        buttonSwapClassroom.setMaxHeight(40);
        buttonSwapClassroom.setMaxWidth(150);

        // Creating horizontal layout
        HBox ButtonHBoxGroup = new HBox(10);
        ButtonHBoxGroup.setAlignment(Pos.CENTER);

        // Adding buttons to it
        ButtonHBoxGroup.getChildren().addAll(buttonSwapGroup,buttonSwapTeacher,buttonSwapClassroom);

        // Creating a drop-down list
        ChoiceBox<String> ChoiceBoxFaculty = new ChoiceBox<>();
        ChoiceBoxFaculty.setStyle(ButtonFontSize);

        //Adding values to the ChoiceBox
        ChoiceBoxFaculty.getItems().addAll("Информационные технологии и радиоэлектроника (ФИТР)","Машиностроительный (МСФ)","Гуманитарный (ГФ)","Определение среднего профессионального образования (СПО)");

        // Setting the size of the ChoiceBox
        ChoiceBoxFaculty.setMaxSize(468,40);
        ChoiceBoxFaculty.setMinSize(468,40);

        // Adding the hbox layout to control the choicebox
        HBox ChoiceBoxControl = new HBox();
        ChoiceBoxControl.setAlignment(Pos.CENTER);
        ChoiceBoxControl.getChildren().add(ChoiceBoxFaculty);
        ChoiceBoxControl.setPadding(new Insets(-20,-20,-20,-20));

        // --------------------------------------------------------------------------------------

        GridPane GridPaneGroupFITR = new GridPane();
        GridPaneGroupFITR.setPadding(new Insets(10,10,10,10));
        GridPaneGroupFITR.setVgap(10);
        GridPaneGroupFITR.setHgap(10);

        int countStrFITRColumnZero = 0;

        for(int i = 0; i < ArrayCheckBoxGrFITR.size()/2; i++){
            GridPane.setConstraints(ArrayCheckBoxGrFITR.get(i),0,countStrFITRColumnZero);
            countStrFITRColumnZero++;

        }

        int countStrFITRColumnOne = 0;

        for(int i = ArrayCheckBoxGrFITR.size()/2; i < ArrayCheckBoxGrFITR.size(); i++){
            GridPane.setConstraints(ArrayCheckBoxGrFITR.get(i),1,countStrFITRColumnOne);
            countStrFITRColumnOne++;
        }

        for(int i = 0; i < ArrayCheckBoxGrFITR.size(); i++){
            GridPaneGroupFITR.getChildren().add(ArrayCheckBoxGrFITR.get(i));
        }

        VBox VBoxGridPaneGroupFITR = new VBox();
        VBoxGridPaneGroupFITR.getChildren().addAll(GridPaneGroupFITR);

        GridPane GridPaneGroupSPO = new GridPane();
        GridPaneGroupSPO.setPadding(new Insets(10,10,10,10));
        GridPaneGroupSPO.setHgap(10);
        GridPaneGroupSPO.setVgap(10);

        int countStrSPOColumnZero = 0;

        for(int i = 0; i < ArrayCheckBoxGrSPO.size()/2; i++){
            GridPane.setConstraints(ArrayCheckBoxGrSPO.get(i),0,countStrSPOColumnZero);
            countStrSPOColumnZero++;
        }

        int countStrSPOColumnOne = 0;

        for(int i = ArrayCheckBoxGrSPO.size()/2; i < ArrayCheckBoxGrSPO.size(); i++){
            GridPane.setConstraints(ArrayCheckBoxGrSPO.get(i),1,countStrSPOColumnOne);
            countStrSPOColumnOne++;
        }

        for(int i = 0; i < ArrayCheckBoxGrSPO.size(); i++){
            GridPaneGroupSPO.getChildren().add(ArrayCheckBoxGrSPO.get(i));
        }

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
        ScrollPane GroupScrollPane = new ScrollPane();

        // Setting the size of the ScrollPane
        GroupScrollPane.setMinHeight(300);
        GroupScrollPane.setMinWidth(400);
        GroupScrollPane.setMaxSize(400,300);

        // Creating a button for creating a schedule
        Button buttonCreatorTimeTable = new Button("Создать расписание");
        buttonCreatorTimeTable.setStyle(ButtonFontSize);

        // Setting the size of the Button
        buttonCreatorTimeTable.setMinWidth(200);
        buttonCreatorTimeTable.setMinHeight(40);
        buttonCreatorTimeTable.setMaxHeight(40);
        buttonCreatorTimeTable.setMaxWidth(200);

        // Adding the hbox layout to control the button(buttonCreatorTimeTable)
        HBox ButtonTimeTableControl = new HBox();
        ButtonTimeTableControl.setAlignment(Pos.CENTER);
        ButtonTimeTableControl.setPadding(new Insets(-20,-20,-20,-20));
        ButtonTimeTableControl.getChildren().add(buttonCreatorTimeTable);

        // Creating a vbox to add all the elements scene
        VBox VBoxGroup  = new VBox(50);
        VBoxGroup.setAlignment(Pos.BASELINE_CENTER);
        VBoxGroup.getChildren().addAll(TitleGroup,ButtonHBoxGroup,ChoiceBoxControl,GroupScrollPane,ButtonTimeTableControl);

        // Creating a group scene
        Scene sceneGroup = new Scene(VBoxGroup,1300,900);

        // Scene teacher
        // --------------------------------------------------------------------------------------

        Pane paneTeacher = new Pane();

        Scene sceneTeacher = new Scene(paneTeacher,1300,900);


        // Scene classroom
        // ---------------------------------------------------------------------------------------

        Pane paneClassroom = new Pane();

        Scene sceneClassroom = new Scene(paneClassroom,1300,900);


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
}