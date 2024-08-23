package timetablekrylov.timetablekrylovgr;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.scene.text.Font;
import javafx.stage.Stage;


import java.io.*;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class HelloApplication extends Application {

    @Override
    public void start(Stage stage) throws IOException {

        // Creating a universal separator
        String separator = File.separator;

        // CSS style options
        String ButtonFontSize = "-fx-font-size: 15px";

        // Variables for storing names for buttons
        String NameButtonGroup = "Группы";
        String NameButtonTeacher = "Преподаватели";
        String NameButtonClassroom = "Аудитории";

        // Scene groups
        // --------------------------------------------------------------------------------------

        // ArraysCheckBox
        // ---------------------------------------------------------------------------------------

        Comparator<CheckBox> CheckBoxComparator = new ComparatorCheckBox();

        // I get an array (Checkbox) of the Faculty of FITR (from ReadGroupFITR)
        ArrayList<CheckBox> ArrayCheckBoxGrFITR = ReadGroupFITR(separator);
        // Setting the CheckBoxes style
        ArrayCheckBoxGrFITR = CheckBoxStyleChanges(ArrayCheckBoxGrFITR);
        // For the convenience of work, I pass the values to a new variable
        ArrayList<CheckBox> finalArrayCheckBoxGrFITR = ArrayCheckBoxGrFITR;

        Collections.sort(finalArrayCheckBoxGrFITR,CheckBoxComparator);

        // I get an array (Checkbox) of the Faculty of SPO (from ReadGroupSPO)
        ArrayList<CheckBox> ArrayCheckBoxGrSPO = ReadGroupSPO(separator);
        ArrayCheckBoxGrSPO = CheckBoxStyleChanges(ArrayCheckBoxGrSPO);
        ArrayList<CheckBox> finalArrayCheckBoxGrSPO = ArrayCheckBoxGrSPO;

        Collections.sort(finalArrayCheckBoxGrSPO,CheckBoxComparator);

        // I get an array (Checkbox) of the Faculty of GF (from ReadGroupGF)
        ArrayList<CheckBox> ArrayCheckBoxGF = ReadGroupGF(separator);
        ArrayCheckBoxGF = CheckBoxStyleChanges(ArrayCheckBoxGF);
        ArrayList<CheckBox> finalArrayCheckBoxGF = ArrayCheckBoxGF;

        Collections.sort(finalArrayCheckBoxGF,CheckBoxComparator);

        // I get an array (Checkbox) of the Faculty of MSF (from ReadGroupMSF)
        ArrayList<CheckBox> ArrayCheckBoxMSF = ReadGroupMSF(separator);
        ArrayCheckBoxMSF = CheckBoxStyleChanges(ArrayCheckBoxMSF);
        ArrayList<CheckBox> finalArrayCheckBoxMSF = ArrayCheckBoxMSF;

        Collections.sort(finalArrayCheckBoxMSF,CheckBoxComparator);

        // I get an array (Checkbox) of the Teachers (from ReadTeacher)
        ArrayList<CheckBox> ArrayCheckBoxTeacher = ReadTeacher(separator);
        ArrayCheckBoxTeacher = CheckBoxStyleChanges(ArrayCheckBoxTeacher);
        ArrayList<CheckBox> finalArrayCheckBoxTeacher = ArrayCheckBoxTeacher;

        Collections.sort(finalArrayCheckBoxTeacher,CheckBoxComparator);

        // I get an array (Checkbox) of the Classroom (from ReadClassroom)
        ArrayList<CheckBox> ArrayCheckBoxClassroom = ReadClassroom(separator);
        ArrayCheckBoxClassroom = CheckBoxStyleChanges(ArrayCheckBoxClassroom);
        ArrayList<CheckBox> finalArrayCheckBoxClassroom = ArrayCheckBoxClassroom;

        Collections.sort(finalArrayCheckBoxClassroom,CheckBoxComparator);

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
        ComboBox<String> ComboBoxBoxFaculty = SetStyleComboBox(ButtonFontSize,"Выберите факультет");
        //Adding values to the ChoiceBox
        ComboBoxBoxFaculty.getItems().addAll("Информационные технологии и радиоэлектроника (ФИТР)","Машиностроительный (МСФ)","Гуманитарный (ГФ)","Определение среднего профессионального образования (СПО)");

        // Adding the hbox layout to control the choicebox
        HBox ComboBoxControl = new HBox();
        ComboBoxControl.setAlignment(Pos.CENTER);
        ComboBoxControl.getChildren().add(ComboBoxBoxFaculty);
        ComboBoxControl.setPadding(new Insets(-20,-20,-20,-20));

        // --------------------------------------------------------------------------------------

        // I use a grid template to neatly place the elements in 2 columns
        GridPane GridPaneGroupFITR = CreatorGridPaneCheckBox(ArrayCheckBoxGrFITR);

        // I use the Vbox template to control the elements
        VBox VBoxGridPaneGroupFITR = new VBox();
        VBoxGridPaneGroupFITR.getChildren().addAll(GridPaneGroupFITR);

        // I use a grid template to neatly place the elements in 2 columns
        GridPane GridPaneGroupSPO = CreatorGridPaneCheckBox(ArrayCheckBoxGrSPO);

        // I use the Vbox template to control the elements
        VBox VBoxGridPaneGroupSPO = new VBox();
        VBoxGridPaneGroupSPO.getChildren().add(GridPaneGroupSPO);

        GridPane GridPaneGroupGF = CreatorGridPaneCheckBox(ArrayCheckBoxGF);

        VBox VBoxGridPaneGroupGF = new VBox();
        VBoxGridPaneGroupGF.getChildren().add(GridPaneGroupGF);

        GridPane GridPaneGroupMSF = CreatorGridPaneCheckBox(ArrayCheckBoxMSF);

        VBox VBoxGridPaneGroupMSF = new VBox();
        VBoxGridPaneGroupMSF.getChildren().add(GridPaneGroupMSF);

        // -----------------------------------------------------------------------------------------------------

        // Creating a scroll pane
        ScrollPane GroupScrollPane = CreatorScrollPane();

        // I am creating a ChoiceBox so that the user can set semester
        ComboBox<String> ComboBoxSemesterGroup = SetStyleComboBox(ButtonFontSize,"Выберите семестр");
        // 1 or 2
        ComboBoxSemesterGroup.getItems().add("1");
        ComboBoxSemesterGroup.getItems().add("2");

        // I am creating a text field for the user to set the year
        TextField TextFieldYearGroup = new TextField();
        TextFieldYearGroup.setStyle(ButtonFontSize);
        TextFieldYearGroup.setPromptText("Укажите год");
        TextFieldYearGroup.setMinSize(468,40);
        TextFieldYearGroup.setMaxSize(468,40);

        // Creating a button for creating a schedule
        Button buttonCreatorTimeTableOne = CreatorButtonCreatorTimetable("Создать индивидуальное расписание",ButtonFontSize);

        // Creating a button for creating a schedule
        Button buttonCreatorTimeTableTwo = CreatorButtonCreatorTimetable("Создать групповое расписание",ButtonFontSize);

        VBox BottomMenuTimeTableControl = CreatorBottomMenu(ComboBoxSemesterGroup,TextFieldYearGroup,buttonCreatorTimeTableOne,buttonCreatorTimeTableTwo);

        // Creating a vbox to add all the elements scene
        VBox VBoxGroup  = new VBox(50);
        VBoxGroup.setAlignment(Pos.BASELINE_CENTER);
        VBoxGroup.getChildren().addAll(TitleGroup,ButtonHBoxGroup,ComboBoxControl,GroupScrollPane,BottomMenuTimeTableControl);

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

        // I am creating an HBox to control the buttons
        HBox ButtonTeacherHbox = new HBox(10);
        ButtonTeacherHbox.setAlignment(Pos.CENTER);
        ButtonTeacherHbox.getChildren().addAll(buttonSwapGroupTeacher,buttonSwapTeacherTeacher,buttonSwapClassroomTeacher);

        // I am creating a ScrollPane so that the lists of teacher groups and classrooms can be scaled.
        ScrollPane TeacherScrollPane = CreatorScrollPane();

        // Creating a teacher grid
        GridPane TeacherGridPane = CreatorGridPaneCheckBox(ArrayCheckBoxTeacher);

        // I am passing the teachers' grid
        TeacherScrollPane.setContent(TeacherGridPane);

        // I am creating a ChoiceBox so that the user can set semester
        ComboBox<String> ComboBoxSemesterTeacher = SetStyleComboBox(ButtonFontSize,"Выберите семестр");
        ComboBoxSemesterTeacher.getItems().add("1");
        ComboBoxSemesterTeacher.getItems().add("2");

        // I am creating a text field for the user to set the year
        TextField TextFieldYearTeacher = new TextField();
        TextFieldYearTeacher.setStyle(ButtonFontSize);
        TextFieldYearTeacher.setPromptText("Укажите год");
        TextFieldYearTeacher.setMinSize(468,40);
        TextFieldYearTeacher.setMaxSize(468,40);

        CheckBox checkBoxDistantFalse = new CheckBox();
        checkBoxDistantFalse.setText("Не учитывать заочную форму");
        checkBoxDistantFalse.setFont(new Font("Arial",15));

        // Creating a button for creating a schedule
        Button buttonCreatorTimeTableTeacherOne = CreatorButtonCreatorTimetable("Создать индивидуальное расписание",ButtonFontSize);

        // Creating a button for creating a schedule
        Button buttonCreatorTimeTableTeacherTwo = CreatorButtonCreatorTimetable("Создать групповое расписание",ButtonFontSize);

        // I use the function to form the bottom menu (CreatorBottomMenu)
        VBox BottomMenuTimeTableControlTeacher = CreatorBottomMenuTeacher(ComboBoxSemesterTeacher,TextFieldYearTeacher,checkBoxDistantFalse,buttonCreatorTimeTableTeacherOne,buttonCreatorTimeTableTeacherTwo);

        // Creating the final template for the teachers' scene
        VBox VBoxTeacher = new VBox(50);
        VBoxTeacher.setAlignment(Pos.BASELINE_CENTER);
        VBoxTeacher.getChildren().addAll(TitleTeacher,ButtonTeacherHbox,TeacherScrollPane,BottomMenuTimeTableControlTeacher);

        // Creating a scene
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

        // I am creating an HBox to control the buttons
        HBox buttonHBoxClassroom = new HBox(10);
        buttonHBoxClassroom.setAlignment(Pos.CENTER);
        buttonHBoxClassroom.getChildren().addAll(buttonSwapGroupClassroom,buttonSwapTeacherClassroom,buttonSwapClassroomClassroom);

        // I am creating a ScrollPane so that the lists of teacher groups and classrooms can be scaled.
        ScrollPane ClassroomScrollPane = CreatorScrollPane();

        // Creating a classroom grid
        GridPane ClassroomGridPane = CreatorGridPaneCheckBox(ArrayCheckBoxClassroom);

        // I am passing the classroom grid
        ClassroomScrollPane.setContent(ClassroomGridPane);

        // Creating a button for creating a schedule
        Button buttonCreatorTimeTableClassroomOne = CreatorButtonCreatorTimetable("Создать индивидуальное расписание",ButtonFontSize);

        // Creating a button for creating a schedule
        Button buttonCreatorTimeTableClassroomTwo = CreatorButtonCreatorTimetable("Создать групповое расписание",ButtonFontSize);

        // I am creating a ChoiceBox so that the user can set semester
        ComboBox<String> ComboBoxSemesterClassroom = SetStyleComboBox(ButtonFontSize,"Выберите семестр");
        ComboBoxSemesterClassroom.getItems().add("1");
        ComboBoxSemesterClassroom.getItems().add("2");

        // I am creating a text field for the user to set the year
        TextField TextFieldYearClassroom = new TextField();
        TextFieldYearClassroom.setStyle(ButtonFontSize);
        TextFieldYearClassroom.setPromptText("Укажите год");
        TextFieldYearClassroom.setMinSize(468,40);
        TextFieldYearClassroom.setMaxSize(468,40);

        VBox BottomMenuClassroom = CreatorBottomMenu(ComboBoxSemesterClassroom,TextFieldYearClassroom,buttonCreatorTimeTableClassroomOne,buttonCreatorTimeTableClassroomTwo);

        // Creating the final template for the classroom scene
        VBox VBoxClassroom = new VBox(50);
        VBoxClassroom.setAlignment(Pos.BASELINE_CENTER);
        VBoxClassroom.getChildren().addAll(TitleClassroom,buttonHBoxClassroom,ClassroomScrollPane,BottomMenuClassroom);

        // Creating a scene
        Scene sceneClassroom = new Scene(VBoxClassroom,1300,900);


        // Button
        // ---------------------------------------------------------------------------------------

        // Scene change button
        buttonSwapTeacher.setOnAction(Event -> {
            stage.setScene(sceneTeacher);
        });

        // Scene change button
        buttonSwapGroup.setOnAction(Event -> {
            stage.setScene(sceneGroup);
        });

        // Scene change button
        buttonSwapClassroom.setOnAction(Event -> {
            stage.setScene(sceneClassroom);
        });

        // Scene change button
        buttonSwapTeacherTeacher.setOnAction(Event -> {
            stage.setScene(sceneTeacher);
        });

        // Scene change button
        buttonSwapGroupTeacher.setOnAction(Event -> {
            stage.setScene(sceneGroup);
        });

        // Scene change button
        buttonSwapClassroomTeacher.setOnAction(Event -> {
            stage.setScene(sceneClassroom);
        });

        // Scene change button
        buttonSwapGroupClassroom.setOnAction(Event -> {
            stage.setScene(sceneGroup);
        });

        // Scene change button
        buttonSwapTeacherClassroom.setOnAction(Event -> {
            stage.setScene(sceneTeacher);
        });

        // Scene change button
        buttonSwapClassroomClassroom.setOnAction(Event -> {
            stage.setScene(sceneClassroom);
        });

        // The button for getting the schedule of one group
        buttonCreatorTimeTableOne.setOnAction(Event -> {
            String Json = CreatorOneURL(finalArrayCheckBoxGrFITR,finalArrayCheckBoxGrSPO,finalArrayCheckBoxGF,finalArrayCheckBoxMSF,ComboBoxSemesterGroup,TextFieldYearGroup);
            Group group = new Group();
            group.CreatorCouples(Json);

            CreatorTableExelGroup creatorTableExel = new CreatorTableExelGroup();
            try {
                creatorTableExel.CreatorTimeTableOneGroup(group);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }

        });

        // A button for getting the schedule of several groups
        buttonCreatorTimeTableTwo.setOnAction(Event -> {
            ArrayList<String> ArrayURLAddressGroup = CreatorAllURL(finalArrayCheckBoxGrFITR,finalArrayCheckBoxGrSPO,finalArrayCheckBoxGF,finalArrayCheckBoxMSF,ComboBoxSemesterGroup,TextFieldYearGroup);
            ArrayList<Group> ArrayGroup = new ArrayList<>();

            for(int i = 0; i < ArrayURLAddressGroup.size(); i++){
                ArrayGroup.add(new Group());
            }

            for(int i = 0; i < ArrayURLAddressGroup.size(); i++){
                ArrayGroup.get(i).CreatorCouples(ArrayURLAddressGroup.get(i));
            }

            CreatorTableExelGroups creatorTableExelGroups = new CreatorTableExelGroups();
            try {
                creatorTableExelGroups.CreatorTimeTableGroups(ArrayGroup);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }

        });

        // A button for getting one teacher's schedule
        buttonCreatorTimeTableTeacherOne.setOnAction(Event -> {
            try {
                String Json = CreatorURLTeacherOne(finalArrayCheckBoxTeacher,ComboBoxSemesterTeacher,TextFieldYearTeacher);
                Teacher teacher = new Teacher(checkBoxDistantFalse.isSelected());
                teacher.CreatorCouples(Json);

                CreatorTableExelTeacher creatorTableExelTeacher = new CreatorTableExelTeacher();
                creatorTableExelTeacher.CreatorTimeTableTeacherOne(teacher);

            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        });

        // A button for getting the schedule of several teachers
        buttonCreatorTimeTableTeacherTwo.setOnAction(Event -> {
            try {
                ArrayList<String> ArrayJsonTeachers = CreatorURLTeacherAll(finalArrayCheckBoxTeacher,ComboBoxSemesterTeacher,TextFieldYearTeacher);
                ArrayList<Teacher> ArrayTeacher = new ArrayList<>();

                for(int i = 0; i < ArrayJsonTeachers.size(); i++){
                    ArrayTeacher.add(new Teacher(checkBoxDistantFalse.isSelected()));
                }

                for(int i = 0; i < ArrayTeacher.size(); i++){
                    ArrayTeacher.get(i).CreatorCouples(ArrayJsonTeachers.get(i));
                }

                CreatorTableExelTeachers creatorTableExelTeachers = new CreatorTableExelTeachers();
                creatorTableExelTeachers.CreateTimeTableTeachers(ArrayTeacher);

            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        });

        buttonCreatorTimeTableClassroomOne.setOnAction(Event -> {

            int countSelect = 0;

            for(int i = 0; i < finalArrayCheckBoxClassroom.size(); i++){
                if(finalArrayCheckBoxClassroom.get(i).isSelected()){
                    countSelect++;
                }
            }

            if(countSelect == 1) {
                try {
                    ArrayList<String> Array = CreatorCoupleAllGroup(finalArrayCheckBoxGrFITR, finalArrayCheckBoxGrSPO, finalArrayCheckBoxGF, finalArrayCheckBoxMSF, ComboBoxSemesterClassroom, TextFieldYearClassroom, finalArrayCheckBoxClassroom);
                    ArrayList<CoupleGroup> ArrayCouple = new ArrayList<>();

                    for (int i = 0; i < Array.size(); i++) {
                        ArrayCouple.add(new CoupleGroup());
                    }

                    for (int i = 0; i < ArrayCouple.size(); i++) {
                        ArrayCouple.get(i).CreatorCouple(Array.get(i));
                    }

                    CreatorTableExelClassroom creatorTableExelClassroom = new CreatorTableExelClassroom();
                    creatorTableExelClassroom.CreatorTimeTableClassroomOne(ArrayCouple);

                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
            }

        });

        buttonCreatorTimeTableClassroomTwo.setOnAction(Event -> {

            try {
                ArrayList<String> Array = CreatorCoupleAllGroup(finalArrayCheckBoxGrFITR,finalArrayCheckBoxGrSPO,finalArrayCheckBoxGF,finalArrayCheckBoxMSF,ComboBoxSemesterClassroom,TextFieldYearClassroom,finalArrayCheckBoxClassroom);
                ArrayList<CoupleGroup> ArrayCouple = new ArrayList<>();

                for(int i = 0; i < Array.size(); i++){
                    ArrayCouple.add(new CoupleGroup());
                }

                for(int i = 0; i < ArrayCouple.size(); i++){
                    ArrayCouple.get(i).CreatorCouple(Array.get(i));
                }

                CreatorTableExelClassrooms creatorTableExelClassrooms = new CreatorTableExelClassrooms();
                creatorTableExelClassrooms.CreateTableExelClassroom(ArrayCouple,finalArrayCheckBoxClassroom);

            } catch (IOException e) {
                throw new RuntimeException(e);
            }

        });

        // Scene Timetable Group
        // -----------------------------------------------------------------------------------------




        // Scene Timetable Teacher
        // -----------------------------------------------------------------------------------------



        // Scene Timetable Classroom
        // -----------------------------------------------------------------------------------------

        // Changing information in the ScrollPane depending on the selected ChoiceBox item
        ComboBoxBoxFaculty.getSelectionModel().selectedItemProperty().addListener((V,OldView,NewView) -> {
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
                GroupScrollPane.setContent(VBoxGridPaneGroupGF);
            }
        });


        // Stage
        // ---------------------------------------------------------------------------------------

        // Setting the default scene
        stage.setScene(sceneGroup);
        // Window (stage) demonstration
        stage.show();

    }

    public static void main(String[] args) {
        launch();
    }

    // Method for creating a CheckBox array for FITR faculty
    public ArrayList<CheckBox> ReadGroupFITR(String separator) throws FileNotFoundException {

        // Creating a new array
        ArrayList<CheckBox> CheckBoxGroup = new ArrayList<>();

        // I set the path to the file where our values are stored
        String patch = "D:" + separator + "Javal" + separator + "TimetableKrylovGR" + separator + "Group MIVLGU" + separator + "(FITR) Group.txt";

        // Creating an object of the file class
        File file = new File(patch);

        // Creating a scanner
        Scanner scanner = new Scanner(file);

        // While the scanner is reading the file, I create CheckBox objects in the array
        while(scanner.hasNextLine()){
            CheckBoxGroup.add(new CheckBox(scanner.nextLine()));
        }

        // I am returning the result
        return CheckBoxGroup;

    }

    // Method for creating a CheckBox array for SPO faculty
    public ArrayList<CheckBox> ReadGroupSPO(String separator) throws FileNotFoundException {

        // Creating a new array
        ArrayList<CheckBox> ArrayCheckBoxGroup = new ArrayList<>();

        // I set the path to the file where our values are stored
        String patch = "D:" + separator + "Javal" + separator + "TimetableKrylovGR" + separator + "Group MIVLGU" + separator + "(SPO) Group.txt";

        // Creating an object of the file class
        File file = new File(patch);

        // Creating a scanner
        Scanner scanner = new Scanner(file);

        // While the scanner is reading the file, I create CheckBox objects in the array
        while(scanner.hasNextLine()){
            ArrayCheckBoxGroup.add(new CheckBox(scanner.nextLine()));
        }

        // I am returning the result
        return ArrayCheckBoxGroup;

    }

    // Method for creating a CheckBox array for GF faculty
    public ArrayList<CheckBox> ReadGroupGF(String separator) throws FileNotFoundException {

        // Creating a new array
        ArrayList<CheckBox> ArrayCheckBoxGroup = new ArrayList<>();

        // I set the path to the file where our values are stored
        String patch = "D:" + separator + "Javal" + separator + "TimetableKrylovGR" + separator + "Group MIVLGU" + separator + "(GF) Group.txt";

        // Creating an object of the file class
        File file = new File(patch);

        // Creating a scanner
        Scanner scanner = new Scanner(file);

        // While the scanner is reading the file, I create CheckBox objects in the array
        while(scanner.hasNextLine()){
            ArrayCheckBoxGroup.add(new CheckBox(scanner.nextLine()));
        }

        // I am returning the result
        return ArrayCheckBoxGroup;
    }

    // Method for creating a CheckBox array for MSF faculty
    public ArrayList<CheckBox> ReadGroupMSF(String separator) throws FileNotFoundException {

        // Creating a new array
        ArrayList<CheckBox> ArrayCheckBoxGroup = new ArrayList<>();

        // I set the path to the file where our values are stored
        String patch = "D:" + separator + "Javal" + separator + "TimetableKrylovGR" + separator + "Group MIVLGU" + separator + "(MSF) Group.txt";

        // Creating an object of the file class
        File file = new File(patch);

        // Creating a scanner
        Scanner scanner = new Scanner(file);

        // While the scanner is reading the file, I create CheckBox objects in the array
        while(scanner.hasNextLine()){
            ArrayCheckBoxGroup.add(new CheckBox(scanner.nextLine()));
        }

        // I am returning the result
        return ArrayCheckBoxGroup;
    }

    // The method of creating an array of teacher CheckBox
    public ArrayList<CheckBox> ReadTeacher(String separator) throws FileNotFoundException {

        // Creating a new array
        ArrayList<CheckBox> ArrayTeacher = new ArrayList<>();

        // I set the path to the file where our values are stored
        String path = "D:" + separator + "Javal" + separator + "TimetableKrylovGR" + separator + "Teacher.txt";

        // Creating an object of the file class
        File file = new File(path);

        // Creating a scanner
        Scanner scanner = new Scanner(file);

        // While the scanner is reading the file, I create CheckBox objects in the array
        while(scanner.hasNextLine()){
            ArrayTeacher.add(new CheckBox(scanner.nextLine()));
        }

        // I am returning the result
        return ArrayTeacher;
    }

    // The method of creating an array of classroom CheckBox
    public ArrayList<CheckBox> ReadClassroom(String separator) throws FileNotFoundException {

        // Creating a new array
        ArrayList<CheckBox> ArrayClassroom = new ArrayList<>();

        // I set the path to the file where our values are stored
        String patch = "D:" + separator + "Javal" + separator + "TimetableKrylovGR" + separator + "Classroom.txt";

        // Creating an object of the file class
        File file = new File(patch);

        // Creating a scanner
        Scanner scanner = new Scanner(file);

        // While the scanner is reading the file, I create CheckBox objects in the array
        while(scanner.hasNextLine()){
            ArrayClassroom.add(new CheckBox(scanner.nextLine()));
        }

        // I am returning the result
        return ArrayClassroom;
    }

    // Method for creating buttons
    public Button CreatorButtonSwap(String nameButton, String ButtonStyle){

        Button button =  new Button(nameButton);
        button.setStyle(ButtonStyle);

        button.setMinWidth(150);
        button.setMinHeight(40);
        button.setMaxHeight(40);
        button.setMaxWidth(150);

        return button;
    }

    // Method for creating buttons
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

    // Method for changing the CheckBox style
    public ArrayList<CheckBox> CheckBoxStyleChanges(ArrayList<CheckBox> Array){

        for(int i = 0; i < Array.size(); i++){
            Array.get(i).setFont(new Font("Time New Roman",15));
        }

        return Array;
    }

    // Method for creating Text Title
    public Label SetStyleTitle(String NameLabel){

        Label label = new Label(NameLabel);
        label.setFont(new Font("Arial",40));
        return label;

    }

    // Method for creating ChoiceBox
    public ComboBox<String> SetStyleComboBox(String ButtonStyle,String Text){

        ComboBox<String> comboBox = new ComboBox<>();
        comboBox.setStyle(ButtonStyle);
        comboBox.setPromptText(Text);

        // Setting the size of the ChoiceBox
        comboBox.setMaxSize(468,40);
        comboBox.setMinSize(468,40);

        return comboBox;
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

    // Method for creating ScrollPane
    public ScrollPane CreatorScrollPane(){

        ScrollPane scrollPane = new ScrollPane();

        // Setting the size of the ScrollPane
        scrollPane.setMinHeight(300);
        scrollPane.setMinWidth(450);
        scrollPane.setMaxSize(470,300);

        return scrollPane;
    }

    // Creating a class for reading Json
    public String readAll(Reader reader) throws IOException {

        // Creating a StringBuilder object for convenient concatenation
        StringBuilder stringBuilder = new StringBuilder();

        // Variable for writing the index of the symbol
        int CheckStr;

        // Conditions for reading by character
        while((CheckStr = reader.read()) != -1){
            stringBuilder.append((char) CheckStr);
        }

        // Getting the result
        return stringBuilder.toString();
    }

    // I'm creating a function that will return Json to us
    public String ReadJsonInURL(String url) throws MalformedURLException, IOException {

        // Creating an InputStream that counts our URL
        InputStream inputStream = new URL(url).openConnection().getInputStream();

        try {

            // Creating a BufferedReader for writing data
            BufferedReader rd = new BufferedReader(new InputStreamReader(inputStream, StandardCharsets.UTF_8));

            // I'm writing our result in a string
            String jsonText = readAll(rd);

            // I am returning the result
            return jsonText;

        } finally {
            // Closing the reading stream
            inputStream.close();
        }

    }

    public String CreatorOneURL(ArrayList<CheckBox> ArrayFITR, ArrayList<CheckBox> ArraySPO, ArrayList<CheckBox> ArrayGF, ArrayList<CheckBox> ArrayMSF, ComboBox<String> Sem, TextField year) {

        for (int i = 0; i < ArrayFITR.size(); i++) {
            if (ArrayFITR.get(i).isSelected()) {

                String FirstURLData = "https://scala.mivlgu.ru/core/frontend/index.php?r=schedulecash/group&group=";
                String URLDataAndGroup = URLEncoder.encode(ArrayFITR.get(i).getText(), StandardCharsets.UTF_8);
                String Semester = "&semester=";
                String ChoiceBoxSem = Sem.getValue();
                String Year = "&year=";
                String TextFieldYear = year.getText();
                String Format = "&format=json";
                String FinalUrl = FirstURLData + URLDataAndGroup + Semester + ChoiceBoxSem + Year + TextFieldYear + Format;

                try {
                    String Json = ReadJsonInURL(FinalUrl);
                    return Json;
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }

            }
        }

        for (int i = 0; i < ArraySPO.size(); i++) {
            if (ArraySPO.get(i).isSelected()) {

                String FirstURLData = "https://scala.mivlgu.ru/core/frontend/index.php?r=schedulecash/group&group=";
                String URLDataAndGroup = URLEncoder.encode(ArraySPO.get(i).getText(), StandardCharsets.UTF_8);
                String Semester = "&semester=";
                String ChoiceBoxSem = Sem.getValue();
                String Year = "&year=";
                String TextFieldYear = year.getText();
                String Format = "&format=json";
                String FinalUrl = FirstURLData + URLDataAndGroup + Semester + ChoiceBoxSem + Year + TextFieldYear + Format;

                try {
                    String Json = ReadJsonInURL(FinalUrl);
                    return Json;
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
            }
        }

        for (int i = 0; i < ArrayGF.size(); i++) {
            if (ArrayGF.get(i).isSelected()) {

                String FirstURLData = "https://scala.mivlgu.ru/core/frontend/index.php?r=schedulecash/group&group=";
                String URLDataAndGroup = URLEncoder.encode(ArrayGF.get(i).getText(), StandardCharsets.UTF_8);
                String Semester = "&semester=";
                String ChoiceBoxSem = Sem.getValue();
                String Year = "&year=";
                String TextFieldYear = year.getText();
                String Format = "&format=json";
                String FinalUrl = FirstURLData + URLDataAndGroup + Semester + ChoiceBoxSem + Year + TextFieldYear + Format;

                try {
                    String Json = ReadJsonInURL(FinalUrl);
                    return Json;
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
            }
        }

        for (int i = 0; i < ArrayMSF.size(); i++) {
            if (ArrayMSF.get(i).isSelected()) {

                String FirstURLData = "https://scala.mivlgu.ru/core/frontend/index.php?r=schedulecash/group&group=";
                String URLDataAndGroup = URLEncoder.encode(ArrayMSF.get(i).getText(), StandardCharsets.UTF_8);
                String Semester = "&semester=";
                String ChoiceBoxSem = Sem.getValue();
                String Year = "&year=";
                String TextFieldYear = year.getText();
                String Format = "&format=json";
                String FinalUrl = FirstURLData + URLDataAndGroup + Semester + ChoiceBoxSem + Year + TextFieldYear + Format;

                try {
                    String Json = ReadJsonInURL(FinalUrl);
                    return Json;
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
            }
        }

        return "Сбор не удался!";
    }

    public ArrayList<String> CreatorAllURL(ArrayList<CheckBox> ArrayFITR, ArrayList<CheckBox> ArraySPO, ArrayList<CheckBox> ArrayGF, ArrayList<CheckBox> ArrayMSF, ComboBox<String> Sem, TextField year){

        ArrayList<String> ArrayURLAddress = new ArrayList<>();

        for(int i = 0; i < ArrayFITR.size(); i++){
            if(ArrayFITR.get(i).isSelected()){

                String FirstURLData = "https://scala.mivlgu.ru/core/frontend/index.php?r=schedulecash/group&group=";
                String URLDataAndGroup = URLEncoder.encode(ArrayFITR.get(i).getText(),StandardCharsets.UTF_8);
                String Semester = "&semester=";
                String ChoiceBoxSem = Sem.getValue();
                String Year = "&year=";
                String TextFieldYear = year.getText();
                String Format = "&format=json";
                String FinalUrl = FirstURLData+URLDataAndGroup+Semester+ChoiceBoxSem+Year+TextFieldYear+Format;

                try {
                    String Json = ReadJsonInURL(FinalUrl);
                    ArrayURLAddress.add(Json);
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }

            }
        }

        for (int i = 0; i < ArraySPO.size(); i++) {
            if (ArraySPO.get(i).isSelected()) {

                String FirstURLData = "https://scala.mivlgu.ru/core/frontend/index.php?r=schedulecash/group&group=";
                String URLDataAndGroup = URLEncoder.encode(ArraySPO.get(i).getText(), StandardCharsets.UTF_8);
                String Semester = "&semester=";
                String ChoiceBoxSem = Sem.getValue();
                String Year = "&year=";
                String TextFieldYear = year.getText();
                String Format = "&format=json";
                String FinalUrl = FirstURLData + URLDataAndGroup + Semester + ChoiceBoxSem + Year + TextFieldYear + Format;

                try {
                    String Json = ReadJsonInURL(FinalUrl);
                    ArrayURLAddress.add(Json);
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }

            }
        }

        for (int i = 0; i < ArrayGF.size(); i++) {
            if (ArrayGF.get(i).isSelected()) {

                String FirstURLData = "https://scala.mivlgu.ru/core/frontend/index.php?r=schedulecash/group&group=";
                String URLDataAndGroup = URLEncoder.encode(ArrayGF.get(i).getText(), StandardCharsets.UTF_8);
                String Semester = "&semester=";
                String ChoiceBoxSem = Sem.getValue();
                String Year = "&year=";
                String TextFieldYear = year.getText();
                String Format = "&format=json";
                String FinalUrl = FirstURLData + URLDataAndGroup + Semester + ChoiceBoxSem + Year + TextFieldYear + Format;

                try {
                    String Json = ReadJsonInURL(FinalUrl);
                    ArrayURLAddress.add(Json);
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }

            }
        }

        for (int i = 0; i < ArrayMSF.size(); i++) {
            if (ArrayMSF.get(i).isSelected()) {

                String FirstURLData = "https://scala.mivlgu.ru/core/frontend/index.php?r=schedulecash/group&group=";
                String URLDataAndGroup = URLEncoder.encode(ArrayMSF.get(i).getText(), StandardCharsets.UTF_8);
                String Semester = "&semester=";
                String ChoiceBoxSem = Sem.getValue();
                String Year = "&year=";
                String TextFieldYear = year.getText();
                String Format = "&format=json";
                String FinalUrl = FirstURLData + URLDataAndGroup + Semester + ChoiceBoxSem + Year + TextFieldYear + Format;

                try {
                    String Json = ReadJsonInURL(FinalUrl);
                    ArrayURLAddress.add(Json);
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }

            }
        }

        return ArrayURLAddress;

    }

    public ArrayList<String> CreatorCoupleAllGroup(ArrayList<CheckBox> ArrayFITR, ArrayList<CheckBox> ArraySPO, ArrayList<CheckBox> ArrayGF, ArrayList<CheckBox> ArrayMSF, ComboBox<String> Sem, TextField Year, ArrayList<CheckBox> ArrayClassroom) throws IOException {

        ArrayList<String> ArrayCoupleJson = new ArrayList<>();
        ArrayList<String> ArrayCoupleItogString = new ArrayList<>();

        ArrayList<String> ArrayJsonAllGroup = new ArrayList<>();
        String FirstURLData = "https://scala.mivlgu.ru/core/frontend/index.php?r=schedulecash/group&group=";
        String semester = "&semester=";
        String year = "&year=";
        String format = "&format=json";
        String ChoiceBoxSem = Sem.getValue();
        String TextFieldYear = Year.getText();

        for(int i = 0; i < ArrayFITR.size(); i++){

            String URLDataAndGroup = URLEncoder.encode(ArrayFITR.get(i).getText(), StandardCharsets.UTF_8);
            String FinalUrl = FirstURLData + URLDataAndGroup + semester + ChoiceBoxSem + year + TextFieldYear + format;

            String Json = ReadJsonInURL(FinalUrl);
            ArrayJsonAllGroup.add(Json);

        }

        for(int i = 0; i < ArraySPO.size(); i++){

            String URLDataAndGroup = URLEncoder.encode(ArraySPO.get(i).getText(), StandardCharsets.UTF_8);
            String FinalUrl = FirstURLData + URLDataAndGroup + semester + ChoiceBoxSem + year + TextFieldYear + format;

            String Json = ReadJsonInURL(FinalUrl);
            ArrayJsonAllGroup.add(Json);

        }

        for(int i = 0; i < ArrayGF.size(); i++){

            String URLDataAndGroup = URLEncoder.encode(ArrayGF.get(i).getText(), StandardCharsets.UTF_8);
            String FinalUrl = FirstURLData + URLDataAndGroup + semester + ChoiceBoxSem + year + TextFieldYear + format;

            String Json = ReadJsonInURL(FinalUrl);
            ArrayJsonAllGroup.add(Json);

        }

        for(int i = 0; i < ArrayMSF.size(); i++){

            String URLDataAndGroup = URLEncoder.encode(ArrayMSF.get(i).getText(), StandardCharsets.UTF_8);
            String FinalUrl = FirstURLData + URLDataAndGroup + semester + ChoiceBoxSem + year + TextFieldYear + format;

            String Json = ReadJsonInURL(FinalUrl);
            ArrayJsonAllGroup.add(Json);

        }

        for(int i = 0; i < ArrayJsonAllGroup.size(); i++){

            Pattern pattern = Pattern.compile("\"id_day\":\"\\d\",\"number_para\":\"\\d\",\"discipline\":\"[A-zА-я \\.\\-\\/]+\",\"type\":\"[А-яA-z]+\",\"type_week\":\"[А-яA-z]+\",\"aud\":\"[А-я\\. 0-9\\/\\-]+\",\"number_week\":\"[0-9\\/\\,\\-]+\",\"comment\":\"(|[A-zА-я\\.\\/\\,\\-])\",\"zaoch\":(true|false),\"name\":\"[A-zА-яё. ]+\"");
            Matcher matcher = pattern.matcher(ArrayJsonAllGroup.get(i));

            while(matcher.find()){
                ArrayCoupleJson.add(matcher.group());
            }

        }

        for(int i = 0; i < ArrayCoupleJson.size(); i++){
            Pattern pattern = Pattern.compile("\"aud\":\"((ауд|)[\\.0-9A-zА-я\\/ ]+|[A-zА-я])\"");
            Matcher matcher = pattern.matcher(ArrayCoupleJson.get(i));

            String RegXAudFull = "";

            while(matcher.find()){
                RegXAudFull = matcher.group();
            }

            Pattern pattern1 = Pattern.compile("[\\d]{1,3}(|\\-)([А-яA-z0-9]|)\\/\\d");
            Matcher matcher1 = pattern1.matcher(RegXAudFull);

            String Itog = "";

            while(matcher1.find()){
                Itog = matcher1.group();
            }

            for(int i2 = 0; i2 < ArrayClassroom.size(); i2++){

                if(ArrayClassroom.get(i2).isSelected()){
                    if(ArrayClassroom.get(i2).getText().equals(Itog)){
                       ArrayCoupleItogString.add(ArrayCoupleJson.get(i));
                    }
                }
            }
        }

        return ArrayCoupleItogString;
    }

    public String CreatorURLTeacherOne(ArrayList<CheckBox> ArrayTeacher,ComboBox<String> Sem,TextField year) throws IOException {

        for(int i = 0; i < ArrayTeacher.size(); i++){
            if(ArrayTeacher.get(i).isSelected()){

                String SemesterURL = "&semester=";
                String SemesterURLData = Sem.getValue();
                String YearURL = "&year=";
                String YearURLData = year.getText();
                String FormatURLData = "&format=json";

                String LastURLData = SemesterURL+SemesterURLData+YearURL+YearURLData+FormatURLData;

                String FirstURLData = "https://scala.mivlgu.ru/core/frontend/index.php?r=schedulecash/findteacher&fio=";
                String TeacherFio = URLEncoder.encode(ArrayTeacher.get(i).getText(),StandardCharsets.UTF_8);

                String FinalURL = FirstURLData+TeacherFio+LastURLData;

                String Json = ReadJsonInURL(FinalURL);

                Pattern pattern = Pattern.compile("\"[0-9]+\":\"[A-zА-яё]+ [A-zА-я\\ё]+ [A-zА-яё]+\"");
                Matcher matcher = pattern.matcher(Json);

                String RegXRez = "";

                while(matcher.find()){
                    RegXRez = matcher.group();
                }

                Pattern pattern1 = Pattern.compile("\\d+");
                Matcher matcher1 = pattern1.matcher(RegXRez);

                String TeacherId = "";

                while(matcher1.find()){
                    TeacherId = matcher1.group();
                }

                String FirstURL2Data = "https://scala.mivlgu.ru/core/frontend/index.php?r=schedulecash/teacher&teacher_id=";

                String FinalURL2 = FirstURL2Data+TeacherId+LastURLData;

                String TeacherJsonTimeTable = ReadJsonInURL(FinalURL2);
                return TeacherJsonTimeTable;
            }

        }

        return "Вы ничего не отметили!";
    }

    public ArrayList<String> CreatorURLTeacherAll(ArrayList<CheckBox> ArrayTeacher,ComboBox<String> Sem,TextField year) throws IOException {

        ArrayList<String> ArrayJsonTeacher = new ArrayList<>();

        for(int i = 0; i < ArrayTeacher.size(); i++){
            if(ArrayTeacher.get(i).isSelected()){

                String SemesterURL = "&semester=";
                String SemesterURLData = Sem.getValue();
                String YearURL = "&year=";
                String YearURLData = year.getText();
                String FormatURLData = "&format=json";

                String LastURLData = SemesterURL+SemesterURLData+YearURL+YearURLData+FormatURLData;

                String FirstURLData = "https://scala.mivlgu.ru/core/frontend/index.php?r=schedulecash/findteacher&fio=";
                String TeacherFio = URLEncoder.encode(ArrayTeacher.get(i).getText(),StandardCharsets.UTF_8);

                String FinalURL = FirstURLData+TeacherFio+LastURLData;

                String Json = ReadJsonInURL(FinalURL);

                Pattern pattern = Pattern.compile("\"[0-9]+\":\"[A-zА-яё]+ [A-zА-я\\ё]+ [A-zА-яё]+\"");
                Matcher matcher = pattern.matcher(Json);

                String RegXRez = "";

                while(matcher.find()){
                    RegXRez = matcher.group();
                }

                Pattern pattern1 = Pattern.compile("\\d+");
                Matcher matcher1 = pattern1.matcher(RegXRez);

                String TeacherId = "";

                while(matcher1.find()){
                    TeacherId = matcher1.group();
                }

                String FirstURL2Data = "https://scala.mivlgu.ru/core/frontend/index.php?r=schedulecash/teacher&teacher_id=";

                String FinalURL2 = FirstURL2Data+TeacherId+LastURLData;

                String TeacherJsonTimeTable = ReadJsonInURL(FinalURL2);
                ArrayJsonTeacher.add(TeacherJsonTimeTable);
            }

        }

        return ArrayJsonTeacher;

    }

    // Method for creating Bottom Menu
    public VBox CreatorBottomMenu(ComboBox<String> choiceBox,TextField textField,Button button1,Button button2){

        // Adding the hbox layout to control the button(buttonCreatorTimeTable)
        VBox BottomMenuTimeTableControl = new VBox(15);
        BottomMenuTimeTableControl.setAlignment(Pos.CENTER);
        BottomMenuTimeTableControl.setPadding(new Insets(-100,-100,-100,-100));
        BottomMenuTimeTableControl.getChildren().addAll(choiceBox,textField,button1,button2);

        return BottomMenuTimeTableControl;
    }

    // Method for creating Bottom Menu
    public VBox CreatorBottomMenuTeacher(ComboBox<String> choiceBox,TextField textField,CheckBox checkBoxDistantFalse,Button button1,Button button2){

        // Adding the hbox layout to control the button(buttonCreatorTimeTable)
        VBox BottomMenuTimeTableControl = new VBox(15);
        BottomMenuTimeTableControl.setAlignment(Pos.CENTER);
        BottomMenuTimeTableControl.setPadding(new Insets(-120,-120,-120,-120));
        BottomMenuTimeTableControl.getChildren().addAll(choiceBox,textField,checkBoxDistantFalse,button1,button2);

        return BottomMenuTimeTableControl;
    }


}