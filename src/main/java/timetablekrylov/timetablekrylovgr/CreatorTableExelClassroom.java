package timetablekrylov.timetablekrylovgr;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.Date;

public class CreatorTableExelClassroom {

    private Workbook workbookOneClassroom = new HSSFWorkbook();

    private Sheet Classroom = workbookOneClassroom.createSheet("Аудитория");

    private Row rowZero = Classroom.createRow(0);
    private Cell cellDayWeekAndNumberCouple = rowZero.createCell(0);

    private Cell Monday = rowZero.createCell(1);
    private Cell Tuesday = rowZero.createCell(2);
    private Cell Wednesday = rowZero.createCell(3);
    private Cell Thursday = rowZero.createCell(4);
    private Cell Friday = rowZero.createCell(5);
    private Cell Saturday = rowZero.createCell(6);

    private Row StrOneCouple = Classroom.createRow(1);
    private Row StrTwoCouple = Classroom.createRow(2);
    private Row StrThreeCouple = Classroom.createRow(3);
    private Row StrFourCouple = Classroom.createRow(4);
    private Row StrFiveCouple = Classroom.createRow(5);
    private Row StrSixCouple = Classroom.createRow(6);
    private Row StrSevenCouple = Classroom.createRow(7);

    private Cell OneCouple = StrOneCouple.createCell(0);
    private Cell TwoCouple = StrTwoCouple.createCell(0);
    private Cell ThreeCouple = StrThreeCouple.createCell(0);
    private Cell FourCouple = StrFourCouple.createCell(0);
    private Cell FiveCouple = StrFiveCouple.createCell(0);
    private Cell SixCouple = StrSixCouple.createCell(0);
    private Cell SevenCouple = StrSevenCouple.createCell(0);

    public ArrayList<String> ConcatCoupleDayWeek(ArrayList<CoupleGroup> CoupleOne,ArrayList<CoupleGroup> CoupleTwo,ArrayList<CoupleGroup> CoupleThree,ArrayList<CoupleGroup> CoupleFour, ArrayList<CoupleGroup> CoupleFive,ArrayList<CoupleGroup> CoupleSix,ArrayList<CoupleGroup> CoupleSeven){

        String CoupleOneDayWeekOne = "";
        String CoupleTwoDayWeekTwo = "";
        String CoupleThreeDayWeekThree = "";
        String CoupleFourDayWeekFour = "";
        String CoupleFiveDayWeekFive = "";
        String CoupleSixDayWeekSix = "";
        String CoupleSevenDayWeekSeven = "";

        for(int i = 0; i < CoupleOne.size(); i++){
            if(i < CoupleOne.size()-1) {
                CoupleOneDayWeekOne = CoupleOneDayWeekOne + CoupleOne.get(i).GetDiscipline() + " (" + CoupleOne.get(i).GetCoupleType() + ")\n" + "недели. " + CoupleOne.get(i).GetNumberWeek() + " " + CoupleOne.get(i).GetTeacherName() + " " + CoupleOne.get(i).GetAud() + " подгруппы. " + CoupleOne.get(i).GetUnderGroup() + "\n" + "\n";
            }
            else{
                CoupleOneDayWeekOne = CoupleOneDayWeekOne + CoupleOne.get(i).GetDiscipline() + " (" + CoupleOne.get(i).GetCoupleType() + ")\n" + "недели. " + CoupleOne.get(i).GetNumberWeek() + " " + CoupleOne.get(i).GetTeacherName() + " " + CoupleOne.get(i).GetAud() + " подгруппы. " + CoupleOne.get(i).GetUnderGroup();
            }
        }

        for(int i = 0; i < CoupleTwo.size(); i++){
            if(i < CoupleTwo.size()-1) {
                CoupleTwoDayWeekTwo = CoupleTwoDayWeekTwo + CoupleTwo.get(i).GetDiscipline() + " (" + CoupleTwo.get(i).GetCoupleType() + ")\n" + "недели. " + CoupleTwo.get(i).GetNumberWeek() + " " + CoupleTwo.get(i).GetTeacherName() + " " + CoupleTwo.get(i).GetAud() + " подгруппы. " + CoupleTwo.get(i).GetUnderGroup() + "\n" + "\n";
            }
            else{
                CoupleTwoDayWeekTwo = CoupleTwoDayWeekTwo + CoupleTwo.get(i).GetDiscipline() + " (" + CoupleTwo.get(i).GetCoupleType() + ")\n" + "недели. " + CoupleTwo.get(i).GetNumberWeek() + " " + CoupleTwo.get(i).GetTeacherName() + " " + CoupleTwo.get(i).GetAud() + " подгруппы. " + CoupleTwo.get(i).GetUnderGroup();
            }
        }

        for(int i = 0; i < CoupleThree.size(); i++){
            if(i < CoupleThree.size()-1) {
                CoupleThreeDayWeekThree = CoupleThreeDayWeekThree + CoupleThree.get(i).GetDiscipline() + " (" + CoupleThree.get(i).GetCoupleType() + ")\n" + "недели. " + CoupleThree.get(i).GetNumberWeek() + " " + CoupleThree.get(i).GetTeacherName() + " " + CoupleThree.get(i).GetAud() + " подгруппы. " + CoupleThree.get(i).GetUnderGroup() + "\n" + "\n";
            }
            else{
                CoupleThreeDayWeekThree = CoupleThreeDayWeekThree + CoupleThree.get(i).GetDiscipline() + " (" + CoupleThree.get(i).GetCoupleType() + ")\n" + "недели. " + CoupleThree.get(i).GetNumberWeek() + " " + CoupleThree.get(i).GetTeacherName() + " " + CoupleThree.get(i).GetAud() + " подгруппы. " + CoupleThree.get(i).GetUnderGroup();
            }
        }

        for(int i = 0; i < CoupleFour.size(); i++){
            if(i < CoupleFour.size()-1) {
                CoupleFourDayWeekFour = CoupleFourDayWeekFour + CoupleFour.get(i).GetDiscipline() + " (" + CoupleFour.get(i).GetCoupleType() + ")\n" + "недели. " + CoupleFour.get(i).GetNumberWeek() + " " + CoupleFour.get(i).GetTeacherName() + " " + CoupleFour.get(i).GetAud() + " подгруппы. " + CoupleFour.get(i).GetUnderGroup() + "\n" + "\n";
            }
            else{
                CoupleFourDayWeekFour = CoupleFourDayWeekFour + CoupleFour.get(i).GetDiscipline() + " (" + CoupleFour.get(i).GetCoupleType() + ")\n" + "недели. " + CoupleFour.get(i).GetNumberWeek() + " " + CoupleFour.get(i).GetTeacherName() + " " + CoupleFour.get(i).GetAud() + " подгруппы. " + CoupleFour.get(i).GetUnderGroup();
            }
        }

        for(int i = 0; i < CoupleFive.size(); i++){
            if(i < CoupleFive.size()-1) {
                CoupleFiveDayWeekFive = CoupleFiveDayWeekFive + CoupleFive.get(i).GetDiscipline() + " (" + CoupleFive.get(i).GetCoupleType() + ")\n" + "недели. " + CoupleFive.get(i).GetNumberWeek() + " " + CoupleFive.get(i).GetTeacherName() + " " + CoupleFive.get(i).GetAud() + " подгруппы. " + CoupleFive.get(i).GetUnderGroup() + "\n" + "\n";
            }
            else{
                CoupleFiveDayWeekFive = CoupleFiveDayWeekFive + CoupleFive.get(i).GetDiscipline() + " (" + CoupleFive.get(i).GetCoupleType() + ")\n" + "недели. " + CoupleFive.get(i).GetNumberWeek() + " " + CoupleFive.get(i).GetTeacherName() + " " + CoupleFive.get(i).GetAud() + " подгруппы. " + CoupleFive.get(i).GetUnderGroup();
            }
        }

        for(int i = 0; i < CoupleSix.size(); i++){
            if(i < CoupleSix.size()-1) {
                CoupleSixDayWeekSix = CoupleSixDayWeekSix + CoupleSix.get(i).GetDiscipline() + " (" + CoupleSix.get(i).GetCoupleType() + ")\n" + "недели. " + CoupleSix.get(i).GetNumberWeek() + " " + CoupleSix.get(i).GetTeacherName() + " " + CoupleSix.get(i).GetAud() + " подгруппы. " + CoupleSix.get(i).GetUnderGroup() + "\n" + "\n";
            }
            else{
                CoupleSixDayWeekSix = CoupleSixDayWeekSix + CoupleSix.get(i).GetDiscipline() + " (" + CoupleSix.get(i).GetCoupleType() + ")\n" + "недели. " + CoupleSix.get(i).GetNumberWeek() + " " + CoupleSix.get(i).GetTeacherName() + " " + CoupleSix.get(i).GetAud() + " подгруппы. " + CoupleSix.get(i).GetUnderGroup();
            }
        }

        for(int i = 0; i < CoupleSeven.size(); i++){
            if(i < CoupleSeven.size()-1) {
                CoupleSevenDayWeekSeven = CoupleSevenDayWeekSeven + CoupleSeven.get(i).GetDiscipline() + " (" + CoupleSeven.get(i).GetCoupleType() + ")\n" + "недели. " + CoupleSeven.get(i).GetNumberWeek() + " " + CoupleSeven.get(i).GetTeacherName() + " " + CoupleSeven.get(i).GetAud() + " подгруппы. " + CoupleSeven.get(i).GetUnderGroup() + "\n" + "\n";
            }
            else{
                CoupleSevenDayWeekSeven = CoupleSevenDayWeekSeven + CoupleSeven.get(i).GetDiscipline() + " (" + CoupleSeven.get(i).GetCoupleType() + ")\n" + "недели. " + CoupleSeven.get(i).GetNumberWeek() + " " + CoupleSeven.get(i).GetTeacherName() + " " + CoupleSeven.get(i).GetAud() + " подгруппы. " + CoupleSeven.get(i).GetUnderGroup();
            }
        }

        ArrayList<String> CoupleDayWeek = new ArrayList<>();
        CoupleDayWeek.add(CoupleOneDayWeekOne);
        CoupleDayWeek.add(CoupleTwoDayWeekTwo);
        CoupleDayWeek.add(CoupleThreeDayWeekThree);
        CoupleDayWeek.add(CoupleFourDayWeekFour);
        CoupleDayWeek.add(CoupleFiveDayWeekFive);
        CoupleDayWeek.add(CoupleSixDayWeekSix);
        CoupleDayWeek.add(CoupleSevenDayWeekSeven);

        return CoupleDayWeek;
    }

    public void CreatorTimeTableClassroomOne(ArrayList<CoupleGroup> Array) throws IOException {

        int HeightPoints = 50;
        int ColumnWidth = 12000;

        ArrayList<CoupleGroup> CoupleOne = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleTwo = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleThree = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleFour = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleFive = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleSix = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleSeven = new ArrayList<>();

        String Aud = Array.get(0).GetAud();

        Pattern pattern = Pattern.compile("\\d{3}");
        Matcher matcher = pattern.matcher(Aud);

        String RezAud = "";

        while(matcher.find()){
            RezAud = matcher.group();
        }

        workbookOneClassroom.setSheetName(0,RezAud);

        CellStyle cellStyle = workbookOneClassroom.createCellStyle();
        cellStyle.setWrapText(true);

        ArrayList<CoupleGroup> ArrayMonday = new ArrayList<>();
        ArrayList<CoupleGroup> ArrayTuesday = new ArrayList<>();
        ArrayList<CoupleGroup> ArrayWednesday = new ArrayList<>();
        ArrayList<CoupleGroup> ArrayThursday = new ArrayList<>();
        ArrayList<CoupleGroup> ArrayFriday = new ArrayList<>();
        ArrayList<CoupleGroup> ArraySaturday = new ArrayList<>();

        cellDayWeekAndNumberCouple.setCellStyle(cellStyle);
        cellDayWeekAndNumberCouple.setCellValue("День недели" + "\n" + "Номер пары");
        Classroom.setColumnWidth(0,8000);
        rowZero.setHeightInPoints(30);

        OneCouple.setCellValue("1 пара");
        TwoCouple.setCellValue("2 пара");
        ThreeCouple.setCellValue("3 пара");
        FourCouple.setCellValue("4 пара");
        FiveCouple.setCellValue("5 пара");
        SixCouple.setCellValue("6 пара");
        SevenCouple.setCellValue("7 пара");

        Monday.setCellValue("Понедельник");
        Classroom.setColumnWidth(1,8000);

        Tuesday.setCellValue("Вторник");
        Classroom.setColumnWidth(2,8000);

        Wednesday.setCellValue("Среде");
        Classroom.setColumnWidth(3,8000);

        Thursday.setCellValue("Четверг");
        Classroom.setColumnWidth(4,8000);

        Friday.setCellValue("Пятница");
        Classroom.setColumnWidth(5,8000);

        Saturday.setCellValue("Суббота");
        Classroom.setColumnWidth(6,8000);

        for(int i = 0; i < Array.size(); i++){

            int IdDay = Array.get(i).GetIDDay();

            switch (IdDay){
                case (1):{
                    ArrayMonday.add(Array.get(i));
                    break;
                }
                case (2):{
                    ArrayTuesday.add(Array.get(i));
                    break;
                }
                case (3):{
                    ArrayWednesday.add(Array.get(i));
                    break;
                }
                case (4):{
                    ArrayThursday.add(Array.get(i));
                    break;
                }
                case (5):{
                    ArrayFriday.add(Array.get(i));
                    break;
                }
                case (6):{
                    ArraySaturday.add(Array.get(i));
                    break;
                }
            }
        }

        // Couple Monday

        String CoupleOneMonday = "";
        String CoupleTwoMonday = "";
        String CoupleThreeMonday = "";
        String CoupleFourMonday = "";
        String CoupleFiveMonday = "";
        String CoupleSixMonday = "";
        String CoupleSevenMonday = "";

        for(int i = 0; i < ArrayMonday.size(); i++){

            switch (ArrayMonday.get(i).GetCoupleNumber()) {
                case (1): {
                    CoupleOne.add(ArrayMonday.get(i));
                    break;
                }
                case (2): {
                    CoupleTwo.add(ArrayMonday.get(i));
                    break;
                }
                case (3): {
                    CoupleThree.add(ArrayMonday.get(i));
                    break;
                }
                case (4): {
                    CoupleFour.add(ArrayMonday.get(i));
                    break;
                }
                case (5): {
                    CoupleFive.add(ArrayMonday.get(i));
                    break;
                }
                case (6): {
                    CoupleSix.add(ArrayMonday.get(i));
                    break;
                }
                case (7): {
                    CoupleSeven.add(ArrayMonday.get(i));
                    break;
                }
            }
        }

        ArrayList<String> ArrayStrMonday = ConcatCoupleDayWeek(CoupleOne,CoupleTwo,CoupleThree,CoupleFour,CoupleFive,CoupleSix,CoupleSeven);

        CoupleOneMonday = ArrayStrMonday.get(0);
        CoupleTwoMonday = ArrayStrMonday.get(1);
        CoupleThreeMonday = ArrayStrMonday.get(2);
        CoupleFourMonday = ArrayStrMonday.get(3);
        CoupleFiveMonday = ArrayStrMonday.get(4);
        CoupleSixMonday = ArrayStrMonday.get(5);
        CoupleSevenMonday = ArrayStrMonday.get(6);

        Cell cellOneCoupleMonday = StrOneCouple.createCell(1);
        cellOneCoupleMonday.setCellValue(CoupleOneMonday);
        cellOneCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(1,ColumnWidth);
        StrOneCouple.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleMonday = StrTwoCouple.createCell(1);
        cellTwoCoupleMonday.setCellValue(CoupleTwoMonday);
        cellTwoCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(1,ColumnWidth);
        StrTwoCouple.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleMonday = StrThreeCouple.createCell(1);
        cellThreeCoupleMonday.setCellValue(CoupleThreeMonday);
        cellThreeCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(1,ColumnWidth);
        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleMonday = StrFourCouple.createCell(1);
        cellFourCoupleMonday.setCellValue(CoupleFourMonday);
        cellFourCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(1,ColumnWidth);
        StrFourCouple.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleMonday = StrFiveCouple.createCell(1);
        cellFiveCoupleMonday.setCellValue(CoupleFiveMonday);
        cellFiveCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(1,ColumnWidth);
        StrFiveCouple.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleMonday = StrSixCouple.createCell(1);
        cellSixCoupleMonday.setCellValue(CoupleSixMonday);
        cellSixCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(1,ColumnWidth);
        StrSixCouple.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleMonday = StrSevenCouple.createCell(1);
        cellSevenCoupleMonday.setCellValue(CoupleSevenMonday);
        cellSevenCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(1,ColumnWidth);
        StrSevenCouple.setHeightInPoints(HeightPoints);

        CoupleOne.clear();
        CoupleTwo.clear();
        CoupleThree.clear();
        CoupleFour.clear();
        CoupleFive.clear();
        CoupleSix.clear();
        CoupleSeven.clear();

//        HeightOneCouple = HeightPoints;
//        HeightTwoCouple = HeightPoints;
//        HeightThreeCouple = HeightPoints;
//        HeightFourCouple = HeightPoints;
//        HeightFiveCouple = HeightPoints;
//        HeightSixCouple = HeightPoints;
//        HeightSevenCouple = HeightPoints;

        // Couple Tuesday

        String CoupleOneTuesday = "";
        String CoupleTwoTuesday = "";
        String CoupleThreeTuesday = "";
        String CoupleFourTuesday = "";
        String CoupleFiveTuesday = "";
        String CoupleSixTuesday = "";
        String CoupleSevenTuesday = "";

        for(int i = 0; i < ArrayTuesday.size(); i++){

            switch (ArrayTuesday.get(i).GetCoupleNumber()) {
                case (1): {
                    CoupleOne.add(ArrayTuesday.get(i));
                    break;
                }
                case (2): {
                    CoupleTwo.add(ArrayTuesday.get(i));
                    break;
                }
                case (3): {
                    CoupleThree.add(ArrayTuesday.get(i));
                    break;
                }
                case (4): {
                    CoupleFour.add(ArrayTuesday.get(i));
                    break;
                }
                case (5): {
                    CoupleFive.add(ArrayTuesday.get(i));
                    break;
                }
                case (6): {
                    CoupleSix.add(ArrayTuesday.get(i));
                    break;
                }
                case (7): {
                    CoupleSeven.add(ArrayTuesday.get(i));
                    break;
                }
            }
        }

        ArrayList<String> ArrayStrTuesday = ConcatCoupleDayWeek(CoupleOne,CoupleTwo,CoupleThree,CoupleFour,CoupleFive,CoupleSix,CoupleSeven);

        CoupleOneTuesday = ArrayStrTuesday.get(0);
        CoupleTwoTuesday = ArrayStrTuesday.get(1);
        CoupleThreeTuesday = ArrayStrTuesday.get(2);
        CoupleFourTuesday = ArrayStrTuesday.get(3);
        CoupleFiveTuesday = ArrayStrTuesday.get(4);
        CoupleSixTuesday = ArrayStrTuesday.get(5);
        CoupleSevenTuesday = ArrayStrTuesday.get(6);

        Cell cellOneCoupleTuesday = StrOneCouple.createCell(2);
        cellOneCoupleTuesday.setCellValue(CoupleOneTuesday);
        cellOneCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(2,ColumnWidth);

//        if(MaxHeightOneCouple < HeightOneCouple){
//            MaxHeightOneCouple = HeightOneCouple;
//        }

        StrOneCouple.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleTuesday = StrTwoCouple.createCell(2);
        cellTwoCoupleTuesday.setCellValue(CoupleTwoTuesday);
        cellTwoCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(2,ColumnWidth);

//        if(MaxHeightTwoCouple < HeightTwoCouple){
//            MaxHeightTwoCouple = HeightTwoCouple;
//        }

        StrTwoCouple.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleTuesday = StrThreeCouple.createCell(2);
        cellThreeCoupleTuesday.setCellValue(CoupleThreeTuesday);
        cellThreeCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(2,ColumnWidth);

//        if(MaxHeightThreeCouple < HeightThreeCouple){
//            MaxHeightThreeCouple = HeightThreeCouple;
//        }

        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleTuesday = StrFourCouple.createCell(2);
        cellFourCoupleTuesday.setCellValue(CoupleFourTuesday);
        cellFourCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(2,ColumnWidth);

//        if(MaxHeightFourCouple < HeightFourCouple){
//            MaxHeightFourCouple = HeightFourCouple;
//        }

        StrFourCouple.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleTuesday = StrFiveCouple.createCell(2);
        cellFiveCoupleTuesday.setCellValue(CoupleFiveTuesday);
        cellFiveCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(2,ColumnWidth);

//        if(MaxHeightFiveCouple < HeightFiveCouple){
//            MaxHeightFiveCouple = HeightFiveCouple;
//        }

        StrFiveCouple.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleTuesday = StrSixCouple.createCell(2);
        cellSixCoupleTuesday.setCellValue(CoupleSixTuesday);
        cellSixCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(2,ColumnWidth);

//        if(MaxHeightSixCouple < HeightSixCouple){
//            MaxHeightSixCouple = HeightSixCouple;
//        }

        StrSixCouple.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleTuesday = StrSevenCouple.createCell(2);
        cellSevenCoupleTuesday.setCellValue(CoupleSevenTuesday);
        cellSevenCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(2,ColumnWidth);

//        if(MaxHeightSevenCouple < HeightSevenCouple){
//            MaxHeightSevenCouple = HeightSevenCouple;
//        }

        StrSevenCouple.setHeightInPoints(HeightPoints);

        CoupleOne.clear();
        CoupleTwo.clear();
        CoupleThree.clear();
        CoupleFour.clear();
        CoupleFive.clear();
        CoupleSix.clear();
        CoupleSeven.clear();

//        HeightOneCouple = HeightPoints;
//        HeightTwoCouple = HeightPoints;
//        HeightThreeCouple = HeightPoints;
//        HeightFourCouple = HeightPoints;
//        HeightFiveCouple = HeightPoints;
//        HeightSixCouple = HeightPoints;
//        HeightSevenCouple = HeightPoints;

        // Couple Wednesday

        String CoupleOneWednesday = "";
        String CoupleTwoWednesday = "";
        String CoupleThreeWednesday = "";
        String CoupleFourWednesday = "";
        String CoupleFiveWednesday = "";
        String CoupleSixWednesday = "";
        String CoupleSevenWednesday = "";

        for(int i = 0; i < ArrayWednesday.size(); i++){

            switch (ArrayWednesday.get(i).GetCoupleNumber()) {
                case (1): {
                    CoupleOne.add(ArrayWednesday.get(i));
                    break;
                }
                case (2): {
                    CoupleTwo.add(ArrayWednesday.get(i));
                    break;
                }
                case (3): {
                    CoupleThree.add(ArrayWednesday.get(i));
                    break;
                }
                case (4): {
                    CoupleFour.add(ArrayWednesday.get(i));
                    break;
                }
                case (5): {
                    CoupleFive.add(ArrayWednesday.get(i));
                    break;
                }
                case (6): {
                    CoupleSix.add(ArrayWednesday.get(i));
                    break;
                }
                case (7): {
                    CoupleSeven.add(ArrayWednesday.get(i));
                    break;
                }
            }
        }

        ArrayList<String> ArrayStrWednesday = ConcatCoupleDayWeek(CoupleOne,CoupleTwo,CoupleThree,CoupleFour,CoupleFive,CoupleSix,CoupleSeven);

        CoupleOneWednesday = ArrayStrWednesday.get(0);
        CoupleTwoWednesday = ArrayStrWednesday.get(1);
        CoupleThreeWednesday = ArrayStrWednesday.get(2);
        CoupleFourWednesday = ArrayStrWednesday.get(3);
        CoupleFiveWednesday = ArrayStrWednesday.get(4);
        CoupleSixWednesday = ArrayStrWednesday.get(5);
        CoupleSevenWednesday = ArrayStrWednesday.get(6);

        Cell cellOneCoupleWednesday = StrOneCouple.createCell(3);
        cellOneCoupleWednesday.setCellValue(CoupleOneWednesday);
        cellOneCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(3,ColumnWidth);

//        if(MaxHeightOneCouple < HeightOneCouple){
//            MaxHeightOneCouple = HeightOneCouple;
//        }

        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleWednesday = StrTwoCouple.createCell(3);
        cellTwoCoupleWednesday.setCellValue(CoupleTwoWednesday);
        cellTwoCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(3,ColumnWidth);

//        if(MaxHeightTwoCouple < HeightTwoCouple){
//            MaxHeightTwoCouple = HeightTwoCouple;
//        }

        StrTwoCouple.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleWednesday = StrThreeCouple.createCell(3);
        cellThreeCoupleWednesday.setCellValue(CoupleThreeWednesday);
        cellThreeCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(3,ColumnWidth);

//        if(MaxHeightThreeCouple < HeightThreeCouple){
//            MaxHeightThreeCouple = HeightThreeCouple;
//        }

        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleWednesday = StrFourCouple.createCell(3);
        cellFourCoupleWednesday.setCellValue(CoupleFourWednesday);
        cellFourCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(3,ColumnWidth);

//        if(MaxHeightFourCouple < HeightFourCouple){
//            MaxHeightFourCouple = HeightFourCouple;
//        }

        StrFourCouple.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleWednesday = StrFiveCouple.createCell(3);
        cellFiveCoupleWednesday.setCellValue(CoupleFiveWednesday);
        cellFiveCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(3,ColumnWidth);

//        if(MaxHeightFiveCouple < HeightFiveCouple){
//            MaxHeightFiveCouple = HeightFiveCouple;
//        }

        StrFiveCouple.setHeightInPoints(HeightPoints);


        Cell cellSixCoupleWednesday = StrSixCouple.createCell(3);
        cellSixCoupleWednesday.setCellValue(CoupleSixWednesday);
        cellSixCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(3,ColumnWidth);

//        if(MaxHeightSixCouple < HeightSixCouple){
//            MaxHeightSixCouple = HeightSixCouple;
//        }

        StrSixCouple.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleWednesday = StrSevenCouple.createCell(3);
        cellSevenCoupleWednesday.setCellValue(CoupleSevenWednesday);
        cellSevenCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(3,ColumnWidth);

//        if(MaxHeightSevenCouple < HeightSevenCouple){
//            MaxHeightSevenCouple = HeightSevenCouple;
//        }

        StrSevenCouple.setHeightInPoints(HeightPoints);

        CoupleOne.clear();
        CoupleTwo.clear();
        CoupleThree.clear();
        CoupleFour.clear();
        CoupleFive.clear();
        CoupleSix.clear();
        CoupleSeven.clear();

//        HeightOneCouple = HeightPoints;
//        HeightTwoCouple = HeightPoints;
//        HeightThreeCouple = HeightPoints;
//        HeightFourCouple = HeightPoints;
//        HeightFiveCouple = HeightPoints;
//        HeightSixCouple = HeightPoints;
//        HeightSevenCouple = HeightPoints;

        // Couple Thursday

        String CoupleOneThursday = "";
        String CoupleTwoThursday = "";
        String CoupleThreeThursday = "";
        String CoupleFourThursday= "";
        String CoupleFiveThursday = "";
        String CoupleSixThursday = "";
        String CoupleSevenThursday = "";

        for(int i = 0; i < ArrayThursday.size(); i++){

            switch (ArrayThursday.get(i).GetCoupleNumber()) {
                case (1): {
                    CoupleOne.add(ArrayThursday.get(i));
                    break;
                }
                case (2): {
                    CoupleTwo.add(ArrayThursday.get(i));
                    break;
                }
                case (3): {
                    CoupleThree.add(ArrayThursday.get(i));
                    break;
                }
                case (4): {
                    CoupleFour.add(ArrayThursday.get(i));
                    break;
                }
                case (5): {
                    CoupleFive.add(ArrayThursday.get(i));
                    break;
                }
                case (6): {
                    CoupleSix.add(ArrayThursday.get(i));
                    break;
                }
                case (7): {
                    CoupleSeven.add(ArrayThursday.get(i));
                    break;
                }
            }
        }

        ArrayList<String> ArrayStrThursday = ConcatCoupleDayWeek(CoupleOne,CoupleTwo,CoupleThree,CoupleFour,CoupleFive,CoupleSix,CoupleSeven);

        CoupleOneThursday = ArrayStrThursday.get(0);
        CoupleTwoThursday = ArrayStrThursday.get(1);
        CoupleThreeThursday = ArrayStrThursday.get(2);
        CoupleFourThursday = ArrayStrThursday.get(3);
        CoupleFiveThursday = ArrayStrThursday.get(4);
        CoupleSixThursday = ArrayStrThursday.get(5);
        CoupleSevenThursday = ArrayStrThursday.get(6);

        Cell cellOneCoupleThursday = StrOneCouple.createCell(4);
        cellOneCoupleThursday.setCellValue(CoupleOneThursday);
        cellOneCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(4,ColumnWidth);

//        if(MaxHeightOneCouple < HeightOneCouple){
//            MaxHeightOneCouple = HeightOneCouple;
//        }

        StrOneCouple.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleThursday = StrTwoCouple.createCell(4);
        cellTwoCoupleThursday.setCellValue(CoupleTwoThursday);
        cellTwoCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(4,ColumnWidth);

//        if(MaxHeightTwoCouple < HeightTwoCouple){
//            MaxHeightTwoCouple = HeightTwoCouple;
//        }

        StrTwoCouple.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleThursday = StrThreeCouple.createCell(4);
        cellThreeCoupleThursday.setCellValue(CoupleThreeThursday);
        cellThreeCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(4,ColumnWidth);

//        if(MaxHeightThreeCouple < HeightThreeCouple){
//            MaxHeightThreeCouple = HeightThreeCouple;
//        }

        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleThursday = StrFourCouple.createCell(4);
        cellFourCoupleThursday.setCellValue(CoupleFourThursday);
        cellFourCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(4,ColumnWidth);

//        if(MaxHeightFourCouple < HeightFourCouple){
//            MaxHeightFourCouple = HeightFourCouple;
//        }

        StrFourCouple.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleThursday = StrFiveCouple.createCell(4);
        cellFiveCoupleThursday.setCellValue(CoupleFiveThursday);
        cellFiveCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(4,ColumnWidth);

//        if(MaxHeightFiveCouple < HeightFiveCouple){
//            MaxHeightFiveCouple = HeightFiveCouple;
//        }

        StrFiveCouple.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleThursday = StrSixCouple.createCell(4);
        cellSixCoupleThursday.setCellValue(CoupleSixThursday);
        cellSixCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(4,ColumnWidth);

//        if(MaxHeightSixCouple < HeightSixCouple){
//            MaxHeightSixCouple = HeightSixCouple;
//        }

        StrSixCouple.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleThursday = StrSevenCouple.createCell(4);
        cellSevenCoupleThursday.setCellValue(CoupleSevenThursday);
        cellSevenCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(4,ColumnWidth);

//        if(MaxHeightSevenCouple < HeightSevenCouple){
//            MaxHeightSevenCouple = HeightSevenCouple;
//        }

        StrSevenCouple.setHeightInPoints(HeightPoints);

        CoupleOne.clear();
        CoupleTwo.clear();
        CoupleThree.clear();
        CoupleFour.clear();
        CoupleFive.clear();
        CoupleSix.clear();
        CoupleSeven.clear();

//        HeightOneCouple = HeightPoints;
//        HeightTwoCouple = HeightPoints;
//        HeightThreeCouple = HeightPoints;
//        HeightFourCouple = HeightPoints;
//        HeightFiveCouple = HeightPoints;
//        HeightSixCouple = HeightPoints;
//        HeightSevenCouple = HeightPoints;

        // Couple Friday

        String CoupleOneFriday = "";
        String CoupleTwoFriday = "";
        String CoupleThreeFriday = "";
        String CoupleFourFriday = "";
        String CoupleFiveFriday = "";
        String CoupleSixFriday = "";
        String CoupleSevenFriday = "";

        for(int i = 0; i < ArrayFriday.size(); i++){

            switch (ArrayFriday.get(i).GetCoupleNumber()) {
                case (1): {
                    CoupleOne.add(ArrayFriday.get(i));
                    break;
                }
                case (2): {
                    CoupleTwo.add(ArrayFriday.get(i));
                    break;
                }
                case (3): {
                    CoupleThree.add(ArrayFriday.get(i));
                    break;
                }
                case (4): {
                    CoupleFour.add(ArrayFriday.get(i));
                    break;
                }
                case (5): {
                    CoupleFive.add(ArrayFriday.get(i));
                    break;
                }
                case (6): {
                    CoupleSix.add(ArrayFriday.get(i));
                    break;
                }
                case (7): {
                    CoupleSeven.add(ArrayFriday.get(i));
                    break;
                }
            }
        }

        ArrayList<String> ArrayStrFriday = ConcatCoupleDayWeek(CoupleOne,CoupleTwo,CoupleThree,CoupleFour,CoupleFive,CoupleSix,CoupleSeven);

        CoupleOneFriday = ArrayStrFriday.get(0);
        CoupleTwoFriday = ArrayStrFriday.get(1);
        CoupleThreeFriday = ArrayStrFriday.get(2);
        CoupleFourFriday = ArrayStrFriday.get(3);
        CoupleFiveFriday = ArrayStrFriday.get(4);
        CoupleSixFriday = ArrayStrFriday.get(5);
        CoupleSevenFriday = ArrayStrFriday.get(6);

        Cell cellOneCoupleFriday = StrOneCouple.createCell(5);
        cellOneCoupleFriday.setCellValue(CoupleOneFriday);
        cellOneCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(5,ColumnWidth);

//        if(MaxHeightOneCouple < HeightOneCouple){
//            MaxHeightOneCouple = HeightOneCouple;
//        }

        StrOneCouple.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleFriday = StrTwoCouple.createCell(5);
        cellTwoCoupleFriday.setCellValue(CoupleTwoFriday);
        cellTwoCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(5,ColumnWidth);

//        if(MaxHeightTwoCouple < HeightTwoCouple){
//            MaxHeightTwoCouple = HeightTwoCouple;
//        }

        StrTwoCouple.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleFriday = StrThreeCouple.createCell(5);
        cellThreeCoupleFriday.setCellValue(CoupleThreeFriday);
        cellThreeCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(5,ColumnWidth);

//        if(MaxHeightThreeCouple < HeightThreeCouple){
//            MaxHeightThreeCouple = HeightThreeCouple;
//        }

        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleFriday = StrFourCouple.createCell(5);
        cellFourCoupleFriday.setCellValue(CoupleFourFriday);
        cellFourCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(5,ColumnWidth);

//        if(MaxHeightFourCouple < HeightFourCouple){
//            MaxHeightFourCouple = HeightFourCouple;
//        }

        StrFourCouple.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleFriday = StrFiveCouple.createCell(5);
        cellFiveCoupleFriday.setCellValue(CoupleFiveFriday);
        cellFiveCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(5,ColumnWidth);

//        if(MaxHeightFiveCouple < HeightFiveCouple){
//            MaxHeightFiveCouple = HeightFiveCouple;
//        }

        StrFiveCouple.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleFriday = StrSixCouple.createCell(5);
        cellSixCoupleFriday.setCellValue(CoupleSixFriday);
        cellSixCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(5,ColumnWidth);

//        if(MaxHeightSixCouple < HeightSixCouple){
//            MaxHeightSixCouple = HeightSixCouple;
//        }

        StrSixCouple.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleFriday = StrSevenCouple.createCell(5);
        cellSevenCoupleFriday.setCellValue(CoupleSevenFriday);
        cellSevenCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(5,ColumnWidth);

//        if(MaxHeightSevenCouple < HeightSevenCouple){
//            MaxHeightSevenCouple = HeightSevenCouple;
//        }

        StrSevenCouple.setHeightInPoints(HeightPoints);

        CoupleOne.clear();
        CoupleTwo.clear();
        CoupleThree.clear();
        CoupleFour.clear();
        CoupleFive.clear();
        CoupleSix.clear();
        CoupleSeven.clear();

//        HeightOneCouple = HeightPoints;
//        HeightTwoCouple = HeightPoints;
//        HeightThreeCouple = HeightPoints;
//        HeightFourCouple = HeightPoints;
//        HeightFiveCouple = HeightPoints;
//        HeightSixCouple = HeightPoints;
//        HeightSevenCouple = HeightPoints;

        // Couple Saturday

        String CoupleOneSaturday = "";
        String CoupleTwoSaturday = "";
        String CoupleThreeSaturday = "";
        String CoupleFourSaturday = "";
        String CoupleFiveSaturday = "";
        String CoupleSixSaturday = "";
        String CoupleSevenSaturday = "";

        for(int i = 0; i < ArraySaturday.size(); i++){

            switch (ArraySaturday.get(i).GetCoupleNumber()) {
                case (1): {
                    CoupleOne.add(ArraySaturday.get(i));
                    break;
                }
                case (2): {
                    CoupleTwo.add(ArraySaturday.get(i));
                    break;
                }
                case (3): {
                    CoupleThree.add(ArraySaturday.get(i));
                    break;
                }
                case (4): {
                    CoupleFour.add(ArraySaturday.get(i));
                    break;
                }
                case (5): {
                    CoupleFive.add(ArraySaturday.get(i));
                    break;
                }
                case (6): {
                    CoupleSix.add(ArraySaturday.get(i));
                    break;
                }
                case (7): {
                    CoupleSeven.add(ArraySaturday.get(i));
                    break;
                }
            }
        }

        ArrayList<String> ArrayStrSaturday = ConcatCoupleDayWeek(CoupleOne,CoupleTwo,CoupleThree,CoupleFour,CoupleFive,CoupleSix,CoupleSeven);

        CoupleOneSaturday = ArrayStrSaturday.get(0);
        CoupleTwoSaturday = ArrayStrSaturday.get(1);
        CoupleThreeSaturday = ArrayStrSaturday.get(2);
        CoupleFourSaturday = ArrayStrSaturday.get(3);
        CoupleFiveSaturday = ArrayStrSaturday.get(4);
        CoupleSixSaturday = ArrayStrSaturday.get(5);
        CoupleSevenSaturday = ArrayStrSaturday.get(6);

        Cell cellOneCoupleSaturday = StrOneCouple.createCell(6);
        cellOneCoupleSaturday.setCellValue(CoupleOneSaturday);
        cellOneCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(6,ColumnWidth);

//        if(MaxHeightOneCouple < HeightOneCouple){
//            MaxHeightOneCouple = HeightOneCouple;
//        }

        StrOneCouple.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleSaturday = StrTwoCouple.createCell(6);
        cellTwoCoupleSaturday.setCellValue(CoupleTwoSaturday);
        cellTwoCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(6,ColumnWidth);

//        if(MaxHeightTwoCouple < HeightTwoCouple){
//            MaxHeightTwoCouple = HeightTwoCouple;
//        }

        StrTwoCouple.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleSaturday = StrThreeCouple.createCell(6);
        cellThreeCoupleSaturday.setCellValue(CoupleThreeSaturday);
        cellThreeCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(6,ColumnWidth);

//        if(MaxHeightThreeCouple < HeightThreeCouple){
//            MaxHeightThreeCouple = HeightThreeCouple;
//        }

        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleSaturday = StrFourCouple.createCell(6);
        cellFourCoupleSaturday.setCellValue(CoupleFourSaturday);
        cellFourCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(6,ColumnWidth);

//        if(MaxHeightFourCouple < HeightFourCouple){
//            MaxHeightFourCouple = HeightFourCouple;
//        }

        StrFourCouple.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleSaturday = StrFiveCouple.createCell(6);
        cellFiveCoupleSaturday.setCellValue(CoupleFiveSaturday);
        cellFiveCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(6,ColumnWidth);

//        if(MaxHeightFiveCouple < HeightFiveCouple){
//            MaxHeightFiveCouple = HeightFiveCouple;
//        }

        StrFiveCouple.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleSaturday = StrSixCouple.createCell(6);
        cellSixCoupleSaturday.setCellValue(CoupleSixSaturday);
        cellSixCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(6,ColumnWidth);

//        if(MaxHeightSixCouple < HeightSixCouple){
//            MaxHeightSixCouple = HeightSixCouple;
//        }

        StrSixCouple.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleSaturday = StrSevenCouple.createCell(6);
        cellSevenCoupleSaturday.setCellValue(CoupleSevenSaturday);
        cellSevenCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(6,ColumnWidth);

//        if(MaxHeightSevenCouple < HeightSevenCouple){
//            MaxHeightSevenCouple = HeightSevenCouple;
//        }

        StrSevenCouple.setHeightInPoints(HeightPoints);

        CoupleOne.clear();
        CoupleTwo.clear();
        CoupleThree.clear();
        CoupleFour.clear();
        CoupleFive.clear();
        CoupleSix.clear();
        CoupleSeven.clear();

        String separator = File.separator;

        Date date = new Date();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
        File file = new File(simpleDateFormat.format(date));

        FileOutputStream fileOutputStream = new FileOutputStream("TableClassroom(s)" + separator + "One Classroom" + separator + "Аудитория " + workbookOneClassroom.getSheetName(0) + " (" + file + ")");

        workbookOneClassroom.write(fileOutputStream);
        fileOutputStream.close();

    }
}
