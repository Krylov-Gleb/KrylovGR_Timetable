package timetablekrylov.timetablekrylovgr;

import javafx.css.PseudoClass;
import javafx.scene.control.CheckBox;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

public class CreatorTableExelClassrooms {

    private Workbook workbookClassroom = new HSSFWorkbook();

    String Aud = "";

    private Sheet Classroom = workbookClassroom.createSheet("Аудитории");

    private Row rowZero = Classroom.createRow(0);
    private Cell cellDayWeek = rowZero.createCell(0);
    private Cell cellGroup = rowZero.createCell(1);

    private Row rowOne = Classroom.createRow(1);
    private Cell Monday = rowOne.createCell(0);
    private Cell CoupleOneMonday = rowOne.createCell(1);

    private Row rowTwo = Classroom.createRow(2);
    private Cell CoupleTwoMonday = rowTwo.createCell(1);

    private Row rowThree = Classroom.createRow(3);
    private Cell CoupleThreeMonday = rowThree.createCell(1);

    private Row rowFour = Classroom.createRow(4);
    private Cell CoupleFourMonday = rowFour.createCell(1);

    private Row rowFive = Classroom.createRow(5);
    private Cell CoupleFiveMonday = rowFive.createCell(1);

    private Row rowSix = Classroom.createRow(6);
    private Cell CoupleSixMonday = rowSix.createCell(1);

    private Row rowSeven = Classroom.createRow(7);
    private Cell CoupleSevenMonday = rowSeven.createCell(1);

    // ----------------------------------------------------------------------------

    private Row rowEight = Classroom.createRow(8);
    private Cell Tuesday = rowEight.createCell(0);
    private Cell CoupleOneTuesday = rowEight.createCell(1);

    private Row rowNine = Classroom.createRow(9);
    private Cell CoupleTwoTuesday = rowNine.createCell(1);

    private Row rowTen = Classroom.createRow(10);
    private Cell CoupleThreeTuesday = rowTen.createCell(1);

    private Row rowEleven = Classroom.createRow(11);
    private Cell CoupleFourTuesday = rowEleven.createCell(1);

    private Row rowTwelve = Classroom.createRow(12);
    private Cell CoupleFiveTuesday = rowTwelve.createCell(1);

    private Row rowThirteen = Classroom.createRow(13);
    private Cell CoupleSixTuesday = rowThirteen.createCell(1);

    private Row rowFourteen = Classroom.createRow(14);
    private Cell CoupleSevenTuesday = rowFourteen.createCell(1);

    // ----------------------------------------------------------------------------

    private Row rowfifteen = Classroom.createRow(15);
    private Cell Wednesday = rowfifteen.createCell(0);
    private Cell CoupleOneWednesday = rowfifteen.createCell(1);

    private Row rowSixteen = Classroom.createRow(16);
    private Cell CoupleTwoWednesday = rowSixteen.createCell(1);

    private Row rowSeventeen = Classroom.createRow(17);
    private Cell CoupleThreeWednesday = rowSeventeen.createCell(1);

    private Row rowEighteen = Classroom.createRow(18);
    private Cell CoupleFourWednesday = rowEighteen.createCell(1);

    private Row rowNineteen = Classroom.createRow(19);
    private Cell CoupleFiveWednesday = rowNineteen.createCell(1);

    private Row rowTwenty = Classroom.createRow(20);
    private Cell CoupleSixWednesday = rowTwenty.createCell(1);

    private Row rowTwentyOne = Classroom.createRow(21);
    private Cell CoupleSevenWednesday = rowTwentyOne.createCell(1);

    // ----------------------------------------------------------------------------

    private Row rowTwentyTwo = Classroom.createRow(22);
    private Cell Thursday = rowTwentyTwo.createCell(0);
    private Cell CoupleOneThursday = rowTwentyTwo.createCell(1);

    private Row rowTwentyThree = Classroom.createRow(23);
    private Cell CoupleTwoThursday = rowTwentyThree.createCell(1);

    private Row rowTwentyFour = Classroom.createRow(24);
    private Cell CoupleThreeThursday = rowTwentyFour.createCell(1);

    private Row rowTwentyFive = Classroom.createRow(25);
    private Cell CoupleFourThursday = rowTwentyFive.createCell(1);

    private Row rowTwentySix = Classroom.createRow(26);
    private Cell CoupleFiveThursday = rowTwentySix.createCell(1);

    private Row rowTwentySeven = Classroom.createRow(27);
    private Cell CoupleSixThursday = rowTwentySeven.createCell(1);

    private Row rowTwentyEight = Classroom.createRow(28);
    private Cell CoupleSevenThursday = rowTwentyEight.createCell(1);


    // ----------------------------------------------------------------------------

    private Row rowTwentyNine = Classroom.createRow(29);
    private Cell Friday = rowTwentyNine.createCell(0);
    private Cell CoupleOneFriday = rowTwentyNine.createCell(1);

    private Row rowThirty = Classroom.createRow(30);
    private Cell CoupleTwoFriday = rowThirty.createCell(1);

    private Row rowThirtyOne = Classroom.createRow(31);
    private Cell CoupleThreeFriday = rowThirtyOne.createCell(1);

    private Row rowThirtyTwo = Classroom.createRow(32);
    private Cell CoupleFourFriday = rowThirtyTwo.createCell(1);

    private Row rowThirtyThree = Classroom.createRow(33);
    private Cell CoupleFiveFriday = rowThirtyThree.createCell(1);

    private Row rowThirtyFour = Classroom.createRow(34);
    private Cell CoupleSixFriday = rowThirtyFour.createCell(1);

    private Row rowThirtyFive = Classroom.createRow(35);
    private Cell CoupleSevenFriday = rowThirtyFive.createCell(1);

    // ---------------------------------------------------------------------------

    private Row rowThirtySix = Classroom.createRow(36);
    private Cell Saturday = rowThirtySix.createCell(0);
    private Cell CoupleOneSaturday = rowThirtySix.createCell(1);

    private Row rowThirtySeven = Classroom.createRow(37);
    private Cell CoupleTwoSaturday = rowThirtySeven.createCell(1);

    private Row rowThirtyEight = Classroom.createRow(38);
    private Cell CoupleThreeSaturday = rowThirtyEight.createCell(1);

    private Row rowThirtyNine = Classroom.createRow(39);
    private Cell CoupleFourSaturday = rowThirtyNine.createCell(1);

    private Row rowForty = Classroom.createRow(40);
    private Cell CoupleFiveSaturday = rowForty.createCell(1);

    private Row rowFortyOne = Classroom.createRow(41);
    private Cell CoupleSixSaturday = rowFortyOne.createCell(1);

    private Row rowFortyTwo = Classroom.createRow(42);
    private Cell CoupleSevenSaturday = rowFortyTwo.createCell(1);

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

    private void CreatorTable(int ClassroomNumber, ArrayList<CoupleGroup> Array){

        int HeightPoints = 50;
        int ColumnWidth = 12000;

        ArrayList<CoupleGroup> CoupleOne = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleTwo = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleThree = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleFour = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleFive = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleSix = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleSeven = new ArrayList<>();

        CellStyle cellStyle = workbookClassroom.createCellStyle();
        cellStyle.setWrapText(true);

        ArrayList<CoupleGroup> ArrayMonday = new ArrayList<>();
        ArrayList<CoupleGroup> ArrayTuesday = new ArrayList<>();
        ArrayList<CoupleGroup> ArrayWednesday = new ArrayList<>();
        ArrayList<CoupleGroup> ArrayThursday = new ArrayList<>();
        ArrayList<CoupleGroup> ArrayFriday = new ArrayList<>();
        ArrayList<CoupleGroup> ArraySaturday = new ArrayList<>();

        for (int index = 0; index < Array.size(); index++) {

            int IdDay = Array.get(index).GetIDDay();

            switch (IdDay) {
                case (1): {
                    ArrayMonday.add(Array.get(index));
                    break;
                }
                case (2): {
                    ArrayTuesday.add(Array.get(index));
                    break;
                }
                case (3): {
                    ArrayWednesday.add(Array.get(index));
                    break;
                }
                case (4): {
                    ArrayThursday.add(Array.get(index));
                    break;
                }
                case (5): {
                    ArrayFriday.add(Array.get(index));
                    break;
                }
                case (6): {
                    ArraySaturday.add(Array.get(index));
                    break;
                }
            }
        }

        String CoupleOneMonday = "";
        String CoupleTwoMonday = "";
        String CoupleThreeMonday = "";
        String CoupleFourMonday = "";
        String CoupleFiveMonday = "";
        String CoupleSixMonday = "";
        String CoupleSevenMonday = "";

        for (int index2 = 0; index2 < ArrayMonday.size(); index2++) {

            switch (ArrayMonday.get(index2).GetCoupleNumber()) {
                case (1): {
                    CoupleOne.add(ArrayMonday.get(index2));
                    break;
                }
                case (2): {
                    CoupleTwo.add(ArrayMonday.get(index2));
                    break;
                }
                case (3): {
                    CoupleThree.add(ArrayMonday.get(index2));
                    break;
                }
                case (4): {
                    CoupleFour.add(ArrayMonday.get(index2));
                    break;
                }
                case (5): {
                    CoupleFive.add(ArrayMonday.get(index2));
                    break;
                }
                case (6): {
                    CoupleSix.add(ArrayMonday.get(index2));
                    break;
                }
                case (7): {
                    CoupleSeven.add(ArrayMonday.get(index2));
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

        Cell cellOneCoupleMonday = rowOne.createCell(ClassroomNumber);
        cellOneCoupleMonday.setCellValue(CoupleOneMonday);
        cellOneCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightOneCouple < HeightOneCouple){
//            MaxHeightOneCouple = HeightOneCouple;
//        }

        rowOne.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleMonday = rowTwo.createCell(ClassroomNumber);
        cellTwoCoupleMonday.setCellValue(CoupleTwoMonday);
        cellTwoCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightTwoCouple < HeightTwoCouple){
//            MaxHeightTwoCouple = HeightTwoCouple;
//        }

        rowTwo.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleMonday = rowThree.createCell(ClassroomNumber);
        cellThreeCoupleMonday.setCellValue(CoupleThreeMonday);
        cellThreeCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightThreeCouple < HeightThreeCouple){
//            MaxHeightThreeCouple = HeightThreeCouple;
//        }

        rowThree.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleMonday = rowFour.createCell(ClassroomNumber);
        cellFourCoupleMonday.setCellValue(CoupleFourMonday);
        cellFourCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightFourCouple < HeightFourCouple){
//            MaxHeightFourCouple = HeightFourCouple;
//        }

        rowFour.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleMonday = rowFive.createCell(ClassroomNumber);
        cellFiveCoupleMonday.setCellValue(CoupleFiveMonday);
        cellFiveCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightFiveCouple < HeightFiveCouple){
//            MaxHeightFiveCouple = HeightFiveCouple;
//        }

        rowFive.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleMonday = rowSix.createCell(ClassroomNumber);
        cellSixCoupleMonday.setCellValue(CoupleSixMonday);
        cellSixCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightSixCouple < HeightSixCouple){
//            MaxHeightSixCouple = HeightSixCouple;
//        }

        rowSix.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleMonday = rowSeven.createCell(ClassroomNumber);
        cellSevenCoupleMonday.setCellValue(CoupleSevenMonday);
        cellSevenCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightSevenCouple < HeightSevenCouple){
//            MaxHeightSevenCouple = HeightSevenCouple;
//        }

        rowSeven.setHeightInPoints(HeightPoints);

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

        // Cell Tuesday

        String CoupleOneTuesday = "";
        String CoupleTwoTuesday = "";
        String CoupleThreeTuesday = "";
        String CoupleFourTuesday = "";
        String CoupleFiveTuesday = "";
        String CoupleSixTuesday = "";
        String CoupleSevenTuesday = "";

        for(int index2 = 0; index2 < ArrayTuesday.size(); index2++){

            switch (ArrayTuesday.get(index2).GetCoupleNumber()) {
                case (1): {
                    CoupleOne.add(ArrayTuesday.get(index2));
                    break;
                }
                case (2): {
                    CoupleTwo.add(ArrayTuesday.get(index2));
                    break;
                }
                case (3): {
                    CoupleThree.add(ArrayTuesday.get(index2));
                    break;
                }
                case (4): {
                    CoupleFour.add(ArrayTuesday.get(index2));
                    break;
                }
                case (5): {
                    CoupleFive.add(ArrayTuesday.get(index2));
                    break;
                }
                case (6): {
                    CoupleSix.add(ArrayTuesday.get(index2));
                    break;
                }
                case (7): {
                    CoupleSeven.add(ArrayTuesday.get(index2));
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

        Cell cellOneCoupleTuesday = rowEight.createCell(ClassroomNumber);
        cellOneCoupleTuesday.setCellValue(CoupleOneTuesday);
        cellOneCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightOneCouple < HeightOneCouple){
//            MaxHeightOneCouple = HeightOneCouple;
//        }

        rowEight.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleTuesday = rowNine.createCell(ClassroomNumber);
        cellTwoCoupleTuesday.setCellValue(CoupleTwoTuesday);
        cellTwoCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightTwoCouple < HeightTwoCouple){
//            MaxHeightTwoCouple = HeightTwoCouple;
//        }

        rowNine.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleTuesday = rowTen.createCell(ClassroomNumber);
        cellThreeCoupleTuesday.setCellValue(CoupleThreeTuesday);
        cellThreeCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightThreeCouple < HeightThreeCouple){
//            MaxHeightThreeCouple = HeightThreeCouple;
//        }

        rowTen.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleTuesday = rowEleven.createCell(ClassroomNumber);
        cellFourCoupleTuesday.setCellValue(CoupleFourTuesday);
        cellFourCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightFourCouple < HeightFourCouple){
//            MaxHeightFourCouple = HeightFourCouple;
//        }

        rowEleven.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleTuesday = rowTwelve.createCell(ClassroomNumber);
        cellFiveCoupleTuesday.setCellValue(CoupleFiveTuesday);
        cellFiveCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightFiveCouple < HeightFiveCouple){
//            MaxHeightFiveCouple = HeightFiveCouple;
//        }

        rowTwelve.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleTuesday = rowThirteen.createCell(ClassroomNumber);
        cellSixCoupleTuesday.setCellValue(CoupleSixTuesday);
        cellSixCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightSixCouple < HeightSixCouple){
//            MaxHeightSixCouple = HeightSixCouple;
//        }

        rowThirteen.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleTuesday = rowFourteen.createCell(ClassroomNumber);
        cellSevenCoupleTuesday.setCellValue(CoupleSevenTuesday);
        cellSevenCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightSevenCouple < HeightSevenCouple){
//            MaxHeightSevenCouple = HeightSevenCouple;
//        }

        rowFourteen.setHeightInPoints(HeightPoints);

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

        // Cell Wednesday

        String CoupleOneWednesday = "";
        String CoupleTwoWednesday = "";
        String CoupleThreeWednesday = "";
        String CoupleFourWednesday = "";
        String CoupleFiveWednesday = "";
        String CoupleSixWednesday = "";
        String CoupleSevenWednesday = "";

        for(int index2 = 0; index2 < ArrayWednesday.size(); index2++){

            switch (ArrayWednesday.get(index2).GetCoupleNumber()) {
                case (1): {
                    CoupleOne.add(ArrayWednesday.get(index2));
                    break;
                }
                case (2): {
                    CoupleTwo.add(ArrayWednesday.get(index2));
                    break;
                }
                case (3): {
                    CoupleThree.add(ArrayWednesday.get(index2));
                    break;
                }
                case (4): {
                    CoupleFour.add(ArrayWednesday.get(index2));
                    break;
                }
                case (5): {
                    CoupleFive.add(ArrayWednesday.get(index2));
                    break;
                }
                case (6): {
                    CoupleSix.add(ArrayWednesday.get(index2));
                    break;
                }
                case (7): {
                    CoupleSeven.add(ArrayWednesday.get(index2));
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

        Cell cellOneCoupleWednesday = rowfifteen.createCell(ClassroomNumber);
        cellOneCoupleWednesday.setCellValue(CoupleOneWednesday);
        cellOneCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightOneCouple < HeightOneCouple){
//            MaxHeightOneCouple = HeightOneCouple;
//        }

        rowfifteen.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleWednesday = rowSixteen.createCell(ClassroomNumber);
        cellTwoCoupleWednesday.setCellValue(CoupleTwoWednesday);
        cellTwoCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightTwoCouple < HeightTwoCouple){
//            MaxHeightTwoCouple = HeightTwoCouple;
//        }

        rowSixteen.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleWednesday = rowSeventeen.createCell(ClassroomNumber);
        cellThreeCoupleWednesday.setCellValue(CoupleThreeWednesday);
        cellThreeCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightThreeCouple < HeightThreeCouple){
//            MaxHeightThreeCouple = HeightThreeCouple;
//        }

        rowSeventeen.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleWednesday = rowEighteen.createCell(ClassroomNumber);
        cellFourCoupleWednesday.setCellValue(CoupleFourWednesday);
        cellFourCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightFourCouple < HeightFourCouple){
//            MaxHeightFourCouple = HeightFourCouple;
//        }

        rowEighteen.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleWednesday = rowNineteen.createCell(ClassroomNumber);
        cellFiveCoupleWednesday.setCellValue(CoupleFiveWednesday);
        cellFiveCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightFiveCouple < HeightFiveCouple){
//            MaxHeightFiveCouple = HeightFiveCouple;
//        }

        rowNineteen.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleWednesday = rowTwenty.createCell(ClassroomNumber);
        cellSixCoupleWednesday.setCellValue(CoupleSixWednesday);
        cellSixCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightSixCouple < HeightSixCouple){
//            MaxHeightSixCouple = HeightSixCouple;
//        }

        rowTwenty.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleWednesday = rowTwentyOne.createCell(ClassroomNumber);
        cellSevenCoupleWednesday.setCellValue(CoupleSevenWednesday);
        cellSevenCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightSevenCouple < HeightSevenCouple){
//            MaxHeightSevenCouple = HeightSevenCouple;
//        }

        rowTwentyOne.setHeightInPoints(HeightPoints);

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

        // Cell Thursday

        String CoupleOneThursday = "";
        String CoupleTwoThursday = "";
        String CoupleThreeThursday = "";
        String CoupleFourThursday = "";
        String CoupleFiveThursday = "";
        String CoupleSixThursday = "";
        String CoupleSevenThursday = "";

        for(int index2 = 0; index2 < ArrayThursday.size(); index2++){

            switch (ArrayThursday.get(index2).GetCoupleNumber()) {
                case (1): {
                    CoupleOne.add(ArrayThursday.get(index2));
                    break;
                }
                case (2): {
                    CoupleTwo.add(ArrayThursday.get(index2));
                    break;
                }
                case (3): {
                    CoupleThree.add(ArrayThursday.get(index2));
                    break;
                }
                case (4): {
                    CoupleFour.add(ArrayThursday.get(index2));
                    break;
                }
                case (5): {
                    CoupleFive.add(ArrayThursday.get(index2));
                    break;
                }
                case (6): {
                    CoupleSix.add(ArrayThursday.get(index2));
                    break;
                }
                case (7): {
                    CoupleSeven.add(ArrayThursday.get(index2));
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

        Cell cellOneCoupleThursday = rowTwentyTwo.createCell(ClassroomNumber);
        cellOneCoupleThursday.setCellValue(CoupleOneThursday);
        cellOneCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightOneCouple < HeightOneCouple){
//            MaxHeightOneCouple = HeightOneCouple;
//        }

        rowTwentyTwo.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleThursday = rowTwentyThree.createCell(ClassroomNumber);
        cellTwoCoupleThursday.setCellValue(CoupleTwoThursday);
        cellTwoCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightTwoCouple < HeightTwoCouple){
//            MaxHeightTwoCouple = HeightTwoCouple;
//        }

        rowTwentyThree.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleThursday = rowTwentyFour.createCell(ClassroomNumber);
        cellThreeCoupleThursday.setCellValue(CoupleThreeThursday);
        cellThreeCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightThreeCouple < HeightThreeCouple){
//            MaxHeightThreeCouple = HeightThreeCouple;
//        }

        rowTwentyFour.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleThursday = rowTwentyFive.createCell(ClassroomNumber);
        cellFourCoupleThursday.setCellValue(CoupleFourThursday);
        cellFourCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightFourCouple < HeightFourCouple){
//            MaxHeightFourCouple = HeightFourCouple;
//        }

        rowTwentyFive.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleThursday = rowTwentySix.createCell(ClassroomNumber);
        cellFiveCoupleThursday.setCellValue(CoupleFiveThursday);
        cellFiveCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightFiveCouple < HeightFiveCouple){
//            MaxHeightFiveCouple = HeightFiveCouple;
//        }

        rowTwentySix.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleThursday = rowTwentySeven.createCell(ClassroomNumber);
        cellSixCoupleThursday.setCellValue(CoupleSixThursday);
        cellSixCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightSixCouple < HeightSixCouple){
//            MaxHeightSixCouple = HeightSixCouple;
//        }

        rowTwentySeven.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleThursday = rowTwentyEight.createCell(ClassroomNumber);
        cellSevenCoupleThursday.setCellValue(CoupleSevenThursday);
        cellSevenCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightSevenCouple < HeightSevenCouple){
//            MaxHeightSevenCouple = HeightSevenCouple;
//        }

        rowTwentyEight.setHeightInPoints(HeightPoints);

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

        // Cell Friday

        String CoupleOneFriday = "";
        String CoupleTwoFriday = "";
        String CoupleThreeFriday = "";
        String CoupleFourFriday = "";
        String CoupleFiveFriday = "";
        String CoupleSixFriday = "";
        String CoupleSevenFriday = "";

        for(int index2 = 0; index2 < ArrayFriday.size(); index2++){

            switch (ArrayFriday.get(index2).GetCoupleNumber()) {
                case (1): {
                    CoupleOne.add(ArrayFriday.get(index2));
                    break;
                }
                case (2): {
                    CoupleTwo.add(ArrayFriday.get(index2));
                    break;
                }
                case (3): {
                    CoupleThree.add(ArrayFriday.get(index2));
                    break;
                }
                case (4): {
                    CoupleFour.add(ArrayFriday.get(index2));
                    break;
                }
                case (5): {
                    CoupleFive.add(ArrayFriday.get(index2));
                    break;
                }case (6): {
                    CoupleSix.add(ArrayFriday.get(index2));
                    break;
                }
                case (7): {
                    CoupleSeven.add(ArrayFriday.get(index2));
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

        Cell cellOneCoupleFriday = rowTwentyNine.createCell(ClassroomNumber);
        cellOneCoupleFriday.setCellValue(CoupleOneFriday);
        cellOneCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightOneCouple < HeightOneCouple){
//            MaxHeightOneCouple = HeightOneCouple;
//        }

        rowTwentyNine.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleFriday = rowThirty.createCell(ClassroomNumber);
        cellTwoCoupleFriday.setCellValue(CoupleTwoFriday);
        cellTwoCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightTwoCouple < HeightTwoCouple){
//            MaxHeightTwoCouple = HeightTwoCouple;
//        }

        rowThirty.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleFriday = rowThirtyOne.createCell(ClassroomNumber);
        cellThreeCoupleFriday.setCellValue(CoupleThreeFriday);
        cellThreeCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightThreeCouple < HeightThreeCouple){
//            MaxHeightThreeCouple = HeightThreeCouple;
//        }

        rowThirtyOne.setHeightInPoints(HeightPoints);


        Cell cellFourCoupleFriday = rowThirtyTwo.createCell(ClassroomNumber);
        cellFourCoupleFriday.setCellValue(CoupleFourFriday);
        cellFourCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightFourCouple < HeightFourCouple){
//            MaxHeightFourCouple = HeightFourCouple;
//        }

        rowThirtyTwo.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleFriday = rowThirtyThree.createCell(ClassroomNumber);
        cellFiveCoupleFriday.setCellValue(CoupleFiveFriday);
        cellFiveCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightFiveCouple < HeightFiveCouple){
//            MaxHeightFiveCouple = HeightFiveCouple;
//        }

        rowThirtyThree.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleFriday = rowThirtyFour.createCell(ClassroomNumber);
        cellSixCoupleFriday.setCellValue(CoupleSixFriday);
        cellSixCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightSixCouple < HeightSixCouple){
//            MaxHeightSixCouple = HeightSixCouple;
//        }

        rowThirtyFour.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleFriday = rowThirtyFive.createCell(ClassroomNumber);
        cellSevenCoupleFriday.setCellValue(CoupleSevenFriday);
        cellSevenCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightSevenCouple < HeightSevenCouple){
//            MaxHeightSevenCouple = HeightSevenCouple;
//        }

        rowThirtyFive.setHeightInPoints(HeightPoints);

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

        // Cell Saturday

        String CoupleOneSaturday = "";
        String CoupleTwoSaturday = "";
        String CoupleThreeSaturday = "";
        String CoupleFourSaturday = "";
        String CoupleFiveSaturday = "";
        String CoupleSixSaturday = "";
        String CoupleSevenSaturday = "";

        for(int index2 = 0; index2 < ArraySaturday.size(); index2++){

            switch (ArraySaturday.get(index2).GetCoupleNumber()) {
                case (1): {
                    CoupleOne.add(ArraySaturday.get(index2));
                    break;
                }
                case (2): {
                    CoupleTwo.add(ArraySaturday.get(index2));
                    break;
                }
                case (3): {
                    CoupleThree.add(ArraySaturday.get(index2));
                    break;
                }
                case (4): {
                    CoupleFour.add(ArraySaturday.get(index2));
                    break;
                }
                case (5): {
                    CoupleFive.add(ArraySaturday.get(index2));
                    break;
                }
                case (6): {
                    CoupleSix.add(ArraySaturday.get(index2));
                    break;
                }
                case (7): {
                    CoupleSeven.add(ArraySaturday.get(index2));
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

        Cell cellOneCoupleSaturday = rowThirtySix.createCell(ClassroomNumber);
        cellOneCoupleSaturday.setCellValue(CoupleOneSaturday);
        cellOneCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightOneCouple < HeightOneCouple){
//            MaxHeightOneCouple = HeightOneCouple;
//        }

        rowThirtySix.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleSaturday = rowThirtySeven.createCell(ClassroomNumber);
        cellTwoCoupleSaturday.setCellValue(CoupleTwoSaturday);
        cellTwoCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightTwoCouple < HeightTwoCouple){
//            MaxHeightTwoCouple = HeightTwoCouple;
//        }

        rowThirtySeven.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleSaturday = rowThirtyEight.createCell(ClassroomNumber);
        cellThreeCoupleSaturday.setCellValue(CoupleThreeSaturday);
        cellThreeCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightThreeCouple < HeightThreeCouple){
//            MaxHeightThreeCouple = HeightThreeCouple;
//        }

        rowThirtyEight.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleSaturday = rowThirtyNine.createCell(ClassroomNumber);
        cellFourCoupleSaturday.setCellValue(CoupleFourSaturday);
        cellFourCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightFourCouple < HeightFourCouple){
//            MaxHeightFourCouple = HeightFourCouple;
//        }

        rowThirtyNine.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleSaturday = rowForty.createCell(ClassroomNumber);
        cellFiveCoupleSaturday.setCellValue(CoupleFiveSaturday);
        cellFiveCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightFiveCouple < HeightFiveCouple){
//            MaxHeightFiveCouple = HeightFiveCouple;
//        }

        rowForty.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleSaturday = rowFortyOne.createCell(ClassroomNumber);
        cellSixCoupleSaturday.setCellValue(CoupleSixSaturday);
        cellSixCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightSixCouple < HeightSixCouple){
//            MaxHeightSixCouple = HeightSixCouple;
//        }

        rowFortyOne.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleSaturday = rowFortyTwo.createCell(ClassroomNumber);
        cellSevenCoupleSaturday.setCellValue(CoupleSevenSaturday);
        cellSevenCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(ClassroomNumber,ColumnWidth);

//        if(MaxHeightSevenCouple < HeightSevenCouple){
//            MaxHeightSevenCouple = HeightSevenCouple;
//        }

        rowFortyTwo.setHeightInPoints(HeightPoints);

        CoupleOne.clear();
        CoupleTwo.clear();
        CoupleThree.clear();
        CoupleFour.clear();
        CoupleFive.clear();
        CoupleSix.clear();
        CoupleSeven.clear();

//        ArrayList<Integer> ArrayMax = new ArrayList<>(7);
//        ArrayMax.add(MaxHeightOneCouple);
//        ArrayMax.add(MaxHeightTwoCouple);
//        ArrayMax.add(MaxHeightThreeCouple);
//        ArrayMax.add(MaxHeightFourCouple);
//        ArrayMax.add(MaxHeightFiveCouple);
//        ArrayMax.add(MaxHeightSixCouple);
//        ArrayMax.add(MaxHeightSevenCouple);

    }

    public void CreateTableExelClassroom(ArrayList<CoupleGroup> ArrayCouple, ArrayList<CheckBox> ArrayClassroomCheckBox) throws IOException {

        CellStyle cellStyle = workbookClassroom.createCellStyle();
        cellStyle.setWrapText(true);

        cellDayWeek.setCellStyle(cellStyle);
        cellDayWeek.setCellValue("День недели");
        Classroom.setColumnWidth(0,8000);
        rowZero.setHeightInPoints(30);

        cellGroup.setCellStyle(cellStyle);
        cellGroup.setCellValue("Номер пары");
        Classroom.setColumnWidth(1,8000);
        rowZero.setHeightInPoints(30);

        Monday.setCellValue("Понедельник");
        Tuesday.setCellValue("Вторник");
        Wednesday.setCellValue("Среда");
        Thursday.setCellValue("Четверг");
        Friday.setCellValue("Пятница");
        Saturday.setCellValue("Суббота");


        CoupleOneMonday.setCellValue("1 пара");
        CoupleTwoMonday.setCellValue("2 пара");
        CoupleThreeMonday.setCellValue("3 пара");
        CoupleFourMonday.setCellValue("4 пара");
        CoupleFiveMonday.setCellValue("5 пара");
        CoupleSixMonday.setCellValue("6 пара");
        CoupleSevenMonday.setCellValue("7 пара");

        CoupleOneTuesday.setCellValue("1 пара");
        CoupleTwoTuesday.setCellValue("2 пара");
        CoupleThreeTuesday.setCellValue("3 пара");
        CoupleFourTuesday.setCellValue("4 пара");
        CoupleFiveTuesday.setCellValue("5 пара");
        CoupleSixTuesday.setCellValue("6 пара");
        CoupleSevenTuesday.setCellValue("7 пара");

        CoupleOneWednesday.setCellValue("1 пара");
        CoupleTwoWednesday.setCellValue("2 пара");
        CoupleThreeWednesday.setCellValue("3 пара");
        CoupleFourWednesday.setCellValue("4 пара");
        CoupleFiveWednesday.setCellValue("5 пара");
        CoupleSixWednesday.setCellValue("6 пара");
        CoupleSevenWednesday.setCellValue("7 пара");

        CoupleOneThursday.setCellValue("1 пара");
        CoupleTwoThursday.setCellValue("2 пара");
        CoupleThreeThursday.setCellValue("3 пара");
        CoupleFourThursday.setCellValue("4 пара");
        CoupleFiveThursday.setCellValue("5 пара");
        CoupleSixThursday.setCellValue("6 пара");
        CoupleSevenThursday.setCellValue("7 пара");

        CoupleOneFriday.setCellValue("1 пара");
        CoupleTwoFriday.setCellValue("2 пара");
        CoupleThreeFriday.setCellValue("3 пара");
        CoupleFourFriday.setCellValue("4 пара");
        CoupleFiveFriday.setCellValue("5 пара");
        CoupleSixFriday.setCellValue("6 пара");
        CoupleSevenFriday.setCellValue("7 пара");

        CoupleOneSaturday.setCellValue("1 пара");
        CoupleTwoSaturday.setCellValue("2 пара");
        CoupleThreeSaturday.setCellValue("3 пара");
        CoupleFourSaturday.setCellValue("4 пара");
        CoupleFiveSaturday.setCellValue("5 пара");
        CoupleSixSaturday.setCellValue("6 пара");
        CoupleSevenSaturday.setCellValue("7 пара");

        Classroom.addMergedRegion(new CellRangeAddress(36,42,0,0));
        Classroom.addMergedRegion(new CellRangeAddress(29,35,0,0));
        Classroom.addMergedRegion(new CellRangeAddress(22,28,0,0));
        Classroom.addMergedRegion(new CellRangeAddress(15,21,0,0));
        Classroom.addMergedRegion(new CellRangeAddress(8,14,0,0));
        Classroom.addMergedRegion(new CellRangeAddress(1,7,0,0));

        ArrayList<CoupleGroup> Array410 = new ArrayList<>();
        ArrayList<CoupleGroup> Array411 = new ArrayList<>();
        ArrayList<CoupleGroup> Array413 = new ArrayList<>();
        ArrayList<CoupleGroup> Array416 = new ArrayList<>();
        ArrayList<CoupleGroup> Array417 = new ArrayList<>();

        for(int i = 0; i < ArrayCouple.size(); i++){
            switch(ArrayCouple.get(i).GetAud()){
                case ("\"ауд. 410/2\""):{
                    Array410.add(ArrayCouple.get(i));
                    break;
                }
                case("\"ауд. 411/2\""):{
                    Array411.add(ArrayCouple.get(i));
                    break;
                }
                case("\"ауд. 413/2\""):{
                    Array413.add(ArrayCouple.get(i));
                    break;
                }
                case ("\"ауд. 416/2\""):{
                    Array416.add(ArrayCouple.get(i));
                    break;
                }
                case("\"ауд. 417/2\""):{
                    Array417.add(ArrayCouple.get(i));
                    break;
                }
            }
        }

        int ClassroomNumber = 2;

        for(int i = 0; i < ArrayClassroomCheckBox.size(); i++){

            if(ArrayClassroomCheckBox.get(i).isSelected()) {

                Cell cell = rowZero.createCell(ClassroomNumber);
                cell.setCellValue(ArrayClassroomCheckBox.get(i).getText());

                if (ArrayClassroomCheckBox.get(i).getText().equals("410/2")) {
                    CreatorTable(ClassroomNumber,Array410);
//                    HeightOneCouple = Array.get(0);
//                    HeightTwoCouple = Array.get(1);
//                    HeightThreeCouple = Array.get(2);
//                    HeightFourCouple = Array.get(3);
//                    HeightFiveCouple = Array.get(4);
//                    HeightSixCouple = Array.get(5);
//                    HeightSevenCouple = Array.get(6);
                }

                if(ArrayClassroomCheckBox.get(i).getText().equals("411/2")) {
                    CreatorTable(ClassroomNumber,Array411);
//                    HeightOneCouple = Array.get(0);
//                    HeightTwoCouple = Array.get(1);
//                    HeightThreeCouple = Array.get(2);
//                    HeightFourCouple = Array.get(3);
//                    HeightFiveCouple = Array.get(4);
//                    HeightSixCouple = Array.get(5);
//                    HeightSevenCouple = Array.get(6);
                }

                if(ArrayClassroomCheckBox.get(i).getText().equals("413/2")) {
                    CreatorTable(ClassroomNumber,Array413);
//                    HeightOneCouple = Array.get(0);
//                    HeightTwoCouple = Array.get(1);
//                    HeightThreeCouple = Array.get(2);
//                    HeightFourCouple = Array.get(3);
//                    HeightFiveCouple = Array.get(4);
//                    HeightSixCouple = Array.get(5);
//                    HeightSevenCouple = Array.get(6);
                }

                if(ArrayClassroomCheckBox.get(i).getText().equals("416/2")) {
                    CreatorTable(ClassroomNumber,Array416);
//                    HeightOneCouple = Array.get(0);
//                    HeightTwoCouple = Array.get(1);
//                    HeightThreeCouple = Array.get(2);
//                    HeightFourCouple = Array.get(3);
//                    HeightFiveCouple = Array.get(4);
//                    HeightSixCouple = Array.get(5);
//                    HeightSevenCouple = Array.get(6);
                }

                if(ArrayClassroomCheckBox.get(i).getText().equals("417/2")) {
                    CreatorTable(ClassroomNumber,Array417);
//                    HeightOneCouple = Array.get(0);
//                    HeightTwoCouple = Array.get(1);
//                    HeightThreeCouple = Array.get(2);
//                    HeightFourCouple = Array.get(3);
//                    HeightFiveCouple = Array.get(4);
//                    HeightSixCouple = Array.get(5);
//                    HeightSevenCouple = Array.get(6);
                }

                switch(ArrayClassroomCheckBox.get(i).getText()){
                    case ("410/2"):{
                        Aud = Aud + "410 ";
                        break;
                    }
                    case("411/2"):{
                        Aud = Aud + "411 ";
                        break;
                    }
                    case("413/2"):{
                        Aud = Aud + "413 ";
                        break;
                    }
                    case ("416/2"):{
                        Aud = Aud + "416 ";
                        break;
                    }
                    case("417/2"):{
                        Aud = Aud + "417 ";
                        break;
                    }
                }

                ClassroomNumber++;
            }
        }

        String separator = File.separator;

        Date date = new Date();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
        File fileDate = new File(simpleDateFormat.format(date));

        FileOutputStream fileOutputStream = new FileOutputStream("TableClassroom(s)" + separator + "Few Classroom" + separator + "Аудитории " + Aud + "(" + fileDate + ")");

        workbookClassroom.write(fileOutputStream);
        fileOutputStream.close();

    }

}
