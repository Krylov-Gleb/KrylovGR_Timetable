package timetablekrylov.timetablekrylovgr;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class CreatorTableExelGroup {

    // Creating an Excel workbook.
    private Workbook workbookOneElementGroup = new HSSFWorkbook();

    // I am creating a sheet on which our tables will be located.
    private Sheet Group = workbookOneElementGroup.createSheet("Группа");

    // Creating a row located at index 0 (In table 1)
    private Row rowGroupZero = Group.createRow(0);

    // Creating a cell storing the day of the week (0x0)
    private Cell cellDayWeek = rowGroupZero.createCell(0);

    // Creating cells of the days of the week
    // Cell Monday (0x1)
    private Cell cellGroupMonday = rowGroupZero.createCell(1);

    // Cell Tuesday (0x2)
    private Cell cellGroupTuesday = rowGroupZero.createCell(2);

    // Cell Wednesday (0x3)
    private Cell cellGroupWednesday = rowGroupZero.createCell(3);

    // Cell Thursday (0x4)
    private Cell cellGroupThursday = rowGroupZero.createCell(4);

    // Cell Friday (0x5)
    private Cell cellGroupFriday = rowGroupZero.createCell(5);

    // Cell Saturday (0x6)
    private Cell cellGroupSaturday = rowGroupZero.createCell(6);

    // Creating a row located at index 1 (in table 2)
    private Row rowGroupOne = Group.createRow(1);

    // I am creating a cell in which the number of the pair is stored. (1x0).
    private Cell cellOneCouple = rowGroupOne.createCell(0);

    // Creating a row located at index 2 (in table 3)
    private Row rowGroupTwo = Group.createRow(2);

    // I am creating a cell in which the number of the pair is stored. (2x0)
    private Cell cellTwoCouple = rowGroupTwo.createCell(0);

    // Creating a row located at index 3 (in table 4)
    private Row rowGroupThree = Group.createRow(3);

    // I am creating a cell in which the number of the pair is stored. (3x0)
    private Cell cellThreeCouple = rowGroupThree.createCell(0);

    // Creating a row located at index 4 (in table 5)
    private Row rowGroupFour = Group.createRow(4);

    // I am creating a cell in which the number of the pair is stored. (4x0)
    private Cell cellFourCouple = rowGroupFour.createCell(0);

    // Creating a row located at index 5 (in table 6)
    private Row rowGroupFive = Group.createRow(5);

    // I am creating a cell in which the number of the pair is stored. (5x0)
    private Cell cellFiveCouple = rowGroupFive.createCell(0);

    // Creating a row located at index 6 (in table 7)
    private Row rowGroupSix = Group.createRow(6);

    // I am creating a cell in which the number of the pair is stored. (6x0)
    private Cell cellSixCouple = rowGroupSix.createCell(0);

    // Creating a row located at index 7 (in table 8)
    private Row rowGroupSeven = Group.createRow(7);

    // I am creating a cell in which the number of the pair is stored. (7x0)
    private Cell cellSevenCouple = rowGroupSeven.createCell(0);

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
                CoupleOneDayWeekOne = CoupleOneDayWeekOne + CoupleOne.get(i).GetDiscipline() + " (" + CoupleOne.get(i).GetCoupleType() + ")\n" + CoupleOne.get(i).GetNumberWeek() + " " + CoupleOne.get(i).GetTeacherName() + " " + CoupleOne.get(i).GetAud() + "\n" + "\n";
            }
            else{
                CoupleOneDayWeekOne = CoupleOneDayWeekOne + CoupleOne.get(i).GetDiscipline() + " (" + CoupleOne.get(i).GetCoupleType() + ")\n" + CoupleOne.get(i).GetNumberWeek() + " " + CoupleOne.get(i).GetTeacherName() + " " + CoupleOne.get(i).GetAud();
            }
        }

        for(int i = 0; i < CoupleTwo.size(); i++){
            if(i < CoupleTwo.size()-1) {
                CoupleTwoDayWeekTwo = CoupleTwoDayWeekTwo + CoupleTwo.get(i).GetDiscipline() + " (" + CoupleTwo.get(i).GetCoupleType() + ")\n" + CoupleTwo.get(i).GetNumberWeek() + " " + CoupleTwo.get(i).GetTeacherName() + " " + CoupleTwo.get(i).GetAud() + "\n" + "\n";
            }
            else{
                CoupleTwoDayWeekTwo = CoupleTwoDayWeekTwo + CoupleTwo.get(i).GetDiscipline() + " (" + CoupleTwo.get(i).GetCoupleType() + ")\n" + CoupleTwo.get(i).GetNumberWeek() + " " + CoupleTwo.get(i).GetTeacherName() + " " + CoupleTwo.get(i).GetAud();
            }
        }

        for(int i = 0; i < CoupleThree.size(); i++){
            if(i < CoupleThree.size()-1) {
                CoupleThreeDayWeekThree = CoupleThreeDayWeekThree + CoupleThree.get(i).GetDiscipline() + " (" + CoupleThree.get(i).GetCoupleType() + ")\n" + CoupleThree.get(i).GetNumberWeek() + " " + CoupleThree.get(i).GetTeacherName() + " " + CoupleThree.get(i).GetAud() + "\n" + "\n";
            }
            else{
                CoupleThreeDayWeekThree = CoupleThreeDayWeekThree + CoupleThree.get(i).GetDiscipline() + " (" + CoupleThree.get(i).GetCoupleType() + ")\n" + CoupleThree.get(i).GetNumberWeek() + " " + CoupleThree.get(i).GetTeacherName() + " " + CoupleThree.get(i).GetAud();
            }
        }

        for(int i = 0; i < CoupleFour.size(); i++){
            if(i < CoupleFour.size()-1) {
                CoupleFourDayWeekFour = CoupleFourDayWeekFour + CoupleFour.get(i).GetDiscipline() + " (" + CoupleFour.get(i).GetCoupleType() + ")\n" + CoupleFour.get(i).GetNumberWeek() + " " + CoupleFour.get(i).GetTeacherName() + " " + CoupleFour.get(i).GetAud() + "\n" + "\n";
            }
            else{
                CoupleFourDayWeekFour = CoupleFourDayWeekFour + CoupleFour.get(i).GetDiscipline() + " (" + CoupleFour.get(i).GetCoupleType() + ")\n" + CoupleFour.get(i).GetNumberWeek() + " " + CoupleFour.get(i).GetTeacherName() + " " + CoupleFour.get(i).GetAud();
            }
        }

        for(int i = 0; i < CoupleFive.size(); i++){
            if(i < CoupleFive.size()-1) {
                CoupleFiveDayWeekFive = CoupleFiveDayWeekFive + CoupleFive.get(i).GetDiscipline() + " (" + CoupleFive.get(i).GetCoupleType() + ")\n" + CoupleFive.get(i).GetNumberWeek() + " " + CoupleFive.get(i).GetTeacherName() + " " + CoupleFive.get(i).GetAud() + "\n" + "\n";
            }
            else{
                CoupleFiveDayWeekFive = CoupleFiveDayWeekFive + CoupleFive.get(i).GetDiscipline() + " (" + CoupleFive.get(i).GetCoupleType() + ")\n" + CoupleFive.get(i).GetNumberWeek() + " " + CoupleFive.get(i).GetTeacherName() + " " + CoupleFive.get(i).GetAud();
            }
        }

        for(int i = 0; i < CoupleSix.size(); i++){
            if(i < CoupleSix.size()-1) {
                CoupleSixDayWeekSix = CoupleSixDayWeekSix + CoupleSix.get(i).GetDiscipline() + " (" + CoupleSix.get(i).GetCoupleType() + ")\n" + CoupleSix.get(i).GetNumberWeek() + " " + CoupleSix.get(i).GetTeacherName() + " " + CoupleSix.get(i).GetAud() + "\n" + "\n";
            }
            else{
                CoupleSixDayWeekSix = CoupleSixDayWeekSix + CoupleSix.get(i).GetDiscipline() + " (" + CoupleSix.get(i).GetCoupleType() + ")\n" + CoupleSix.get(i).GetNumberWeek() + " " + CoupleSix.get(i).GetTeacherName() + " " + CoupleSix.get(i).GetAud();
            }
        }

        for(int i = 0; i < CoupleSeven.size(); i++){
            if(i < CoupleSeven.size()-1) {
                CoupleSevenDayWeekSeven = CoupleSevenDayWeekSeven + CoupleSeven.get(i).GetDiscipline() + " (" + CoupleSeven.get(i).GetCoupleType() + ")\n" + CoupleSeven.get(i).GetNumberWeek() + " " + CoupleSeven.get(i).GetTeacherName() + " " + CoupleSeven.get(i).GetAud() + "\n" + "\n";
            }
            else{
                CoupleSevenDayWeekSeven = CoupleSevenDayWeekSeven + CoupleSeven.get(i).GetDiscipline() + " (" + CoupleSeven.get(i).GetCoupleType() + ")\n" + CoupleSeven.get(i).GetNumberWeek() + " " + CoupleSeven.get(i).GetTeacherName() + " " + CoupleSeven.get(i).GetAud();
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

    // I'm creating a method to create a table with one group.
    // Passing the Group class to the method.
    public void CreatorTimeTableOneGroup(Group group) throws IOException {

        workbookOneElementGroup.setSheetName(0,group.GetGroupName());

        // I create variables responsible for the dimensions.
        int HeightPoints = 50;
        int ColumnWidth = 12000;

        ArrayList<CoupleGroup> CoupleOne = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleTwo = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleThree = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleFour = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleFive = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleSix = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleSeven = new ArrayList<>();

        // I give permission to move a line in cells.
        CellStyle cellStyle = workbookOneElementGroup.createCellStyle();
        cellStyle.setWrapText(true);

        // I am creating an array in which Monday pairs will be stored.
        ArrayList<CoupleGroup> ArrayMonday = new ArrayList<>();

        // I am creating an array in which Tuesday pairs will be stored.
        ArrayList<CoupleGroup> ArrayTuesday = new ArrayList<>();

        // I am creating an array in which Wednesday pairs will be stored.
        ArrayList<CoupleGroup> ArrayWednesday = new ArrayList<>();

        // I am creating an array in which Thursday pairs will be stored.
        ArrayList<CoupleGroup> ArrayThursday = new ArrayList<>();

        // I am creating an array in which Friday pairs will be stored.
        ArrayList<CoupleGroup> ArrayFriday = new ArrayList<>();

        // I am creating an array in which Saturday pairs will be stored.
        ArrayList<CoupleGroup> ArraySaturday = new ArrayList<>();

        // Creating an array to store all pairs.
        ArrayList<CoupleGroup> ArrayCouple = group.GetArrayCouples();

        // Setting values for the cell (cellDayWeek).
        cellDayWeek.setCellValue("День недели" + "\n" + "Номер пары");

        // I give permission for line wrapping (Changing my style).
        cellDayWeek.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(0,8000);
        rowGroupZero.setHeightInPoints(30);

        // Setting values for the cell (cellOneCouple).
        cellOneCouple.setCellValue("1 пара");

        // Setting values for the cell (cellTwoCouple).
        cellTwoCouple.setCellValue("2 пара");

        // Setting values for the cell (cellThreeCouple).
        cellThreeCouple.setCellValue("3 пара");

        // Setting values for the cell (cellFourCouple).
        cellFourCouple.setCellValue("4 пара");

        // Setting values for the cell (cellFiveCouple).
        cellFiveCouple.setCellValue("5 пара");

        // Setting values for the cell (cellSixCouple).
        cellSixCouple.setCellValue("6 пара");

        // Setting values for the cell (cellSevenCouple).
        cellSevenCouple.setCellValue("7 пара");

        // Setting values for the cell (cellGroupMonday).
        cellGroupMonday.setCellValue("Понедельник");

        // I set the dimensions of the cell.
        Group.setColumnWidth(1,8000);

        // Setting values for the cell (cellGroupTuesday).
        cellGroupTuesday.setCellValue("Вторник");

        // I set the dimensions of the cell.
        Group.setColumnWidth(2,8000);

        // Setting values for the cell (cellGroupWednesday).
        cellGroupWednesday.setCellValue("Среда");

        // I set the dimensions of the cell.
        Group.setColumnWidth(3,8000);

        // Setting values for the cell (cellGroupThursday).
        cellGroupThursday.setCellValue("Четверг");

        // I set the dimensions of the cell.
        Group.setColumnWidth(4,8000);

        // Setting values for the cell (cellGroupFriday).
        cellGroupFriday.setCellValue("Пятница");

        // I set the dimensions of the cell.
        Group.setColumnWidth(5,8000);

        // Setting values for the cell (cellGroupSaturday).
        cellGroupSaturday.setCellValue("Суббота");

        // I set the dimensions of the cell.
        Group.setColumnWidth(6,8000);

        // I'm going through the array of all pairs.
        for(int i = 0; i < ArrayCouple.size(); i++){

            // I save the day of the week of each pair to the idDay variable (in turn).
            int IdDay = ArrayCouple.get(i).GetIDDay();

            // I sort pairs by days of the week by writing them into arrays.
            switch (IdDay){
                case (1):{
                    ArrayMonday.add(ArrayCouple.get(i));
                    break;
                }
                case (2):{
                    ArrayTuesday.add(ArrayCouple.get(i));
                    break;
                }
                case (3):{
                    ArrayWednesday.add(ArrayCouple.get(i));
                    break;
                }
                case (4):{
                    ArrayThursday.add(ArrayCouple.get(i));
                    break;
                }
                case (5):{
                    ArrayFriday.add(ArrayCouple.get(i));
                    break;
                }
                case (6):{
                    ArraySaturday.add(ArrayCouple.get(i));
                    break;
                }
            }
        }

        // Couples Monday

        // A variable for concatenating the first pairs on Monday.
        String CoupleOneMonday = "";

        // A variable for combining the second pairs on Monday
        String CoupleTwoMonday = "";

        // A variable for combining the third pairs on Monday.
        String CoupleThreeMonday = "";

        // A variable for concatenating the fourth pairs on Monday.
        String CoupleFourMonday = "";

        // A variable for combining fifth pairs on Monday.
        String CoupleFiveMonday = "";

        // A variable for concatenating the sixth pairs on Monday.
        String CoupleSixMonday = "";

        // A variable for combining the seventh pairs on Monday.
        String CoupleSevenMonday = "";

        // Going through the array of Monday pairs.
        for(int i = 0; i < ArrayMonday.size(); i++){

            // We are looking at the pair number of the elements from the array.
            switch (ArrayMonday.get(i).GetCoupleNumber()) {
                case (1): {
                    // We combine the first pairs for use in cells.
                    CoupleOne.add(ArrayMonday.get(i));
                    break;
                }
                case (2): {
                    // We combine the second pairs for use in cells.
                    CoupleTwo.add(ArrayMonday.get(i));
                    break;
                }
                case (3): {
                    // We combine the third pairs for use in cells.
                    CoupleThree.add(ArrayMonday.get(i));
                    break;
                }
                case (4): {
                    // We combine the fourth pairs for use in cells.
                    CoupleFour.add(ArrayMonday.get(i));
                    break;
                }
                case (5): {
                    // We combine the fifth pairs for use in cells.
                    CoupleFive.add(ArrayMonday.get(i));
                    break;
                }
                case (6): {
                    // We combine the sixth pairs for use in cells.
                    CoupleSix.add(ArrayMonday.get(i));
                    break;
                }
                case (7): {
                    // We combine the seventh pairs for use in cells.
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


        // I am creating a cell (the first pair is on Monday). (1x1)
        Cell cellOneCoupleMonday = rowGroupOne.createCell(1);

        // Setting the value
        cellOneCoupleMonday.setCellValue(CoupleOneMonday);

        // I give permission to move a line in cells.
        cellOneCoupleMonday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(1,ColumnWidth);
        rowGroupOne.setHeightInPoints(HeightPoints);


        // I'm creating a cell (the second pair is on Monday). (2x1)
        Cell cellTwoCoupleMonday = rowGroupTwo.createCell(1);

        // Setting the value
        cellTwoCoupleMonday.setCellValue(CoupleTwoMonday);

        // I give permission to move a line in cells.
        cellTwoCoupleMonday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(1,ColumnWidth);
        rowGroupTwo.setHeightInPoints(HeightPoints);

        // I'm creating a cell (the third pair is on Monday). (3x1)
        Cell cellThreeCoupleMonday = rowGroupThree.createCell(1);

        // Setting the value
        cellThreeCoupleMonday.setCellValue(CoupleThreeMonday);

        // I give permission to move a line in cells.
        cellThreeCoupleMonday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(1,ColumnWidth);
        rowGroupThree.setHeightInPoints(HeightPoints);

        // I am creating a cell (the fourth pair is on Monday). (4x1)
        Cell cellFourCoupleMonday = rowGroupFour.createCell(1);

        // Setting the value
        cellFourCoupleMonday.setCellValue(CoupleFourMonday);

        // I give permission to move a line in cells.
        cellFourCoupleMonday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(1,ColumnWidth);
        rowGroupFour.setHeightInPoints(HeightPoints);

        // I am creating a cell (the fifth pair is on Monday). (5x1)
        Cell cellFiveCoupleMonday = rowGroupFive.createCell(1);

        // Setting the value
        cellFiveCoupleMonday.setCellValue(CoupleFiveMonday);

        // I give permission to move a line in cells.
        cellFiveCoupleMonday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(1,ColumnWidth);
        rowGroupFive.setHeightInPoints(HeightPoints);

        // I'm creating a cell (the sixth pair is on Monday). (6x1)
        Cell cellSixCoupleMonday = rowGroupSix.createCell(1);

        // Setting the value
        cellSixCoupleMonday.setCellValue(CoupleSixMonday);

        // I give permission to move a line in cells.
        cellSixCoupleMonday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(1,ColumnWidth);
        rowGroupSix.setHeightInPoints(HeightPoints);

        // I am creating a cell (the seventh pair is on Monday). (7x1)
        Cell cellSevenCoupleMonday = rowGroupSeven.createCell(1);

        // Setting the value
        cellSevenCoupleMonday.setCellValue(CoupleSevenMonday);

        // I give permission to move a line in cells.
        cellSevenCoupleMonday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(1,ColumnWidth);
        rowGroupSeven.setHeightInPoints(HeightPoints);

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

        // A variable for concatenating the first pairs on Tuesday.
        String CoupleOneTuesday = "";

        // A variable for combining the second pairs on Tuesday
        String CoupleTwoTuesday = "";

        // A variable for combining the third pairs on Tuesday.
        String CoupleThreeTuesday = "";

        // A variable for concatenating the fourth pairs on Tuesday.
        String CoupleFourTuesday = "";

        // A variable for combining fifth pairs on Tuesday.
        String CoupleFiveTuesday = "";

        // A variable for concatenating the sixth pairs on Tuesday.
        String CoupleSixTuesday = "";

        // A variable for combining the seventh pairs on Tuesday.
        String CoupleSevenTuesday = "";

        // Going through the array of Tuesday pairs.
        for(int i = 0; i < ArrayTuesday.size(); i++){

            // We are looking at the pair number of the elements from the array.
            switch (ArrayTuesday.get(i).GetCoupleNumber()) {
                case (1): {
                    // We combine the first pairs for use in cells.
                    CoupleOne.add(ArrayTuesday.get(i));
                    break;
                }
                case (2): {
                    // We combine the second pairs for use in cells.
                    CoupleTwo.add(ArrayTuesday.get(i));
                    break;
                }
                case (3): {
                    // We combine the third pairs for use in cells.
                    CoupleThree.add(ArrayTuesday.get(i));
                    break;
                }
                case (4): {
                    // We combine the fourth pairs for use in cells.
                    CoupleFour.add(ArrayTuesday.get(i));
                    break;
                }
                case (5): {
                    // We combine the fifth pairs for use in cells.
                    CoupleFive.add(ArrayTuesday.get(i));
                    break;
                }
                case (6): {
                    // We combine the sixth pairs for use in cells.
                    CoupleSix.add(ArrayTuesday.get(i));
                    break;
                }
                case (7): {
                    // We combine the seventh pairs for use in cells.
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

        // I am creating a cell (the first pair is on Tuesday). (1x2)
        Cell cellOneCoupleTuesday = rowGroupOne.createCell(2);

        // Setting the value
        cellOneCoupleTuesday.setCellValue(CoupleOneTuesday);

        // I give permission to move a line in cells.
        cellOneCoupleTuesday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(2,ColumnWidth);

//        if(MaxHeightOneCouple < HeightOneCouple){
//            MaxHeightOneCouple = HeightOneCouple;
//        }

        rowGroupOne.setHeightInPoints(HeightPoints);

        // I'm creating a cell (the second pair is on Tuesday). (2x2)
        Cell cellTwoCoupleTuesday = rowGroupTwo.createCell(2);

        // Setting the value
        cellTwoCoupleTuesday.setCellValue(CoupleTwoTuesday);

        // I give permission to move a line in cells.
        cellTwoCoupleTuesday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(2,ColumnWidth);

//        if(MaxHeightTwoCouple < HeightTwoCouple){
//            MaxHeightTwoCouple = HeightTwoCouple;
//        }

        rowGroupTwo.setHeightInPoints(HeightPoints);

        // I'm creating a cell (the third pair is on Tuesday). (3x2)
        Cell cellThreeCoupleTuesday = rowGroupThree.createCell(2);

        // Setting the value
        cellThreeCoupleTuesday.setCellValue(CoupleThreeTuesday);

        // I give permission to move a line in cells.
        cellThreeCoupleTuesday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(2,ColumnWidth);

//        if(MaxHeightThreeCouple < HeightThreeCouple){
//            MaxHeightThreeCouple = HeightThreeCouple;
//        }

        rowGroupThree.setHeightInPoints(HeightPoints);

        // I am creating a cell (the fourth pair is on Tuesday). (4x2)
        Cell cellFourCoupleTuesday = rowGroupFour.createCell(2);

        // Setting the value
        cellFourCoupleTuesday.setCellValue(CoupleFourTuesday);

        // I give permission to move a line in cells.
        cellFourCoupleTuesday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(2,ColumnWidth);
//
//        if(MaxHeightFourCouple < HeightFourCouple){
//            MaxHeightFourCouple = HeightFourCouple;
//        }

        rowGroupFour.setHeightInPoints(HeightPoints);

        // I am creating a cell (the fifth pair is on Tuesday). (5x2)
        Cell cellFiveCoupleTuesday = rowGroupFive.createCell(2);

        // Setting the value
        cellFiveCoupleTuesday.setCellValue(CoupleFiveTuesday);

        // I give permission to move a line in cells.
        cellFiveCoupleTuesday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(2,ColumnWidth);

//        if(MaxHeightFiveCouple < HeightFiveCouple){
//            MaxHeightFiveCouple = HeightFiveCouple;
//        }

        rowGroupFive.setHeightInPoints(HeightPoints);

        // I'm creating a cell (the sixth pair is on Tuesday). (6x2)
        Cell cellSixCoupleTuesday = rowGroupSix.createCell(2);

        // Setting the value
        cellSixCoupleTuesday.setCellValue(CoupleSixTuesday);
        // I give permission to move a line in cells.
        cellSixCoupleTuesday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(2,ColumnWidth);
//
//        if(MaxHeightSixCouple < HeightSixCouple){
//            MaxHeightSixCouple = HeightSixCouple;
//        }

        rowGroupSix.setHeightInPoints(HeightPoints);

        // I am creating a cell (the seventh pair is on Tuesday). (7x2)
        Cell cellSevenCoupleTuesday = rowGroupSeven.createCell(2);

        // Setting the value
        cellSevenCoupleTuesday.setCellValue(CoupleSevenTuesday);

        // I give permission to move a line in cells.
        cellSevenCoupleTuesday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(2,ColumnWidth);

//        if(MaxHeightSevenCouple < HeightSevenCouple){
//            MaxHeightSevenCouple = HeightSevenCouple;
//        }

        rowGroupSeven.setHeightInPoints(HeightPoints);

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

        Cell cellOneCoupleWednesday = rowGroupOne.createCell(3);
        cellOneCoupleWednesday.setCellValue(CoupleOneWednesday);
        cellOneCoupleWednesday.setCellStyle(cellStyle);
        Group.setColumnWidth(3,ColumnWidth);

//        if(MaxHeightOneCouple < HeightOneCouple){
//            MaxHeightOneCouple = HeightOneCouple;
//        }

        rowGroupOne.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleWednesday = rowGroupTwo.createCell(3);
        cellTwoCoupleWednesday.setCellValue(CoupleTwoWednesday);
        cellTwoCoupleWednesday.setCellStyle(cellStyle);
        Group.setColumnWidth(3,ColumnWidth);

//        if(MaxHeightTwoCouple < HeightTwoCouple){
//            MaxHeightTwoCouple = HeightTwoCouple;
//        }

        rowGroupTwo.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleWednesday = rowGroupThree.createCell(3);
        cellThreeCoupleWednesday.setCellValue(CoupleThreeWednesday);
        cellThreeCoupleWednesday.setCellStyle(cellStyle);
        Group.setColumnWidth(3,ColumnWidth);

//        if(MaxHeightThreeCouple < HeightTwoCouple){
//            MaxHeightThreeCouple = HeightTwoCouple;
//        }

        rowGroupThree.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleWednesday = rowGroupFour.createCell(3);
        cellFourCoupleWednesday.setCellValue(CoupleFourWednesday);
        cellFourCoupleWednesday.setCellStyle(cellStyle);
        Group.setColumnWidth(3,ColumnWidth);

//        if(MaxHeightFourCouple < HeightFourCouple){
//            MaxHeightFourCouple = HeightFourCouple;
//        }

        rowGroupFour.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleWednesday = rowGroupFive.createCell(3);
        cellFiveCoupleWednesday.setCellValue(CoupleFiveWednesday);
        cellFiveCoupleWednesday.setCellStyle(cellStyle);
        Group.setColumnWidth(3,ColumnWidth);

//        if(MaxHeightFiveCouple < HeightFiveCouple){
//            MaxHeightFiveCouple = HeightFiveCouple;
//        }

        rowGroupFive.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleWednesday = rowGroupSix.createCell(3);
        cellSixCoupleWednesday.setCellValue(CoupleSixWednesday);
        cellSixCoupleWednesday.setCellStyle(cellStyle);
        Group.setColumnWidth(3,ColumnWidth);

//        if(MaxHeightSixCouple < HeightSixCouple){
//            MaxHeightSixCouple = HeightSixCouple;
//        }

        rowGroupSix.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleWednesday = rowGroupSeven.createCell(3);
        cellSevenCoupleWednesday.setCellValue(CoupleSevenWednesday);
        cellSevenCoupleWednesday.setCellStyle(cellStyle);
        Group.setColumnWidth(3,ColumnWidth);

//        if(MaxHeightSevenCouple < HeightSevenCouple){
//            MaxHeightSevenCouple = HeightSevenCouple;
//        }

        rowGroupSeven.setHeightInPoints(HeightPoints);

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

        Cell cellOneCoupleThursday = rowGroupOne.createCell(4);
        cellOneCoupleThursday.setCellValue(CoupleOneThursday);
        cellOneCoupleThursday.setCellStyle(cellStyle);
        Group.setColumnWidth(4,ColumnWidth);

//        if(MaxHeightOneCouple < HeightOneCouple){
//            MaxHeightOneCouple = HeightOneCouple;
//        }

        rowGroupOne.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleThursday = rowGroupTwo.createCell(4);
        cellTwoCoupleThursday.setCellValue(CoupleTwoThursday);
        cellTwoCoupleThursday.setCellStyle(cellStyle);
        Group.setColumnWidth(4,ColumnWidth);

//        if(MaxHeightTwoCouple < HeightTwoCouple){
//            MaxHeightTwoCouple = HeightTwoCouple;
//        }

        rowGroupTwo.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleThursday = rowGroupThree.createCell(4);
        cellThreeCoupleThursday.setCellValue(CoupleThreeThursday);
        cellThreeCoupleThursday.setCellStyle(cellStyle);
        Group.setColumnWidth(4,ColumnWidth);

//        if(MaxHeightThreeCouple < HeightThreeCouple){
//            MaxHeightThreeCouple = HeightThreeCouple;
//        }

        rowGroupThree.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleThursday = rowGroupFour.createCell(4);
        cellFourCoupleThursday.setCellValue(CoupleFourThursday);
        cellFourCoupleThursday.setCellStyle(cellStyle);
        Group.setColumnWidth(4,ColumnWidth);

//        if(MaxHeightFourCouple < HeightFourCouple){
//            MaxHeightFourCouple = HeightFourCouple;
//        }

        rowGroupFour.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleThursday = rowGroupFive.createCell(4);
        cellFiveCoupleThursday.setCellValue(CoupleFiveThursday);
        cellFiveCoupleThursday.setCellStyle(cellStyle);
        Group.setColumnWidth(4,ColumnWidth);

//        if(MaxHeightFiveCouple < HeightFiveCouple){
//            MaxHeightFiveCouple = HeightFiveCouple;
//        }

        rowGroupFive.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleThursday = rowGroupSix.createCell(4);
        cellSixCoupleThursday.setCellValue(CoupleSixThursday);
        cellSixCoupleThursday.setCellStyle(cellStyle);
        Group.setColumnWidth(4,ColumnWidth);

//        if(MaxHeightSixCouple < HeightSixCouple){
//            MaxHeightSixCouple = HeightSixCouple;
//        }

        rowGroupSix.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleThursday = rowGroupSeven.createCell(4);
        cellSevenCoupleThursday.setCellValue(CoupleSevenThursday);
        cellSevenCoupleThursday.setCellStyle(cellStyle);
        Group.setColumnWidth(4,ColumnWidth);

//        if(MaxHeightSevenCouple < HeightSevenCouple){
//            MaxHeightSevenCouple = HeightSevenCouple;
//        }

        rowGroupSeven.setHeightInPoints(HeightPoints);

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

        Cell cellOneCoupleFriday = rowGroupOne.createCell(5);
        cellOneCoupleFriday.setCellValue(CoupleOneFriday);
        cellOneCoupleFriday.setCellStyle(cellStyle);
        Group.setColumnWidth(5,ColumnWidth);

//        if(MaxHeightOneCouple < HeightOneCouple){
//            MaxHeightOneCouple = HeightOneCouple;
//        }

        rowGroupOne.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleFriday = rowGroupTwo.createCell(5);
        cellTwoCoupleFriday.setCellValue(CoupleTwoFriday);
        cellTwoCoupleFriday.setCellStyle(cellStyle);
        Group.setColumnWidth(5,ColumnWidth);

//        if(MaxHeightTwoCouple < HeightTwoCouple){
//            MaxHeightTwoCouple = HeightTwoCouple;
//        }

        rowGroupTwo.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleFriday = rowGroupThree.createCell(5);
        cellThreeCoupleFriday.setCellValue(CoupleThreeFriday);
        cellThreeCoupleFriday.setCellStyle(cellStyle);
        Group.setColumnWidth(5,ColumnWidth);

//        if(MaxHeightThreeCouple < HeightThreeCouple){
//            MaxHeightThreeCouple = HeightThreeCouple;
//        }

        rowGroupThree.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleFriday = rowGroupFour.createCell(5);
        cellFourCoupleFriday.setCellValue(CoupleFourFriday);
        cellFourCoupleFriday.setCellStyle(cellStyle);
        Group.setColumnWidth(5,ColumnWidth);

//        if(MaxHeightFourCouple < HeightFourCouple){
//            MaxHeightFourCouple = HeightFourCouple;
//        }

        rowGroupFour.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleFriday = rowGroupFive.createCell(5);
        cellFiveCoupleFriday.setCellValue(CoupleFiveFriday);
        cellFiveCoupleFriday.setCellStyle(cellStyle);
        Group.setColumnWidth(5,ColumnWidth);

//        if(MaxHeightFiveCouple < HeightFiveCouple){
//            MaxHeightFiveCouple = HeightFiveCouple;
//        }

        rowGroupFive.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleFriday = rowGroupSix.createCell(5);
        cellSixCoupleFriday.setCellValue(CoupleSixFriday);
        cellSixCoupleFriday.setCellStyle(cellStyle);
        Group.setColumnWidth(5,ColumnWidth);

//        if(MaxHeightSixCouple < HeightSixCouple){
//            MaxHeightSixCouple = HeightSixCouple;
//        }

        rowGroupSix.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleFriday = rowGroupSeven.createCell(5);
        cellSevenCoupleFriday.setCellValue(CoupleSevenFriday);
        cellSevenCoupleFriday.setCellStyle(cellStyle);
        Group.setColumnWidth(5,ColumnWidth);

//        if(MaxHeightSevenCouple < HeightSevenCouple){
//            MaxHeightSevenCouple = HeightSevenCouple;
//        }

        rowGroupSeven.setHeightInPoints(HeightPoints);

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

        Cell cellOneCoupleSaturday = rowGroupOne.createCell(6);
        cellOneCoupleSaturday.setCellValue(CoupleOneSaturday);
        cellOneCoupleSaturday.setCellStyle(cellStyle);
        Group.setColumnWidth(6,ColumnWidth);

//        if(MaxHeightOneCouple < HeightOneCouple){
//            MaxHeightOneCouple = HeightOneCouple;
//        }

        rowGroupOne.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleSaturday = rowGroupTwo.createCell(6);
        cellTwoCoupleSaturday.setCellValue(CoupleTwoSaturday);
        cellTwoCoupleSaturday.setCellStyle(cellStyle);
        Group.setColumnWidth(6,ColumnWidth);

//        if(MaxHeightTwoCouple < HeightTwoCouple){
//            MaxHeightTwoCouple = HeightTwoCouple;
//        }

        rowGroupTwo.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleSaturday = rowGroupThree.createCell(6);
        cellThreeCoupleSaturday.setCellValue(CoupleThreeSaturday);
        cellThreeCoupleSaturday.setCellStyle(cellStyle);
        Group.setColumnWidth(6,ColumnWidth);

//        if(MaxHeightThreeCouple < HeightThreeCouple){
//            MaxHeightThreeCouple = HeightThreeCouple;
//        }

        rowGroupThree.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleSaturday = rowGroupFour.createCell(6);
        cellFourCoupleSaturday.setCellValue(CoupleFourSaturday);
        cellFourCoupleSaturday.setCellStyle(cellStyle);
        Group.setColumnWidth(6,ColumnWidth);

//        if(MaxHeightFourCouple < HeightFourCouple){
//            MaxHeightFourCouple = HeightFourCouple;
//        }

        rowGroupFour.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleSaturday = rowGroupFive.createCell(6);
        cellFiveCoupleSaturday.setCellValue(CoupleFiveSaturday);
        cellFiveCoupleSaturday.setCellStyle(cellStyle);
        Group.setColumnWidth(6,ColumnWidth);

//        if(MaxHeightFiveCouple < HeightFiveCouple){
//            MaxHeightFiveCouple = HeightFiveCouple;
//        }

        rowGroupFive.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleSaturday = rowGroupSix.createCell(6);
        cellSixCoupleSaturday.setCellValue(CoupleSixSaturday);
        cellSixCoupleSaturday.setCellStyle(cellStyle);
        Group.setColumnWidth(6,ColumnWidth);

//        if(MaxHeightSixCouple < HeightSixCouple){
//            MaxHeightSixCouple = HeightSixCouple;
//        }

        rowGroupSix.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleSaturday = rowGroupSeven.createCell(6);
        cellSevenCoupleSaturday.setCellValue(CoupleSevenSaturday);
        cellSevenCoupleSaturday.setCellStyle(cellStyle);
        Group.setColumnWidth(6,ColumnWidth);

//        if(MaxHeightSevenCouple < HeightSevenCouple){
//            MaxHeightSevenCouple = HeightSevenCouple;
//        }

        rowGroupSeven.setHeightInPoints(HeightPoints);

        CoupleOne.clear();
        CoupleTwo.clear();
        CoupleThree.clear();
        CoupleFour.clear();
        CoupleFive.clear();
        CoupleSix.clear();
        CoupleSeven.clear();

        String separator = File.separator;

        // I write the result to a file named OneGroupExelDoc.
        FileOutputStream fileOutputStream = new FileOutputStream("TableGroup(s)" + separator + "OneGroupExelDoc");

        // I'm writing it down through a workbook.
        workbookOneElementGroup.write(fileOutputStream);

        // Closing the recording stream.
        fileOutputStream.close();
    }

}
