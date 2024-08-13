package timetablekrylov.timetablekrylovgr;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;

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

    // I'm creating a method to create a table with one group.
    // Passing the Group class to the method.
    public void CreatorTimeTableOneGroup(Group group) throws IOException {

        // I create variables responsible for the dimensions.
        int HeightPoints = 250;
        int ColumnWidth = 10000;

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
                    CoupleOneMonday = CoupleOneMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetCoupleType() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetTeacherName() + " " + ArrayMonday.get(i).GetAud() + " (" +  ArrayMonday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (2): {
                    // We combine the second pairs for use in cells.
                    CoupleTwoMonday = CoupleTwoMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetCoupleType() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetTeacherName() + " " + ArrayMonday.get(i).GetAud() + " (" +  ArrayMonday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (3): {
                    // We combine the third pairs for use in cells.
                    CoupleThreeMonday = CoupleThreeMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetCoupleType() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetTeacherName() + " " + ArrayMonday.get(i).GetAud() + " (" +  ArrayMonday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (4): {
                    // We combine the fourth pairs for use in cells.
                    CoupleFourMonday = CoupleFourMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetCoupleType() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetTeacherName() + " " + ArrayMonday.get(i).GetAud() + " (" +  ArrayMonday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (5): {
                    // We combine the fifth pairs for use in cells.
                    CoupleFiveMonday = CoupleFiveMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetCoupleType() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetTeacherName() + " " + ArrayMonday.get(i).GetAud() + " (" +  ArrayMonday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (6): {
                    // We combine the sixth pairs for use in cells.
                    CoupleSixMonday = CoupleSixMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetCoupleType() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetTeacherName() + " " + ArrayMonday.get(i).GetAud() + " (" +  ArrayMonday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (7): {
                    // We combine the seventh pairs for use in cells.
                    CoupleSevenMonday = CoupleSevenMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetCoupleType() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetTeacherName() + " " + ArrayMonday.get(i).GetAud() + " (" +  ArrayMonday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
            }
        }


        // I am creating a cell (the first pair is on Monday). (1x1)
        Cell cellOneCoupleMonday = rowGroupOne.createCell(1);

        // Setting the value
        cellOneCoupleMonday.setCellValue(CoupleOneMonday);

        // I give permission to move a line in cells.
        cellOneCoupleMonday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(1, ColumnWidth);
        rowGroupOne.setHeightInPoints(HeightPoints);


        // I'm creating a cell (the second pair is on Monday). (2x1)
        Cell cellTwoCoupleMonday = rowGroupTwo.createCell(1);

        // Setting the value
        cellTwoCoupleMonday.setCellValue(CoupleTwoMonday);

        // I give permission to move a line in cells.
        cellTwoCoupleMonday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(1, ColumnWidth);
        rowGroupTwo.setHeightInPoints(HeightPoints);

        // I'm creating a cell (the third pair is on Monday). (3x1)
        Cell cellThreeCoupleMonday = rowGroupThree.createCell(1);

        // Setting the value
        cellThreeCoupleMonday.setCellValue(CoupleThreeMonday);

        // I give permission to move a line in cells.
        cellThreeCoupleMonday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(1, ColumnWidth);
        rowGroupThree.setHeightInPoints(HeightPoints);

        // I am creating a cell (the fourth pair is on Monday). (4x1)
        Cell cellFourCoupleMonday = rowGroupFour.createCell(1);

        // Setting the value
        cellFourCoupleMonday.setCellValue(CoupleFourMonday);

        // I give permission to move a line in cells.
        cellFourCoupleMonday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(1, ColumnWidth);
        rowGroupFour.setHeightInPoints(HeightPoints);

        // I am creating a cell (the fifth pair is on Monday). (5x1)
        Cell cellFiveCoupleMonday = rowGroupFive.createCell(1);

        // Setting the value
        cellFiveCoupleMonday.setCellValue(CoupleFiveMonday);

        // I give permission to move a line in cells.
        cellFiveCoupleMonday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(1, ColumnWidth);
        rowGroupFive.setHeightInPoints(HeightPoints);

        // I'm creating a cell (the sixth pair is on Monday). (6x1)
        Cell cellSixCoupleMonday = rowGroupSix.createCell(1);

        // Setting the value
        cellSixCoupleMonday.setCellValue(CoupleSixMonday);

        // I give permission to move a line in cells.
        cellSixCoupleMonday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(1, ColumnWidth);
        rowGroupSix.setHeightInPoints(HeightPoints);

        // I am creating a cell (the seventh pair is on Monday). (7x1)
        Cell cellSevenCoupleMonday = rowGroupSeven.createCell(1);

        // Setting the value
        cellSevenCoupleMonday.setCellValue(CoupleSevenMonday);

        // I give permission to move a line in cells.
        cellSevenCoupleMonday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(1, ColumnWidth);
        rowGroupSeven.setHeightInPoints(HeightPoints);

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
                    CoupleOneTuesday = CoupleOneTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetCoupleType() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetTeacherName() + " " + ArrayTuesday.get(i).GetAud() + " (" +  ArrayTuesday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (2): {
                    // We combine the second pairs for use in cells.
                    CoupleTwoTuesday = CoupleTwoTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetCoupleType() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetTeacherName() + " " + ArrayTuesday.get(i).GetAud() + " (" +  ArrayTuesday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (3): {
                    // We combine the third pairs for use in cells.
                    CoupleThreeTuesday = CoupleThreeTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetCoupleType() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetTeacherName() + " " + ArrayTuesday.get(i).GetAud() + " (" +  ArrayTuesday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (4): {
                    // We combine the fourth pairs for use in cells.
                    CoupleFourTuesday = CoupleFourTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetCoupleType() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetTeacherName() + " " + ArrayTuesday.get(i).GetAud() + " (" +  ArrayTuesday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (5): {
                    // We combine the fifth pairs for use in cells.
                    CoupleFiveTuesday = CoupleFiveTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetCoupleType() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetTeacherName() + " " + ArrayTuesday.get(i).GetAud() + " (" +  ArrayTuesday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (6): {
                    // We combine the sixth pairs for use in cells.
                    CoupleSixTuesday = CoupleSixTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetCoupleType() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetTeacherName() + " " + ArrayTuesday.get(i).GetAud() + " (" +  ArrayTuesday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (7): {
                    // We combine the seventh pairs for use in cells.
                    CoupleSevenTuesday = CoupleSevenTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetCoupleType() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetTeacherName() + " " + ArrayTuesday.get(i).GetAud() + " (" +  ArrayTuesday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
            }
        }

        // I am creating a cell (the first pair is on Tuesday). (1x2)
        Cell cellOneCoupleTuesday = rowGroupOne.createCell(2);

        // Setting the value
        cellOneCoupleTuesday.setCellValue(CoupleOneTuesday);

        // I give permission to move a line in cells.
        cellOneCoupleTuesday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(2, ColumnWidth);
        rowGroupOne.setHeightInPoints(HeightPoints);

        // I'm creating a cell (the second pair is on Tuesday). (2x2)
        Cell cellTwoCoupleTuesday = rowGroupTwo.createCell(2);

        // Setting the value
        cellTwoCoupleTuesday.setCellValue(CoupleTwoTuesday);

        // I give permission to move a line in cells.
        cellTwoCoupleTuesday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(2, ColumnWidth);
        rowGroupTwo.setHeightInPoints(HeightPoints);

        // I'm creating a cell (the third pair is on Tuesday). (3x2)
        Cell cellThreeCoupleTuesday = rowGroupThree.createCell(2);

        // Setting the value
        cellThreeCoupleTuesday.setCellValue(CoupleThreeTuesday);

        // I give permission to move a line in cells.
        cellThreeCoupleTuesday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(2, ColumnWidth);
        rowGroupThree.setHeightInPoints(HeightPoints);

        // I am creating a cell (the fourth pair is on Tuesday). (4x2)
        Cell cellFourCoupleTuesday = rowGroupFour.createCell(2);

        // Setting the value
        cellFourCoupleTuesday.setCellValue(CoupleFourTuesday);

        // I give permission to move a line in cells.
        cellFourCoupleTuesday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(2, ColumnWidth);
        rowGroupFour.setHeightInPoints(HeightPoints);

        // I am creating a cell (the fifth pair is on Tuesday). (5x2)
        Cell cellFiveCoupleTuesday = rowGroupFive.createCell(2);

        // Setting the value
        cellFiveCoupleTuesday.setCellValue(CoupleFiveTuesday);

        // I give permission to move a line in cells.
        cellFiveCoupleTuesday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(2, ColumnWidth);
        rowGroupFive.setHeightInPoints(HeightPoints);

        // I'm creating a cell (the sixth pair is on Tuesday). (6x2)
        Cell cellSixCoupleTuesday = rowGroupSix.createCell(2);

        // Setting the value
        cellSixCoupleTuesday.setCellValue(CoupleSixTuesday);

        // I give permission to move a line in cells.
        cellSixCoupleTuesday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(2, ColumnWidth);
        rowGroupSix.setHeightInPoints(HeightPoints);

        // I am creating a cell (the seventh pair is on Tuesday). (7x2)
        Cell cellSevenCoupleTuesday = rowGroupSeven.createCell(2);

        // Setting the value
        cellSevenCoupleTuesday.setCellValue(CoupleSevenTuesday);

        // I give permission to move a line in cells.
        cellSevenCoupleTuesday.setCellStyle(cellStyle);

        // I set the dimensions of the cell.
        Group.setColumnWidth(2, ColumnWidth);
        rowGroupSeven.setHeightInPoints(HeightPoints);

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
                    CoupleOneWednesday = CoupleOneWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetCoupleType() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetTeacherName() + " " + ArrayWednesday.get(i).GetAud() + " (" +  ArrayWednesday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (2): {
                    CoupleTwoWednesday = CoupleTwoWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetCoupleType() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetTeacherName() + " " + ArrayWednesday.get(i).GetAud() + " (" +  ArrayWednesday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (3): {
                    CoupleThreeWednesday = CoupleThreeWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetCoupleType() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetTeacherName() + " " + ArrayWednesday.get(i).GetAud() + " (" +  ArrayWednesday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (4): {
                    CoupleFourWednesday = CoupleFourWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetCoupleType() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetTeacherName() + " " + ArrayWednesday.get(i).GetAud() + " (" +  ArrayWednesday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (5): {
                    CoupleFiveWednesday = CoupleFiveWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetCoupleType() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetTeacherName() + " " + ArrayWednesday.get(i).GetAud() + " (" +  ArrayWednesday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (6): {
                    CoupleSixWednesday = CoupleSixWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetCoupleType() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetTeacherName() + " " + ArrayWednesday.get(i).GetAud() + " (" +  ArrayWednesday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (7): {
                    CoupleSevenWednesday = CoupleSevenWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetCoupleType() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetTeacherName() + " " + ArrayWednesday.get(i).GetAud() + " (" +  ArrayWednesday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
            }
        }

        Cell cellOneCoupleWednesday = rowGroupOne.createCell(3);
        cellOneCoupleWednesday.setCellValue(CoupleOneWednesday);
        cellOneCoupleWednesday.setCellStyle(cellStyle);
        Group.setColumnWidth(3, ColumnWidth);
        rowGroupOne.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleWednesday = rowGroupTwo.createCell(3);
        cellTwoCoupleWednesday.setCellValue(CoupleTwoWednesday);
        cellTwoCoupleWednesday.setCellStyle(cellStyle);
        Group.setColumnWidth(3, ColumnWidth);
        rowGroupTwo.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleWednesday = rowGroupThree.createCell(3);
        cellThreeCoupleWednesday.setCellValue(CoupleThreeWednesday);
        cellThreeCoupleWednesday.setCellStyle(cellStyle);
        Group.setColumnWidth(3, ColumnWidth);
        rowGroupThree.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleWednesday = rowGroupFour.createCell(3);
        cellFourCoupleWednesday.setCellValue(CoupleFourWednesday);
        cellFourCoupleWednesday.setCellStyle(cellStyle);
        Group.setColumnWidth(3, ColumnWidth);
        rowGroupFour.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleWednesday = rowGroupFive.createCell(3);
        cellFiveCoupleWednesday.setCellValue(CoupleFiveWednesday);
        cellFiveCoupleWednesday.setCellStyle(cellStyle);
        Group.setColumnWidth(3, ColumnWidth);
        rowGroupFive.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleWednesday = rowGroupSix.createCell(3);
        cellSixCoupleWednesday.setCellValue(CoupleSixWednesday);
        cellSixCoupleWednesday.setCellStyle(cellStyle);
        Group.setColumnWidth(3, ColumnWidth);
        rowGroupSix.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleWednesday = rowGroupSeven.createCell(3);
        cellSevenCoupleWednesday.setCellValue(CoupleSevenWednesday);
        cellSevenCoupleWednesday.setCellStyle(cellStyle);
        Group.setColumnWidth(3, ColumnWidth);
        rowGroupSeven.setHeightInPoints(HeightPoints);

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
                    CoupleOneThursday = CoupleOneThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetCoupleType() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetTeacherName() + " " + ArrayThursday.get(i).GetAud() + " (" +  ArrayThursday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (2): {
                    CoupleTwoThursday = CoupleTwoThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetCoupleType() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetTeacherName() + " " + ArrayThursday.get(i).GetAud() + " (" +  ArrayThursday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (3): {
                    CoupleThreeThursday = CoupleThreeThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetCoupleType() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetTeacherName() + " " + ArrayThursday.get(i).GetAud() + " (" +  ArrayThursday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (4): {
                    CoupleFourThursday = CoupleFourThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetCoupleType() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetTeacherName() + " " + ArrayThursday.get(i).GetAud() + " (" +  ArrayThursday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (5): {
                    CoupleFiveThursday = CoupleFiveThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetCoupleType() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetTeacherName() + " " + ArrayThursday.get(i).GetAud() + " (" +  ArrayThursday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (6): {
                    CoupleSixThursday = CoupleSixThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetCoupleType() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetTeacherName() + " " + ArrayThursday.get(i).GetAud() + " (" +  ArrayThursday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (7): {
                    CoupleSevenThursday = CoupleSevenThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetCoupleType() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetTeacherName() + " " + ArrayThursday.get(i).GetAud() + " (" +  ArrayThursday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
            }
        }

        Cell cellOneCoupleThursday = rowGroupOne.createCell(4);
        cellOneCoupleThursday.setCellValue(CoupleOneThursday);
        cellOneCoupleThursday.setCellStyle(cellStyle);
        Group.setColumnWidth(4, ColumnWidth);
        rowGroupOne.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleThursday = rowGroupTwo.createCell(4);
        cellTwoCoupleThursday.setCellValue(CoupleTwoThursday);
        cellTwoCoupleThursday.setCellStyle(cellStyle);
        Group.setColumnWidth(4, ColumnWidth);
        rowGroupTwo.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleThursday = rowGroupThree.createCell(4);
        cellThreeCoupleThursday.setCellValue(CoupleThreeThursday);
        cellThreeCoupleThursday.setCellStyle(cellStyle);
        Group.setColumnWidth(4, ColumnWidth);
        rowGroupThree.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleThursday = rowGroupFour.createCell(4);
        cellFourCoupleThursday.setCellValue(CoupleFourThursday);
        cellFourCoupleThursday.setCellStyle(cellStyle);
        Group.setColumnWidth(4, ColumnWidth);
        rowGroupFour.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleThursday = rowGroupFive.createCell(4);
        cellFiveCoupleThursday.setCellValue(CoupleFiveThursday);
        cellFiveCoupleThursday.setCellStyle(cellStyle);
        Group.setColumnWidth(4, ColumnWidth);
        rowGroupFive.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleThursday = rowGroupSix.createCell(4);
        cellSixCoupleThursday.setCellValue(CoupleSixThursday);
        cellSixCoupleThursday.setCellStyle(cellStyle);
        Group.setColumnWidth(4, ColumnWidth);
        rowGroupSix.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleThursday = rowGroupSeven.createCell(4);
        cellSevenCoupleThursday.setCellValue(CoupleSevenThursday);
        cellSevenCoupleThursday.setCellStyle(cellStyle);
        Group.setColumnWidth(4, ColumnWidth);
        rowGroupSeven.setHeightInPoints(HeightPoints);

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
                    CoupleOneFriday = CoupleOneFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetCoupleType() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetTeacherName() + " " + ArrayFriday.get(i).GetAud() + " (" +  ArrayFriday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (2): {
                    CoupleTwoFriday = CoupleTwoFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetCoupleType() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetTeacherName() + " " + ArrayFriday.get(i).GetAud() + " (" +  ArrayFriday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (3): {
                    CoupleThreeFriday = CoupleThreeFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetCoupleType() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetTeacherName() + " " + ArrayFriday.get(i).GetAud() + " (" +  ArrayFriday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (4): {
                    CoupleFourFriday = CoupleFourFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetCoupleType() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetTeacherName() + " " + ArrayFriday.get(i).GetAud() + " (" +  ArrayFriday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (5): {
                    CoupleFiveFriday = CoupleFiveFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetCoupleType() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetTeacherName() + " " + ArrayFriday.get(i).GetAud() + " (" +  ArrayFriday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (6): {
                    CoupleSixFriday = CoupleSixFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetCoupleType() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetTeacherName() + " " + ArrayFriday.get(i).GetAud() + " (" +  ArrayFriday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (7): {
                    CoupleSevenFriday = CoupleSevenFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetCoupleType() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetTeacherName() + " " + ArrayFriday.get(i).GetAud() + " (" +  ArrayFriday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
            }
        }

        Cell cellOneCoupleFriday = rowGroupOne.createCell(5);
        cellOneCoupleFriday.setCellValue(CoupleOneFriday);
        cellOneCoupleFriday.setCellStyle(cellStyle);
        Group.setColumnWidth(5, ColumnWidth);
        rowGroupOne.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleFriday = rowGroupTwo.createCell(5);
        cellTwoCoupleFriday.setCellValue(CoupleTwoFriday);
        cellTwoCoupleFriday.setCellStyle(cellStyle);
        Group.setColumnWidth(5, ColumnWidth);
        rowGroupTwo.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleFriday = rowGroupThree.createCell(5);
        cellThreeCoupleFriday.setCellValue(CoupleThreeFriday);
        cellThreeCoupleFriday.setCellStyle(cellStyle);
        Group.setColumnWidth(5, ColumnWidth);
        rowGroupThree.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleFriday = rowGroupFour.createCell(5);
        cellFourCoupleFriday.setCellValue(CoupleFourFriday);
        cellFourCoupleFriday.setCellStyle(cellStyle);
        Group.setColumnWidth(5, ColumnWidth);
        rowGroupFour.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleFriday = rowGroupFive.createCell(5);
        cellFiveCoupleFriday.setCellValue(CoupleFiveFriday);
        cellFiveCoupleFriday.setCellStyle(cellStyle);
        Group.setColumnWidth(5, ColumnWidth);
        rowGroupFive.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleFriday = rowGroupSix.createCell(5);
        cellSixCoupleFriday.setCellValue(CoupleSixFriday);
        cellSixCoupleFriday.setCellStyle(cellStyle);
        Group.setColumnWidth(5, ColumnWidth);
        rowGroupSix.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleFriday = rowGroupSeven.createCell(5);
        cellSevenCoupleFriday.setCellValue(CoupleSevenFriday);
        cellSevenCoupleFriday.setCellStyle(cellStyle);
        Group.setColumnWidth(5, ColumnWidth);
        rowGroupSeven.setHeightInPoints(HeightPoints);

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
                    CoupleOneSaturday = CoupleOneSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetCoupleType() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetTeacherName() + " " + ArraySaturday.get(i).GetAud() + " (" +  ArraySaturday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (2): {
                    CoupleTwoSaturday = CoupleTwoSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetCoupleType() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetTeacherName() + " " + ArraySaturday.get(i).GetAud() + " (" +  ArraySaturday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (3): {
                    CoupleThreeSaturday = CoupleThreeSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetCoupleType() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetTeacherName() + " " + ArraySaturday.get(i).GetAud() + " (" +  ArraySaturday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (4): {
                    CoupleFourSaturday = CoupleFourSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetCoupleType() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetTeacherName() + " " + ArraySaturday.get(i).GetAud() + " (" +  ArraySaturday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (5): {
                    CoupleFiveSaturday = CoupleFiveSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetCoupleType() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetTeacherName() + " " + ArraySaturday.get(i).GetAud() + " (" +  ArraySaturday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (6): {
                    CoupleSixSaturday = CoupleSixSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetCoupleType() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetTeacherName() + " " + ArraySaturday.get(i).GetAud() + " (" +  ArraySaturday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
                case (7): {
                    CoupleSevenSaturday = CoupleSevenSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetCoupleType() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetTeacherName() + " " + ArraySaturday.get(i).GetAud() + " (" +  ArraySaturday.get(i).GetTypeWeek() + ")" + "\n" + "\n";
                    break;
                }
            }
        }

        Cell cellOneCoupleSaturday = rowGroupOne.createCell(6);
        cellOneCoupleSaturday.setCellValue(CoupleOneSaturday);
        cellOneCoupleSaturday.setCellStyle(cellStyle);
        Group.setColumnWidth(6, ColumnWidth);
        rowGroupOne.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleSaturday = rowGroupTwo.createCell(6);
        cellTwoCoupleSaturday.setCellValue(CoupleTwoSaturday);
        cellTwoCoupleSaturday.setCellStyle(cellStyle);
        Group.setColumnWidth(6, ColumnWidth);
        rowGroupTwo.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleSaturday = rowGroupThree.createCell(6);
        cellThreeCoupleSaturday.setCellValue(CoupleThreeSaturday);
        cellThreeCoupleSaturday.setCellStyle(cellStyle);
        Group.setColumnWidth(6, ColumnWidth);
        rowGroupThree.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleSaturday = rowGroupFour.createCell(6);
        cellFourCoupleSaturday.setCellValue(CoupleFourSaturday);
        cellFourCoupleSaturday.setCellStyle(cellStyle);
        Group.setColumnWidth(6, ColumnWidth);
        rowGroupFour.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleSaturday = rowGroupFive.createCell(6);
        cellFiveCoupleSaturday.setCellValue(CoupleFiveSaturday);
        cellFiveCoupleSaturday.setCellStyle(cellStyle);
        Group.setColumnWidth(6, ColumnWidth);
        rowGroupFive.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleSaturday = rowGroupSix.createCell(6);
        cellSixCoupleSaturday.setCellValue(CoupleSixSaturday);
        cellSixCoupleSaturday.setCellStyle(cellStyle);
        Group.setColumnWidth(6, ColumnWidth);
        rowGroupSix.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleSaturday = rowGroupSeven.createCell(6);
        cellSevenCoupleSaturday.setCellValue(CoupleSevenSaturday);
        cellSevenCoupleSaturday.setCellStyle(cellStyle);
        Group.setColumnWidth(6, ColumnWidth);
        rowGroupSeven.setHeightInPoints(HeightPoints);


        // I write the result to a file named OneGroupExelDoc.
        FileOutputStream fileOutputStream = new FileOutputStream("OneGroupExelDoc");

        // I'm writing it down through a workbook.
        workbookOneElementGroup.write(fileOutputStream);

        // Closing the recording stream.
        fileOutputStream.close();
    }

}
