package timetablekrylov.timetablekrylovgr;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class CreatorTableExelGroups {

    // Creating an Excel workbook.
    private Workbook workbookGroups = new HSSFWorkbook();

    // I am creating a sheet on which our tables will be located.
    private Sheet Groups = workbookGroups.createSheet("Группы");

    // Here I set the correct location for couples on Monday.
    // Creating a row located at index 0 (In table 1)
    private Row rowZero = Groups.createRow(0);

    // Creating a cell that stores the day of the week header.
    private Cell cellDayWeek = rowZero.createCell(0);

    // Creating a cell that stores the header of the pair number.
    private Cell cellGroup = rowZero.createCell(1);

    // Creating a row located at index 1.
    private Row rowOne = Groups.createRow(1);

    // Creating a Monday cell
    private Cell Monday = rowOne.createCell(0);

    // Creating a cell the first pair on Monday
    private Cell CoupleOneMonday = rowOne.createCell(1);

    // Creating a row located at index 2.
    private Row rowTwo = Groups.createRow(2);

    // I'm creating a cell for the second pair on Monday.
    private Cell CoupleTwoMonday = rowTwo.createCell(1);

    // Creating a row located at index 3.
    private Row rowThree = Groups.createRow(3);

    // I'm creating a cell for the third pair on Monday.
    private Cell CoupleThreeMonday = rowThree.createCell(1);

    // Creating a row located at index 4.
    private Row rowFour = Groups.createRow(4);

    // I'm creating the fourth pair cell on Monday.
    private Cell CoupleFourMonday = rowFour.createCell(1);

    // Creating a row located at index 5.
    private Row rowFive = Groups.createRow(5);

    // I'm creating the fifth pair cell on Monday.
    private Cell CoupleFiveMonday = rowFive.createCell(1);

    // Creating a row located at index 6.
    private Row rowSix = Groups.createRow(6);

    // I'm creating a cell for the sixth pair on Monday.
    private Cell CoupleSixMonday = rowSix.createCell(1);

    // Creating a row located at index 7.
    private Row rowSeven = Groups.createRow(7);

    // I'm creating the seventh pair cell on Monday.
    private Cell CoupleSevenMonday = rowSeven.createCell(1);

    // ----------------------------------------------------------------------------

    // Here I set the correct location for couples on Tuesday.
    private Row rowEight = Groups.createRow(8);
    private Cell Tuesday = rowEight.createCell(0);
    private Cell CoupleOneTuesday = rowEight.createCell(1);

    private Row rowNine = Groups.createRow(9);
    private Cell CoupleTwoTuesday = rowNine.createCell(1);

    private Row rowTen = Groups.createRow(10);
    private Cell CoupleThreeTuesday = rowTen.createCell(1);

    private Row rowEleven = Groups.createRow(11);
    private Cell CoupleFourTuesday = rowEleven.createCell(1);

    private Row rowTwelve = Groups.createRow(12);
    private Cell CoupleFiveTuesday = rowTwelve.createCell(1);

    private Row rowThirteen = Groups.createRow(13);
    private Cell CoupleSixTuesday = rowThirteen.createCell(1);

    private Row rowFourteen = Groups.createRow(14);
    private Cell CoupleSevenTuesday = rowFourteen.createCell(1);

    // ----------------------------------------------------------------------------

    // Here I set the correct location for couples on Wednesday.
    private Row rowfifteen = Groups.createRow(15);
    private Cell Wednesday = rowfifteen.createCell(0);
    private Cell CoupleOneWednesday = rowfifteen.createCell(1);

    private Row rowSixteen = Groups.createRow(16);
    private Cell CoupleTwoWednesday = rowSixteen.createCell(1);

    private Row rowSeventeen = Groups.createRow(17);
    private Cell CoupleThreeWednesday = rowSeventeen.createCell(1);

    private Row rowEighteen = Groups.createRow(18);
    private Cell CoupleFourWednesday = rowEighteen.createCell(1);

    private Row rowNineteen = Groups.createRow(19);
    private Cell CoupleFiveWednesday = rowNineteen.createCell(1);

    private Row rowTwenty = Groups.createRow(20);
    private Cell CoupleSixWednesday = rowTwenty.createCell(1);

    private Row rowTwentyOne = Groups.createRow(21);
    private Cell CoupleSevenWednesday = rowTwentyOne.createCell(1);

    // ----------------------------------------------------------------------------

    // Here I set the correct location for couples on Thursday.
    private Row rowTwentyTwo = Groups.createRow(22);
    private Cell Thursday = rowTwentyTwo.createCell(0);
    private Cell CoupleOneThursday = rowTwentyTwo.createCell(1);

    private Row rowTwentyThree = Groups.createRow(23);
    private Cell CoupleTwoThursday = rowTwentyThree.createCell(1);

    private Row rowTwentyFour = Groups.createRow(24);
    private Cell CoupleThreeThursday = rowTwentyFour.createCell(1);

    private Row rowTwentyFive = Groups.createRow(25);
    private Cell CoupleFourThursday = rowTwentyFive.createCell(1);

    private Row rowTwentySix = Groups.createRow(26);
    private Cell CoupleFiveThursday = rowTwentySix.createCell(1);

    private Row rowTwentySeven = Groups.createRow(27);
    private Cell CoupleSixThursday = rowTwentySeven.createCell(1);

    private Row rowTwentyEight = Groups.createRow(28);
    private Cell CoupleSevenThursday = rowTwentyEight.createCell(1);


    // ----------------------------------------------------------------------------

    // Here I set the correct location for couples on Friday.
    private Row rowTwentyNine = Groups.createRow(29);
    private Cell Friday = rowTwentyNine.createCell(0);
    private Cell CoupleOneFriday = rowTwentyNine.createCell(1);

    private Row rowThirty = Groups.createRow(30);
    private Cell CoupleTwoFriday = rowThirty.createCell(1);

    private Row rowThirtyOne = Groups.createRow(31);
    private Cell CoupleThreeFriday = rowThirtyOne.createCell(1);

    private Row rowThirtyTwo = Groups.createRow(32);
    private Cell CoupleFourFriday = rowThirtyTwo.createCell(1);

    private Row rowThirtyThree = Groups.createRow(33);
    private Cell CoupleFiveFriday = rowThirtyThree.createCell(1);

    private Row rowThirtyFour = Groups.createRow(34);
    private Cell CoupleSixFriday = rowThirtyFour.createCell(1);

    private Row rowThirtyFive = Groups.createRow(35);
    private Cell CoupleSevenFriday = rowThirtyFive.createCell(1);

    // ---------------------------------------------------------------------------

    // Here I set the correct location for couples on Saturday.
    private Row rowThirtySix = Groups.createRow(36);
    private Cell Saturday = rowThirtySix.createCell(0);
    private Cell CoupleOneSaturday = rowThirtySix.createCell(1);

    private Row rowThirtySeven = Groups.createRow(37);
    private Cell CoupleTwoSaturday = rowThirtySeven.createCell(1);

    private Row rowThirtyEight = Groups.createRow(38);
    private Cell CoupleThreeSaturday = rowThirtyEight.createCell(1);

    private Row rowThirtyNine = Groups.createRow(39);
    private Cell CoupleFourSaturday = rowThirtyNine.createCell(1);

    private Row rowForty = Groups.createRow(40);
    private Cell CoupleFiveSaturday = rowForty.createCell(1);

    private Row rowFortyOne = Groups.createRow(41);
    private Cell CoupleSixSaturday = rowFortyOne.createCell(1);

    private Row rowFortyTwo = Groups.createRow(42);
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

    // I am creating a method to create a table from multiple groups.
    // It accepts an array of groups.
    public void CreatorTimeTableGroups(ArrayList<Group> ArrayGroup) throws IOException {

        int HeightPoints = 50;
        int ColumnWidth = 12000;

        ArrayList<CoupleGroup> CoupleOne = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleTwo = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleThree = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleFour = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleFive = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleSix = new ArrayList<>();
        ArrayList<CoupleGroup> CoupleSeven = new ArrayList<>();

//        int HeightOneCouple = HeightPoints;
//        int HeightTwoCouple = HeightPoints;
//        int HeightThreeCouple = HeightPoints;
//        int HeightFourCouple = HeightPoints;
//        int HeightFiveCouple = HeightPoints;
//        int HeightSixCouple = HeightPoints;
//        int HeightSevenCouple = HeightPoints;
//
//        int MaxHeightOneCouple = 0;
//        int MaxHeightTwoCouple = 0;
//        int MaxHeightThreeCouple = 0;
//        int MaxHeightFourCouple = 0;
//        int MaxHeightFiveCouple = 0;
//        int MaxHeightSixCouple = 0;
//        int MaxHeightSevenCouple = 0;

        // I give permission to move a line in cells.
        CellStyle cellStyle = workbookGroups.createCellStyle();
        cellStyle.setWrapText(true);

        // I give permission for line wrapping
        cellDayWeek.setCellStyle(cellStyle);

        // Setting values for the cell (cellDayWeek).
        cellDayWeek.setCellValue("День недели");

        // I set the dimensions of the cell.
        Groups.setColumnWidth(0,8000);
        rowZero.setHeightInPoints(30);

        // I give permission for line wrapping
        cellGroup.setCellStyle(cellStyle);

        // Setting values for the cell (cellGroup).
        cellGroup.setCellValue("Номер пары");

        // I set the dimensions of the cell.
        Groups.setColumnWidth(1,8000);
        rowZero.setHeightInPoints(30);

        // Setting values for the cell (Monday).
        Monday.setCellValue("Понедельник");

        // Setting values for the cell (Tuesday).
        Tuesday.setCellValue("Вторник");

        // Setting values for the cell (Wednesday).
        Wednesday.setCellValue("Среда");

        // Setting values for the cell (Thursday).
        Thursday.setCellValue("Четверг");

        // Setting values for the cell (Friday).
        Friday.setCellValue("Пятница");

        // Setting values for the cell (Saturday).
        Saturday.setCellValue("Суббота");

        // Setting values for the cell (CoupleOneMonday).
        CoupleOneMonday.setCellValue("1 пара");

        // Setting values for the cell (CoupleTwoMonday).
        CoupleTwoMonday.setCellValue("2 пара");

        // Setting values for the cell (CoupleThreeMonday).
        CoupleThreeMonday.setCellValue("3 пара");

        // Setting values for the cell (CoupleFourMonday).
        CoupleFourMonday.setCellValue("4 пара");

        // Setting values for the cell (CoupleFiveMonday).
        CoupleFiveMonday.setCellValue("5 пара");

        // Setting values for the cell (CoupleSixMonday).
        CoupleSixMonday.setCellValue("6 пара");

        // Setting values for the cell (CoupleSevenMonday).
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

        // The cells are combined
        Groups.addMergedRegion(new CellRangeAddress(36,42,0,0));
        Groups.addMergedRegion(new CellRangeAddress(29,35,0,0));
        Groups.addMergedRegion(new CellRangeAddress(22,28,0,0));
        Groups.addMergedRegion(new CellRangeAddress(15,21,0,0));
        Groups.addMergedRegion(new CellRangeAddress(8,14,0,0));
        Groups.addMergedRegion(new CellRangeAddress(1,7,0,0));

        // I create variables for displaying group names in cells.
        Cell cell;
        int numberGroup = 2;

        // I'm going through an array of groups.
        for(int i = 0; i < ArrayGroup.size(); i++){

            // I get the name of the group.
            String GroupName = ArrayGroup.get(i).GetGroupName();

            // I create arrays to distribute pairs by day of the week.
            ArrayList<CoupleGroup> ArrayMonday = new ArrayList<>();
            ArrayList<CoupleGroup> ArrayTuesday = new ArrayList<>();
            ArrayList<CoupleGroup> ArrayWednesday = new ArrayList<>();
            ArrayList<CoupleGroup> ArrayThursday = new ArrayList<>();
            ArrayList<CoupleGroup> ArrayFriday = new ArrayList<>();
            ArrayList<CoupleGroup> ArraySaturday = new ArrayList<>();

            // I take pairs of them from a particular group.
            ArrayList<CoupleGroup> ArrayCouple = ArrayGroup.get(i).GetArrayCouples();

            // I'm sorting through the pairs.
            for(int index = 0; index < ArrayCouple.size(); index++) {

                // I get the day of the week of each pair.
                int IdDay = ArrayCouple.get(index).GetIDDay();

                // I sort pairs by day of the week using the switch.
                switch (IdDay) {
                    case (1): {
                        ArrayMonday.add(ArrayCouple.get(index));
                        break;
                    }
                    case (2): {
                        ArrayTuesday.add(ArrayCouple.get(index));
                        break;
                    }
                    case (3): {
                        ArrayWednesday.add(ArrayCouple.get(index));
                        break;
                    }
                    case (4): {
                        ArrayThursday.add(ArrayCouple.get(index));
                        break;
                    }
                    case (5): {
                        ArrayFriday.add(ArrayCouple.get(index));
                        break;
                    }
                    case (6): {
                        ArraySaturday.add(ArrayCouple.get(index));
                        break;
                    }
                }
            }

                // Couple Monday

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

                // I'm going through Monday's pairs.
                for(int index2 = 0; index2 < ArrayMonday.size(); index2++){

                    // I sort the pairs according to their number.
                    switch (ArrayMonday.get(index2).GetCoupleNumber()) {
                        case (1): {
                            // I combine the pairs that are held by the first couple on Monday.
                            CoupleOne.add(ArrayMonday.get(index2));
                            break;
                        }
                        case (2): {
                            // I combine the pairs that are held by the second pair on Monday.
                            CoupleTwo.add(ArrayMonday.get(index2));
                            break;
                        }
                        case (3): {
                            // I am combining pairs that are held by the third pair on Monday.
                            CoupleThree.add(ArrayMonday.get(index2));
                            break;
                        }
                        case (4): {
                            // I am combining pairs that are held by the fourth pair on Monday.
                            CoupleFour.add(ArrayMonday.get(index2));
                            break;
                        }
                        case (5): {
                            // I combine the pairs that are held by the fifth pair on Monday.
                            CoupleFive.add(ArrayMonday.get(index2));
                            break;
                        }
                        case (6): {
                            // I am combining pairs that are held by the sixth pair on Monday.
                            CoupleSix.add(ArrayMonday.get(index2));
                            break;
                        }
                        case (7): {
                            // I am combining pairs that are held by the seventh pair on Monday.
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

                // I am creating a cell (the first pair is on Monday). (1 x numberGroup)
                // The variable numberGroup is responsible for the column.
                Cell cellOneCoupleMonday = rowOne.createCell(numberGroup);

                // Setting the value
                cellOneCoupleMonday.setCellValue(CoupleOneMonday);

                // I give permission to move a line in cells.
                cellOneCoupleMonday.setCellStyle(cellStyle);

                // I set the dimensions of the cell.
                Groups.setColumnWidth(numberGroup,ColumnWidth);
                rowOne.setHeightInPoints(HeightPoints);

                // I'm creating a cell (the second pair is on Monday). (2 x numberGroup)
                Cell cellTwoCoupleMonday = rowTwo.createCell(numberGroup);

                // Setting the value
                cellTwoCoupleMonday.setCellValue(CoupleTwoMonday);

                // I give permission to move a line in cells.
                cellTwoCoupleMonday.setCellStyle(cellStyle);

                // I set the dimensions of the cell.
                Groups.setColumnWidth(numberGroup,ColumnWidth);
                rowTwo.setHeightInPoints(HeightPoints);

                // I'm creating a cell (the third pair is on Monday). (3 x numberGroup)
                Cell cellThreeCoupleMonday = rowThree.createCell(numberGroup);

                // Setting the value
                cellThreeCoupleMonday.setCellValue(CoupleThreeMonday);

                // I give permission to move a line in cells.
                cellThreeCoupleMonday.setCellStyle(cellStyle);

                // I set the dimensions of the cell.
                Groups.setColumnWidth(numberGroup,ColumnWidth);
                rowThree.setHeightInPoints(HeightPoints);

                // I am creating a cell (the fourth pair is on Monday). (4 x numberGroup)
                Cell cellFourCoupleMonday = rowFour.createCell(numberGroup);

                // Setting the value
                cellFourCoupleMonday.setCellValue(CoupleFourMonday);

                // I give permission to move a line in cells.
                cellFourCoupleMonday.setCellStyle(cellStyle);

                // I set the dimensions of the cell.
                Groups.setColumnWidth(numberGroup,ColumnWidth);
                rowFour.setHeightInPoints(HeightPoints);

                // I am creating a cell (the fifth pair is on Monday). (5 x numberGroup)
                Cell cellFiveCoupleMonday = rowFive.createCell(numberGroup);

                // Setting the value
                cellFiveCoupleMonday.setCellValue(CoupleFiveMonday);

                // I give permission to move a line in cells.
                cellFiveCoupleMonday.setCellStyle(cellStyle);

                // I set the dimensions of the cell.
                Groups.setColumnWidth(numberGroup,ColumnWidth);
                rowFive.setHeightInPoints(HeightPoints);

                // I'm creating a cell (the sixth pair is on Monday). (6 x numberGroup)
                Cell cellSixCoupleMonday = rowSix.createCell(numberGroup);

                // Setting the value
                cellSixCoupleMonday.setCellValue(CoupleSixMonday);

                // I give permission to move a line in cells.
                cellSixCoupleMonday.setCellStyle(cellStyle);

                // I set the dimensions of the cell.
                Groups.setColumnWidth(numberGroup,ColumnWidth);
                rowSix.setHeightInPoints(HeightPoints);

                // I am creating a cell (the seventh pair is on Monday). (7 x numberGroup)
                Cell cellSevenCoupleMonday = rowSeven.createCell(numberGroup);

                // Setting the value
                cellSevenCoupleMonday.setCellValue(CoupleSevenMonday);

                // I give permission to move a line in cells.
                cellSevenCoupleMonday.setCellStyle(cellStyle);

                // I set the dimensions of the cell.
                Groups.setColumnWidth(numberGroup,ColumnWidth);
                rowSeven.setHeightInPoints(HeightPoints);

                CoupleOne.clear();
                CoupleTwo.clear();
                CoupleThree.clear();
                CoupleFour.clear();
                CoupleFive.clear();
                CoupleSix.clear();
                CoupleSeven.clear();
//
//                HeightOneCouple = HeightPoints;
//                HeightTwoCouple = HeightPoints;
//                HeightThreeCouple = HeightPoints;
//                HeightFourCouple = HeightPoints;
//                HeightFiveCouple = HeightPoints;
//                HeightSixCouple = HeightPoints;
//                HeightSevenCouple = HeightPoints;

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

                Cell cellOneCoupleTuesday = rowEight.createCell(numberGroup);
                cellOneCoupleTuesday.setCellValue(CoupleOneTuesday);
                cellOneCoupleTuesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup,ColumnWidth);

//                if(MaxHeightOneCouple < HeightOneCouple){
//                    MaxHeightOneCouple = HeightOneCouple;
//                }

                rowEight.setHeightInPoints(HeightPoints);

                Cell cellTwoCoupleTuesday = rowNine.createCell(numberGroup);
                cellTwoCoupleTuesday.setCellValue(CoupleTwoTuesday);
                cellTwoCoupleTuesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup,ColumnWidth);

//                if(MaxHeightTwoCouple < HeightTwoCouple){
//                    MaxHeightTwoCouple = HeightTwoCouple;
//                }

                rowNine.setHeightInPoints(HeightPoints);

                Cell cellThreeCoupleTuesday = rowTen.createCell(numberGroup);
                cellThreeCoupleTuesday.setCellValue(CoupleThreeTuesday);
                cellThreeCoupleTuesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup,ColumnWidth);

//                if(MaxHeightThreeCouple < HeightThreeCouple){
//                    MaxHeightThreeCouple = HeightThreeCouple;
//                }

                rowTen.setHeightInPoints(HeightPoints);

                Cell cellFourCoupleTuesday = rowEleven.createCell(numberGroup);
                cellFourCoupleTuesday.setCellValue(CoupleFourTuesday);
                cellFourCoupleTuesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup,ColumnWidth);

//                if(MaxHeightFourCouple < HeightFourCouple){
//                    MaxHeightFourCouple = HeightFourCouple;
//                }

                rowEleven.setHeightInPoints(HeightPoints);

                Cell cellFiveCoupleTuesday = rowTwelve.createCell(numberGroup);
                cellFiveCoupleTuesday.setCellValue(CoupleFiveTuesday);
                cellFiveCoupleTuesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup,ColumnWidth);

//                if(MaxHeightFiveCouple < HeightFiveCouple){
//                    MaxHeightFiveCouple = HeightFiveCouple;
//                }

                rowTwelve.setHeightInPoints(HeightPoints);

                Cell cellSixCoupleTuesday = rowThirteen.createCell(numberGroup);
                cellSixCoupleTuesday.setCellValue(CoupleSixTuesday);
                cellSixCoupleTuesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup,ColumnWidth);

//                if(MaxHeightSixCouple < HeightSixCouple){
//                    MaxHeightSixCouple = HeightSixCouple;
//                }

                rowThirteen.setHeightInPoints(HeightPoints);

                Cell cellSevenCoupleTuesday = rowFourteen.createCell(numberGroup);
                cellSevenCoupleTuesday.setCellValue(CoupleSevenTuesday);
                cellSevenCoupleTuesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup,ColumnWidth);

//                if(MaxHeightSevenCouple < HeightSevenCouple){
//                    MaxHeightSevenCouple = HeightSevenCouple;
//                }

                rowFourteen.setHeightInPoints(HeightPoints);

                CoupleOne.clear();
                CoupleTwo.clear();
                CoupleThree.clear();
                CoupleFour.clear();
                CoupleFive.clear();
                CoupleSix.clear();
                CoupleSeven.clear();

//                HeightOneCouple = HeightPoints;
//                HeightTwoCouple = HeightPoints;
//                HeightThreeCouple = HeightPoints;
//                HeightFourCouple = HeightPoints;
//                HeightFiveCouple = HeightPoints;
//                HeightSixCouple = HeightPoints;
//                HeightSevenCouple = HeightPoints;

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

                Cell cellOneCoupleWednesday = rowfifteen.createCell(numberGroup);
                cellOneCoupleWednesday.setCellValue(CoupleOneWednesday);
                cellOneCoupleWednesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup,ColumnWidth);

//                if(MaxHeightOneCouple < HeightOneCouple){
//                    MaxHeightOneCouple = HeightOneCouple;
//                }

                rowfifteen.setHeightInPoints(HeightPoints);

                Cell cellTwoCoupleWednesday = rowSixteen.createCell(numberGroup);
                cellTwoCoupleWednesday.setCellValue(CoupleTwoWednesday);
                cellTwoCoupleWednesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup,ColumnWidth);

//                if(MaxHeightTwoCouple < HeightTwoCouple){
//                    MaxHeightTwoCouple = HeightTwoCouple;
//                }

                rowSixteen.setHeightInPoints(HeightPoints);

                Cell cellThreeCoupleWednesday = rowSeventeen.createCell(numberGroup);
                cellThreeCoupleWednesday.setCellValue(CoupleThreeWednesday);
                cellThreeCoupleWednesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup,ColumnWidth);

//                if(MaxHeightThreeCouple < HeightThreeCouple){
//                    MaxHeightThreeCouple = HeightThreeCouple;
//                }

                rowSeventeen.setHeightInPoints(HeightPoints);

                Cell cellFourCoupleWednesday = rowEighteen.createCell(numberGroup);
                cellFourCoupleWednesday.setCellValue(CoupleFourWednesday);
                cellFourCoupleWednesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup,ColumnWidth);

//                if(MaxHeightFourCouple < HeightFourCouple){
//                    MaxHeightFourCouple = HeightFourCouple;
//                }

                rowEighteen.setHeightInPoints(HeightPoints);

                Cell cellFiveCoupleWednesday = rowNineteen.createCell(numberGroup);
                cellFiveCoupleWednesday.setCellValue(CoupleFiveWednesday);
                cellFiveCoupleWednesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup,ColumnWidth);

//                if(MaxHeightFiveCouple < HeightFiveCouple){
//                    MaxHeightFiveCouple = HeightFiveCouple;
//                }

                rowNineteen.setHeightInPoints(HeightPoints);

                Cell cellSixCoupleWednesday = rowTwenty.createCell(numberGroup);
                cellSixCoupleWednesday.setCellValue(CoupleSixWednesday);
                cellSixCoupleWednesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup,ColumnWidth);

//                if(MaxHeightSixCouple < HeightSixCouple){
//                    MaxHeightSixCouple = HeightSixCouple;
//                }

                rowTwenty.setHeightInPoints(HeightPoints);

                Cell cellSevenCoupleWednesday = rowTwentyOne.createCell(numberGroup);
                cellSevenCoupleWednesday.setCellValue(CoupleSevenWednesday);
                cellSevenCoupleWednesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup,ColumnWidth);

//                if(MaxHeightSevenCouple < HeightSevenCouple){
//                    MaxHeightSevenCouple = HeightSevenCouple;
//                }

                rowTwentyOne.setHeightInPoints(HeightPoints);

                CoupleOne.clear();
                CoupleTwo.clear();
                CoupleThree.clear();
                CoupleFour.clear();
                CoupleFive.clear();
                CoupleSix.clear();
                CoupleSeven.clear();

//            HeightOneCouple = HeightPoints;
//            HeightTwoCouple = HeightPoints;
//            HeightThreeCouple = HeightPoints;
//            HeightFourCouple = HeightPoints;
//            HeightFiveCouple = HeightPoints;
//            HeightSixCouple = HeightPoints;
//            HeightSevenCouple = HeightPoints;

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

            Cell cellOneCoupleThursday = rowTwentyTwo.createCell(numberGroup);
            cellOneCoupleThursday.setCellValue(CoupleOneThursday);
            cellOneCoupleThursday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightOneCouple < HeightOneCouple){
//                MaxHeightOneCouple = HeightOneCouple;
//            }

            rowTwentyTwo.setHeightInPoints(HeightPoints);

            Cell cellTwoCoupleThursday = rowTwentyThree.createCell(numberGroup);
            cellTwoCoupleThursday.setCellValue(CoupleTwoThursday);
            cellTwoCoupleThursday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightTwoCouple < HeightTwoCouple){
//                MaxHeightTwoCouple = HeightTwoCouple;
//            }

            rowTwentyThree.setHeightInPoints(HeightPoints);

            Cell cellThreeCoupleThursday = rowTwentyFour.createCell(numberGroup);
            cellThreeCoupleThursday.setCellValue(CoupleThreeThursday);
            cellThreeCoupleThursday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightThreeCouple < HeightThreeCouple){
//                MaxHeightThreeCouple = HeightThreeCouple;
//            }

            rowTwentyFour.setHeightInPoints(HeightPoints);

            Cell cellFourCoupleThursday = rowTwentyFive.createCell(numberGroup);
            cellFourCoupleThursday.setCellValue(CoupleFourThursday);
            cellFourCoupleThursday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightFourCouple < HeightFourCouple){
//                MaxHeightFourCouple = HeightFourCouple;
//            }

            rowTwentyFive.setHeightInPoints(HeightPoints);

            Cell cellFiveCoupleThursday = rowTwentySix.createCell(numberGroup);
            cellFiveCoupleThursday.setCellValue(CoupleFiveThursday);
            cellFiveCoupleThursday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightFiveCouple < HeightFiveCouple){
//                MaxHeightFiveCouple = HeightFiveCouple;
//            }

            rowTwentySix.setHeightInPoints(HeightPoints);

            Cell cellSixCoupleThursday = rowTwentySeven.createCell(numberGroup);
            cellSixCoupleThursday.setCellValue(CoupleSixThursday);
            cellSixCoupleThursday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightSixCouple < HeightSixCouple){
//                MaxHeightSixCouple = HeightSixCouple;
//            }

            rowTwentySeven.setHeightInPoints(HeightPoints);

            Cell cellSevenCoupleThursday = rowTwentyEight.createCell(numberGroup);
            cellSevenCoupleThursday.setCellValue(CoupleSevenThursday);
            cellSevenCoupleThursday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightSevenCouple < HeightSevenCouple){
//                MaxHeightSevenCouple = HeightSevenCouple;
//            }

            rowTwentyEight.setHeightInPoints(HeightPoints);

            CoupleOne.clear();
            CoupleTwo.clear();
            CoupleThree.clear();
            CoupleFour.clear();
            CoupleFive.clear();
            CoupleSix.clear();
            CoupleSeven.clear();

//            HeightOneCouple = HeightPoints;
//            HeightTwoCouple = HeightPoints;
//            HeightThreeCouple = HeightPoints;
//            HeightFourCouple = HeightPoints;
//            HeightFiveCouple = HeightPoints;
//            HeightSixCouple = HeightPoints;
//            HeightSevenCouple = HeightPoints;

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

            Cell cellOneCoupleFriday = rowTwentyNine.createCell(numberGroup);
            cellOneCoupleFriday.setCellValue(CoupleOneFriday);
            cellOneCoupleFriday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightOneCouple < HeightOneCouple){
//                MaxHeightOneCouple = HeightOneCouple;
//            }

            rowTwentyNine.setHeightInPoints(HeightPoints);

            Cell cellTwoCoupleFriday = rowThirty.createCell(numberGroup);
            cellTwoCoupleFriday.setCellValue(CoupleTwoFriday);
            cellTwoCoupleFriday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightTwoCouple < HeightTwoCouple){
//                MaxHeightTwoCouple = HeightTwoCouple;
//            }

            rowThirty.setHeightInPoints(HeightPoints);

            Cell cellThreeCoupleFriday = rowThirtyOne.createCell(numberGroup);
            cellThreeCoupleFriday.setCellValue(CoupleThreeFriday);
            cellThreeCoupleFriday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightThreeCouple < HeightThreeCouple){
//                MaxHeightThreeCouple = HeightThreeCouple;
//            }

            rowThirtyOne.setHeightInPoints(HeightPoints);

            Cell cellFourCoupleFriday = rowThirtyTwo.createCell(numberGroup);
            cellFourCoupleFriday.setCellValue(CoupleFourFriday);
            cellFourCoupleFriday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightFourCouple < HeightFourCouple){
//                MaxHeightFourCouple = HeightFourCouple;
//            }

            rowThirtyTwo.setHeightInPoints(HeightPoints);

            Cell cellFiveCoupleFriday = rowThirtyThree.createCell(numberGroup);
            cellFiveCoupleFriday.setCellValue(CoupleFiveFriday);
            cellFiveCoupleFriday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightFiveCouple < HeightFiveCouple){
//                MaxHeightFiveCouple = HeightFiveCouple;
//            }

            rowThirtyThree.setHeightInPoints(HeightPoints);

            Cell cellSixCoupleFriday = rowThirtyFour.createCell(numberGroup);
            cellSixCoupleFriday.setCellValue(CoupleSixFriday);
            cellSixCoupleFriday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightSixCouple < HeightSixCouple){
//                MaxHeightSixCouple = HeightSixCouple;
//            }

            rowThirtyFour.setHeightInPoints(HeightPoints);

            Cell cellSevenCoupleFriday = rowThirtyFive.createCell(numberGroup);
            cellSevenCoupleFriday.setCellValue(CoupleSevenFriday);
            cellSevenCoupleFriday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightSevenCouple < HeightSevenCouple){
//                MaxHeightSevenCouple = HeightSevenCouple;
//            }

            rowThirtyFive.setHeightInPoints(HeightPoints);

            CoupleOne.clear();
            CoupleTwo.clear();
            CoupleThree.clear();
            CoupleFour.clear();
            CoupleFive.clear();
            CoupleSix.clear();
            CoupleSeven.clear();

//            HeightOneCouple = HeightPoints;
//            HeightTwoCouple = HeightPoints;
//            HeightThreeCouple = HeightPoints;
//            HeightFourCouple = HeightPoints;
//            HeightFiveCouple = HeightPoints;
//            HeightSixCouple = HeightPoints;
//            HeightSevenCouple = HeightPoints;

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

            Cell cellOneCoupleSaturday = rowThirtySix.createCell(numberGroup);
            cellOneCoupleSaturday.setCellValue(CoupleOneSaturday);
            cellOneCoupleSaturday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightOneCouple < HeightOneCouple){
//                MaxHeightOneCouple = HeightOneCouple;
//            }

            rowThirtySix.setHeightInPoints(HeightPoints);

            Cell cellTwoCoupleSaturday = rowThirtySeven.createCell(numberGroup);
            cellTwoCoupleSaturday.setCellValue(CoupleTwoSaturday);
            cellTwoCoupleSaturday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightTwoCouple < HeightTwoCouple){
//                MaxHeightTwoCouple = HeightTwoCouple;
//            }

            rowThirtySeven.setHeightInPoints(HeightPoints);

            Cell cellThreeCoupleSaturday = rowThirtyEight.createCell(numberGroup);
            cellThreeCoupleSaturday.setCellValue(CoupleThreeSaturday);
            cellThreeCoupleSaturday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightThreeCouple < HeightThreeCouple){
//                MaxHeightThreeCouple = HeightThreeCouple;
//            }

            rowThirtyEight.setHeightInPoints(HeightPoints);

            Cell cellFourCoupleSaturday = rowThirtyNine.createCell(numberGroup);
            cellFourCoupleSaturday.setCellValue(CoupleFourSaturday);
            cellFourCoupleSaturday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightFourCouple < HeightFourCouple){
//                MaxHeightFourCouple = HeightFourCouple;
//            }

            rowThirtyNine.setHeightInPoints(HeightPoints);

            Cell cellFiveCoupleSaturday = rowForty.createCell(numberGroup);
            cellFiveCoupleSaturday.setCellValue(CoupleFiveSaturday);
            cellFiveCoupleSaturday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightFiveCouple < HeightFiveCouple){
//                MaxHeightFiveCouple = HeightFiveCouple;
//            }

            rowForty.setHeightInPoints(HeightPoints);

            Cell cellSixCoupleSaturday = rowFortyOne.createCell(numberGroup);
            cellSixCoupleSaturday.setCellValue(CoupleSixSaturday);
            cellSixCoupleSaturday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightSixCouple < HeightSixCouple){
//                MaxHeightSixCouple = HeightSixCouple;
//            }

            rowFortyOne.setHeightInPoints(HeightPoints);

            Cell cellSevenCoupleSaturday = rowFortyTwo.createCell(numberGroup);
            cellSevenCoupleSaturday.setCellValue(CoupleSevenSaturday);
            cellSevenCoupleSaturday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup,ColumnWidth);

//            if(MaxHeightSevenCouple < HeightSevenCouple){
//                MaxHeightSevenCouple = HeightSevenCouple;
//            }

            rowFortyTwo.setHeightInPoints(HeightPoints);

            CoupleOne.clear();
            CoupleTwo.clear();
            CoupleThree.clear();
            CoupleFour.clear();
            CoupleFive.clear();
            CoupleSix.clear();
            CoupleSeven.clear();

            // I'm creating a cell that stores the name of the group.
            cell = rowZero.createCell(numberGroup);

            // Setting the value.
            cell.setCellValue(GroupName);

            // I'm increasing it
            numberGroup++;

        }

        String separator = File.separator;

        // I write the result to a file named AllGroupExelDoc.
        FileOutputStream fileOutputStream = new FileOutputStream("TableGroup(s)" + separator + "AllGroupExelDoc");

        // I'm writing it down through a workbook.
        workbookGroups.write(fileOutputStream);

        // // Closing the recording stream.
        workbookGroups.close();

    }

}
