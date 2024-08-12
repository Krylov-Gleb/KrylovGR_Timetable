package timetablekrylov.timetablekrylovgr;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class CreatorTableExelGroups {

    int HeightPoints = 250;
    int ColumnWidth = 10000;

    private Workbook workbookGroups = new HSSFWorkbook();

    private Sheet Groups = workbookGroups.createSheet("Группы");

    private Row rowZero = Groups.createRow(0);
    private Cell cellDayWeek = rowZero.createCell(0);
    private Cell cellGroup = rowZero.createCell(1);

    private Row rowOne = Groups.createRow(1);
    private Cell Monday = rowOne.createCell(0);
    private Cell CoupleOneMonday = rowOne.createCell(1);

    private Row rowTwo = Groups.createRow(2);
    private Cell CoupleTwoMonday = rowTwo.createCell(1);

    private Row rowThree = Groups.createRow(3);
    private Cell CoupleThreeMonday = rowThree.createCell(1);

    private Row rowFour = Groups.createRow(4);
    private Cell CoupleFourMonday = rowFour.createCell(1);

    private Row rowFive = Groups.createRow(5);
    private Cell CoupleFiveMonday = rowFive.createCell(1);

    private Row rowSix = Groups.createRow(6);
    private Cell CoupleSixMonday = rowSix.createCell(1);

    private Row rowSeven = Groups.createRow(7);
    private Cell CoupleSevenMonday = rowSeven.createCell(1);

    // ----------------------------------------------------------------------------

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


    public void CreatorTimeTableGroups(ArrayList<Group> ArrayGroup) throws IOException {

        CellStyle cellStyle = workbookGroups.createCellStyle();
        cellStyle.setWrapText(true);

        cellDayWeek.setCellStyle(cellStyle);
        cellDayWeek.setCellValue("День недели");
        Groups.setColumnWidth(0,8000);
        rowZero.setHeightInPoints(30);

        cellGroup.setCellStyle(cellStyle);
        cellGroup.setCellValue("Номер пары");
        Groups.setColumnWidth(1,8000);
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

        Groups.addMergedRegion(new CellRangeAddress(36,42,0,0));
        Groups.addMergedRegion(new CellRangeAddress(29,35,0,0));
        Groups.addMergedRegion(new CellRangeAddress(22,28,0,0));
        Groups.addMergedRegion(new CellRangeAddress(15,21,0,0));
        Groups.addMergedRegion(new CellRangeAddress(8,14,0,0));
        Groups.addMergedRegion(new CellRangeAddress(1,7,0,0));

        Cell cell;
        int numberGroup = 2;

        for(int i = 0; i < ArrayGroup.size(); i++){

            String GroupName = ArrayGroup.get(i).GetGroupName();

            ArrayList<CoupleGroup> ArrayMonday = new ArrayList<>();
            ArrayList<CoupleGroup> ArrayTuesday = new ArrayList<>();
            ArrayList<CoupleGroup> ArrayWednesday = new ArrayList<>();
            ArrayList<CoupleGroup> ArrayThursday = new ArrayList<>();
            ArrayList<CoupleGroup> ArrayFriday = new ArrayList<>();
            ArrayList<CoupleGroup> ArraySaturday = new ArrayList<>();

            ArrayList<CoupleGroup> ArrayCouple = ArrayGroup.get(i).GetArrayCouples();

            for(int index = 0; index < ArrayCouple.size(); index++){

                int IdDay = ArrayCouple.get(index).GetIDDay();

                switch (IdDay){
                    case (1):{
                        ArrayMonday.add(ArrayCouple.get(index));
                        break;
                    }
                    case (2):{
                        ArrayTuesday.add(ArrayCouple.get(index));
                        break;
                    }
                    case (3):{
                        ArrayWednesday.add(ArrayCouple.get(index));
                        break;
                    }
                    case (4):{
                        ArrayThursday.add(ArrayCouple.get(index));
                        break;
                    }
                    case (5):{
                        ArrayFriday.add(ArrayCouple.get(index));
                        break;
                    }
                    case (6):{
                        ArraySaturday.add(ArrayCouple.get(index));
                        break;
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

                for(int index2 = 0; index2 < ArrayMonday.size(); index2++){

                    switch (ArrayMonday.get(index2).GetCoupleNumber()) {
                        case (1): {
                            CoupleOneMonday = CoupleOneMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (2): {
                            CoupleTwoMonday = CoupleTwoMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (3): {
                            CoupleThreeMonday = CoupleThreeMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (4): {
                            CoupleFourMonday = CoupleFourMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (5): {
                            CoupleFiveMonday = CoupleFiveMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (6): {
                            CoupleSixMonday = CoupleSixMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (7): {
                            CoupleSevenMonday = CoupleSevenMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                    }
                }

                Cell cellOneCoupleMonday = rowOne.createCell(numberGroup);
                cellOneCoupleMonday.setCellValue(CoupleOneMonday);
                cellOneCoupleMonday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowOne.setHeightInPoints(HeightPoints);

                Cell cellTwoCoupleMonday = rowTwo.createCell(numberGroup);
                cellTwoCoupleMonday.setCellValue(CoupleTwoMonday);
                cellTwoCoupleMonday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowTwo.setHeightInPoints(HeightPoints);

                Cell cellThreeCoupleMonday = rowThree.createCell(numberGroup);
                cellThreeCoupleMonday.setCellValue(CoupleThreeMonday);
                cellThreeCoupleMonday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowThree.setHeightInPoints(HeightPoints);

                Cell cellFourCoupleMonday = rowFour.createCell(numberGroup);
                cellFourCoupleMonday.setCellValue(CoupleFourMonday);
                cellFourCoupleMonday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowFour.setHeightInPoints(HeightPoints);

                Cell cellFiveCoupleMonday = rowFive.createCell(numberGroup);
                cellFiveCoupleMonday.setCellValue(CoupleFiveMonday);
                cellFiveCoupleMonday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowFive.setHeightInPoints(HeightPoints);

                Cell cellSixCoupleMonday = rowSix.createCell(numberGroup);
                cellSixCoupleMonday.setCellValue(CoupleSixMonday);
                cellSixCoupleMonday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowSix.setHeightInPoints(HeightPoints);

                Cell cellSevenCoupleMonday = rowSeven.createCell(numberGroup);
                cellSevenCoupleMonday.setCellValue(CoupleSevenMonday);
                cellSevenCoupleMonday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowSeven.setHeightInPoints(HeightPoints);

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
                            CoupleOneTuesday = CoupleOneTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetTypeWeek() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (2): {
                            CoupleTwoTuesday = CoupleTwoTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetTypeWeek() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (3): {
                            CoupleThreeTuesday = CoupleThreeTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetTypeWeek() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (4): {
                            CoupleFourTuesday = CoupleFourTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetTypeWeek() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (5): {
                            CoupleFiveTuesday = CoupleFiveTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetTypeWeek() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (6): {
                            CoupleSixTuesday = CoupleSixTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetTypeWeek() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (7): {
                            CoupleSevenTuesday = CoupleSevenTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetTypeWeek() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                    }
                }

                Cell cellOneCoupleTuesday = rowEight.createCell(numberGroup);
                cellOneCoupleTuesday.setCellValue(CoupleOneTuesday);
                cellOneCoupleTuesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowEight.setHeightInPoints(HeightPoints);

                Cell cellTwoCoupleTuesday = rowNine.createCell(numberGroup);
                cellTwoCoupleTuesday.setCellValue(CoupleTwoTuesday);
                cellTwoCoupleTuesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowNine.setHeightInPoints(HeightPoints);

                Cell cellThreeCoupleTuesday = rowTen.createCell(numberGroup);
                cellThreeCoupleTuesday.setCellValue(CoupleThreeTuesday);
                cellThreeCoupleTuesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowTen.setHeightInPoints(HeightPoints);

                Cell cellFourCoupleTuesday = rowEleven.createCell(numberGroup);
                cellFourCoupleTuesday.setCellValue(CoupleFourTuesday);
                cellFourCoupleTuesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowEleven.setHeightInPoints(HeightPoints);

                Cell cellFiveCoupleTuesday = rowTwelve.createCell(numberGroup);
                cellFiveCoupleTuesday.setCellValue(CoupleFiveTuesday);
                cellFiveCoupleTuesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowTwelve.setHeightInPoints(HeightPoints);

                Cell cellSixCoupleTuesday = rowThirteen.createCell(numberGroup);
                cellSixCoupleTuesday.setCellValue(CoupleSixTuesday);
                cellSixCoupleTuesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowThirteen.setHeightInPoints(HeightPoints);

                Cell cellSevenCoupleTuesday = rowFourteen.createCell(numberGroup);
                cellSevenCoupleTuesday.setCellValue(CoupleSevenTuesday);
                cellSevenCoupleTuesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowFourteen.setHeightInPoints(HeightPoints);

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
                            CoupleOneWednesday = CoupleOneWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetTypeWeek() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (2): {
                            CoupleTwoWednesday = CoupleTwoWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetTypeWeek() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (3): {
                            CoupleThreeWednesday = CoupleThreeWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetTypeWeek() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (4): {
                            CoupleFourWednesday = CoupleFourWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetTypeWeek() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (5): {
                            CoupleFiveWednesday = CoupleFiveWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetTypeWeek() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (6): {
                            CoupleSixWednesday = CoupleSixWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetTypeWeek() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (7): {
                            CoupleSevenWednesday = CoupleSevenWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetTypeWeek() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                    }
                }

                Cell cellOneCoupleWednesday = rowfifteen.createCell(numberGroup);
                cellOneCoupleWednesday.setCellValue(CoupleOneWednesday);
                cellOneCoupleWednesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowfifteen.setHeightInPoints(HeightPoints);

                Cell cellTwoCoupleWednesday = rowSixteen.createCell(numberGroup);
                cellTwoCoupleWednesday.setCellValue(CoupleTwoWednesday);
                cellTwoCoupleWednesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowSixteen.setHeightInPoints(HeightPoints);

                Cell cellThreeCoupleWednesday = rowSeventeen.createCell(numberGroup);
                cellThreeCoupleWednesday.setCellValue(CoupleThreeWednesday);
                cellThreeCoupleWednesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowSeventeen.setHeightInPoints(HeightPoints);

                Cell cellFourCoupleWednesday = rowEighteen.createCell(numberGroup);
                cellFourCoupleWednesday.setCellValue(CoupleFourWednesday);
                cellFourCoupleWednesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowEighteen.setHeightInPoints(HeightPoints);

                Cell cellFiveCoupleWednesday = rowNineteen.createCell(numberGroup);
                cellFiveCoupleWednesday.setCellValue(CoupleFiveWednesday);
                cellFiveCoupleWednesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowNineteen.setHeightInPoints(HeightPoints);

                Cell cellSixCoupleWednesday = rowTwenty.createCell(numberGroup);
                cellSixCoupleWednesday.setCellValue(CoupleSixWednesday);
                cellSixCoupleWednesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowTwenty.setHeightInPoints(HeightPoints);

                Cell cellSevenCoupleWednesday = rowTwentyOne.createCell(numberGroup);
                cellSevenCoupleWednesday.setCellValue(CoupleSevenWednesday);
                cellSevenCoupleWednesday.setCellStyle(cellStyle);
                Groups.setColumnWidth(numberGroup, ColumnWidth);
                rowTwentyOne.setHeightInPoints(HeightPoints);

            }

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
                        CoupleOneThursday = CoupleOneThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetTypeWeek() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (2): {
                        CoupleTwoThursday = CoupleTwoThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetTypeWeek() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (3): {
                        CoupleThreeThursday = CoupleThreeThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetTypeWeek() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (4): {
                        CoupleFourThursday = CoupleFourThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetTypeWeek() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (5): {
                        CoupleFiveThursday = CoupleFiveThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetTypeWeek() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (6): {
                        CoupleSixThursday = CoupleSixThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetTypeWeek() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (7): {
                        CoupleSevenThursday = CoupleSevenThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetTypeWeek() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                }
            }

            Cell cellOneCoupleThursday = rowTwentyTwo.createCell(numberGroup);
            cellOneCoupleThursday.setCellValue(CoupleOneThursday);
            cellOneCoupleThursday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowTwentyTwo.setHeightInPoints(HeightPoints);

            Cell cellTwoCoupleThursday = rowTwentyThree.createCell(numberGroup);
            cellTwoCoupleThursday.setCellValue(CoupleTwoThursday);
            cellTwoCoupleThursday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowTwentyThree.setHeightInPoints(HeightPoints);

            Cell cellThreeCoupleThursday = rowTwentyFour.createCell(numberGroup);
            cellThreeCoupleThursday.setCellValue(CoupleThreeThursday);
            cellThreeCoupleThursday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowTwentyFour.setHeightInPoints(HeightPoints);

            Cell cellFourCoupleThursday = rowTwentyFive.createCell(numberGroup);
            cellFourCoupleThursday.setCellValue(CoupleFourThursday);
            cellFourCoupleThursday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowTwentyFive.setHeightInPoints(HeightPoints);

            Cell cellFiveCoupleThursday = rowTwentySix.createCell(numberGroup);
            cellFiveCoupleThursday.setCellValue(CoupleFiveThursday);
            cellFiveCoupleThursday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowTwentySix.setHeightInPoints(HeightPoints);

            Cell cellSixCoupleThursday = rowTwentySeven.createCell(numberGroup);
            cellSixCoupleThursday.setCellValue(CoupleSixThursday);
            cellSixCoupleThursday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowTwentySeven.setHeightInPoints(HeightPoints);

            Cell cellSevenCoupleThursday = rowTwentyEight.createCell(numberGroup);
            cellSevenCoupleThursday.setCellValue(CoupleSevenThursday);
            cellSevenCoupleThursday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowTwentyEight.setHeightInPoints(HeightPoints);

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
                        CoupleOneFriday = CoupleOneFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetTypeWeek() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (2): {
                        CoupleTwoFriday = CoupleTwoFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetTypeWeek() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (3): {
                        CoupleThreeFriday = CoupleThreeFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetTypeWeek() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (4): {
                        CoupleFourFriday = CoupleFourFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetTypeWeek() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (5): {
                        CoupleFiveFriday = CoupleFiveFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetTypeWeek() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }case (6): {
                        CoupleSixFriday = CoupleSixFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetTypeWeek() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (7): {
                        CoupleSevenFriday = CoupleSevenFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetTypeWeek() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                }
            }

            Cell cellOneCoupleFriday = rowTwentyNine.createCell(numberGroup);
            cellOneCoupleFriday.setCellValue(CoupleOneFriday);
            cellOneCoupleFriday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowTwentyNine.setHeightInPoints(HeightPoints);

            Cell cellTwoCoupleFriday = rowThirty.createCell(numberGroup);
            cellTwoCoupleFriday.setCellValue(CoupleTwoFriday);
            cellTwoCoupleFriday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowThirty.setHeightInPoints(HeightPoints);

            Cell cellThreeCoupleFriday = rowThirtyOne.createCell(numberGroup);
            cellThreeCoupleFriday.setCellValue(CoupleThreeFriday);
            cellThreeCoupleFriday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowThirtyOne.setHeightInPoints(HeightPoints);

            Cell cellFourCoupleFriday = rowThirtyTwo.createCell(numberGroup);
            cellFourCoupleFriday.setCellValue(CoupleFourFriday);
            cellFourCoupleFriday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowThirtyTwo.setHeightInPoints(HeightPoints);

            Cell cellFiveCoupleFriday = rowThirtyThree.createCell(numberGroup);
            cellFiveCoupleFriday.setCellValue(CoupleFiveFriday);
            cellFiveCoupleFriday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowThirtyThree.setHeightInPoints(HeightPoints);

            Cell cellSixCoupleFriday = rowThirtyFour.createCell(numberGroup);
            cellSixCoupleFriday.setCellValue(CoupleSixFriday);
            cellSixCoupleFriday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowThirtyFour.setHeightInPoints(HeightPoints);

            Cell cellSevenCoupleFriday = rowThirtyFive.createCell(numberGroup);
            cellSevenCoupleFriday.setCellValue(CoupleSevenFriday);
            cellSevenCoupleFriday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowThirtyFive.setHeightInPoints(HeightPoints);

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
                        CoupleOneSaturday = CoupleOneSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetTypeWeek() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (2): {
                        CoupleTwoSaturday = CoupleTwoSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetTypeWeek() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (3): {
                        CoupleThreeSaturday = CoupleThreeSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetTypeWeek() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (4): {
                        CoupleFourSaturday = CoupleFourSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetTypeWeek() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (5): {
                        CoupleFiveSaturday = CoupleFiveSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetTypeWeek() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (6): {
                        CoupleSixSaturday = CoupleSixSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetTypeWeek() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (7): {
                        CoupleSevenSaturday = CoupleSevenSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetTypeWeek() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                }
            }

            Cell cellOneCoupleSaturday = rowThirtySix.createCell(numberGroup);
            cellOneCoupleSaturday.setCellValue(CoupleOneSaturday);
            cellOneCoupleSaturday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowThirtySix.setHeightInPoints(HeightPoints);

            Cell cellTwoCoupleSaturday = rowThirtySeven.createCell(numberGroup);
            cellTwoCoupleSaturday.setCellValue(CoupleTwoSaturday);
            cellTwoCoupleSaturday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowThirtySeven.setHeightInPoints(HeightPoints);

            Cell cellThreeCoupleSaturday = rowThirtyEight.createCell(numberGroup);
            cellThreeCoupleSaturday.setCellValue(CoupleThreeSaturday);
            cellThreeCoupleSaturday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowThirtyEight.setHeightInPoints(HeightPoints);

            Cell cellFourCoupleSaturday = rowThirtyNine.createCell(numberGroup);
            cellFourCoupleSaturday.setCellValue(CoupleFourSaturday);
            cellFourCoupleSaturday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowThirtyNine.setHeightInPoints(HeightPoints);

            Cell cellFiveCoupleSaturday = rowForty.createCell(numberGroup);
            cellFiveCoupleSaturday.setCellValue(CoupleFiveSaturday);
            cellFiveCoupleSaturday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowForty.setHeightInPoints(HeightPoints);

            Cell cellSixCoupleSaturday = rowFortyOne.createCell(numberGroup);
            cellSixCoupleSaturday.setCellValue(CoupleSixSaturday);
            cellSixCoupleSaturday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowFortyOne.setHeightInPoints(HeightPoints);

            Cell cellSevenCoupleSaturday = rowFortyTwo.createCell(numberGroup);
            cellSevenCoupleSaturday.setCellValue(CoupleSevenSaturday);
            cellSevenCoupleSaturday.setCellStyle(cellStyle);
            Groups.setColumnWidth(numberGroup, ColumnWidth);
            rowFortyTwo.setHeightInPoints(HeightPoints);

            cell = rowZero.createCell(numberGroup);
            cell.setCellValue(GroupName);
            numberGroup++;

        }

        FileOutputStream fileOutputStream = new FileOutputStream("AllGroupExelDoc");
        workbookGroups.write(fileOutputStream);
        workbookGroups.close();

    }

}
