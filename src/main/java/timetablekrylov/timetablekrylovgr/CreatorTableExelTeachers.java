package timetablekrylov.timetablekrylovgr;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class CreatorTableExelTeachers {

    int HeightPoints = 250;
    int ColumnWidth = 10000;

    private Workbook workbookTeachers = new HSSFWorkbook();

    private Sheet Teachers = workbookTeachers.createSheet("Преподаватели");

    private Row rowZero = Teachers.createRow(0);
    private Cell cellDayWeek = rowZero.createCell(0);
    private Cell CoupleNumber = rowZero.createCell(1);

    private Row rowOne = Teachers.createRow(1);
    private Cell Monday = rowOne.createCell(0);
    private Cell CoupleOneMonday = rowOne.createCell(1);

    private Row rowTwo = Teachers.createRow(2);
    private Cell CoupleTwoMonday = rowTwo.createCell(1);

    private Row rowThree = Teachers.createRow(3);
    private Cell CoupleThreeMonday = rowThree.createCell(1);

    private Row rowFour = Teachers.createRow(4);
    private Cell CoupleFourMonday = rowFour.createCell(1);

    private Row rowFive = Teachers.createRow(5);
    private Cell CoupleFiveMonday = rowFive.createCell(1);

    private Row rowSix = Teachers.createRow(6);
    private Cell CoupleSixMonday = rowSix.createCell(1);

    private Row rowSeven = Teachers.createRow(7);
    private Cell CoupleSevenMonday = rowSeven.createCell(1);

    // -------------------------------------------------------------------------------------

    private Row rowEight = Teachers.createRow(8);
    private Cell Tuesday = rowEight.createCell(0);
    private Cell CoupleOneTuesday = rowEight.createCell(1);

    private Row rowNine = Teachers.createRow(9);
    private Cell CoupleTwoTuesday = rowNine.createCell(1);

    private Row rowTen = Teachers.createRow(10);
    private Cell CoupleThreeTuesday = rowTen.createCell(1);

    private Row rowEleven = Teachers.createRow(11);
    private Cell CoupleFourTuesday = rowEleven.createCell(1);

    private Row rowTwelve = Teachers.createRow(12);
    private Cell CoupleFiveTuesday = rowTwelve.createCell(1);

    private Row rowThirteen = Teachers.createRow(13);
    private Cell CoupleSixTuesday = rowThirteen.createCell(1);

    private Row rowFourteen = Teachers.createRow(14);
    private Cell CoupleSevenTuesday = rowFourteen.createCell(1);

    // --------------------------------------------------------------------------------------

    private Row rowfifteen = Teachers.createRow(15);
    private Cell Wednesday = rowfifteen.createCell(0);
    private Cell CoupleOneWednesday = rowfifteen.createCell(1);

    private Row rowSixteen = Teachers.createRow(16);
    private Cell CoupleTwoWednesday = rowSixteen.createCell(1);

    private Row rowSeventeen = Teachers.createRow(17);
    private Cell CoupleThreeWednesday = rowSeventeen.createCell(1);

    private Row rowEighteen = Teachers.createRow(18);
    private Cell CoupleFourWednesday = rowEighteen.createCell(1);

    private Row rowNineteen = Teachers.createRow(19);
    private Cell CoupleFiveWednesday = rowNineteen.createCell(1);

    private Row rowTwenty = Teachers.createRow(20);
    private Cell CoupleSixWednesday = rowTwenty.createCell(1);

    private Row rowTwentyOne = Teachers.createRow(21);
    private Cell CoupleSevenWednesday = rowTwentyOne.createCell(1);

    // ------------------------------------------------------------------------------------------

    private Row rowTwentyTwo = Teachers.createRow(22);
    private Cell Thursday = rowTwentyTwo.createCell(0);
    private Cell CoupleOneThursday = rowTwentyTwo.createCell(1);

    private Row rowTwentyThree = Teachers.createRow(23);
    private Cell CoupleTwoThursday = rowTwentyThree.createCell(1);

    private Row rowTwentyFour = Teachers.createRow(24);
    private Cell CoupleThreeThursday = rowTwentyFour.createCell(1);

    private Row rowTwentyFive = Teachers.createRow(25);
    private Cell CoupleFourThursday = rowTwentyFive.createCell(1);

    private Row rowTwentySix = Teachers.createRow(26);
    private Cell CoupleFiveThursday = rowTwentySix.createCell(1);

    private Row rowTwentySeven = Teachers.createRow(27);
    private Cell CoupleSixThursday = rowTwentySeven.createCell(1);

    private Row rowTwentyEight = Teachers.createRow(28);
    private Cell CoupleSevenThursday = rowTwentyEight.createCell(1);

    // --------------------------------------------------------------------------------------------

    private Row rowTwentyNine = Teachers.createRow(29);
    private Cell Friday = rowTwentyNine.createCell(0);
    private Cell CoupleOneFriday = rowTwentyNine.createCell(1);

    private Row rowThirty = Teachers.createRow(30);
    private Cell CoupleTwoFriday = rowThirty.createCell(1);

    private Row rowThirtyOne = Teachers.createRow(31);
    private Cell CoupleThreeFriday = rowThirtyOne.createCell(1);

    private Row rowThirtyTwo = Teachers.createRow(32);
    private Cell CoupleFourFriday = rowThirtyTwo.createCell(1);

    private Row rowThirtyThree = Teachers.createRow(33);
    private Cell CoupleFiveFriday = rowThirtyThree.createCell(1);

    private Row rowThirtyFour = Teachers.createRow(34);
    private Cell CoupleSixFriday = rowThirtyFour.createCell(1);

    private Row rowThirtyFive = Teachers.createRow(35);
    private Cell CoupleSevenFriday = rowThirtyFive.createCell(1);

    // ---------------------------------------------------------------------------

    private Row rowThirtySix = Teachers.createRow(36);
    private Cell Saturday = rowThirtySix.createCell(0);
    private Cell CoupleOneSaturday = rowThirtySix.createCell(1);

    private Row rowThirtySeven = Teachers.createRow(37);
    private Cell CoupleTwoSaturday = rowThirtySeven.createCell(1);

    private Row rowThirtyEight = Teachers.createRow(38);
    private Cell CoupleThreeSaturday = rowThirtyEight.createCell(1);

    private Row rowThirtyNine = Teachers.createRow(39);
    private Cell CoupleFourSaturday = rowThirtyNine.createCell(1);

    private Row rowForty = Teachers.createRow(40);
    private Cell CoupleFiveSaturday = rowForty.createCell(1);

    private Row rowFortyOne = Teachers.createRow(41);
    private Cell CoupleSixSaturday = rowFortyOne.createCell(1);

    private Row rowFortyTwo = Teachers.createRow(42);
    private Cell CoupleSevenSaturday = rowFortyTwo.createCell(1);

    public void CreateTimeTableTeachers(ArrayList<Teacher> ArrayTeacher) throws IOException {

        CellStyle cellStyle = workbookTeachers.createCellStyle();
        cellStyle.setWrapText(true);

        cellDayWeek.setCellStyle(cellStyle);
        cellDayWeek.setCellValue("День недели");
        Teachers.setColumnWidth(0,8000);
        rowZero.setHeightInPoints(30);

        CoupleNumber.setCellStyle(cellStyle);
        CoupleNumber.setCellValue("Номер пары");
        Teachers.setColumnWidth(1,8000);
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

        Teachers.addMergedRegion(new CellRangeAddress(36,42,0,0));
        Teachers.addMergedRegion(new CellRangeAddress(29,35,0,0));
        Teachers.addMergedRegion(new CellRangeAddress(22,28,0,0));
        Teachers.addMergedRegion(new CellRangeAddress(15,21,0,0));
        Teachers.addMergedRegion(new CellRangeAddress(8,14,0,0));
        Teachers.addMergedRegion(new CellRangeAddress(1,7,0,0));

        Cell cell;
        int numberTeacher = 2;

        for(int i = 0; i < ArrayTeacher.size(); i++) {

            String TeacherName = ArrayTeacher.get(i).GetTeacherName();

            ArrayList<CoupleTeacher> ArrayMonday = new ArrayList<>();
            ArrayList<CoupleTeacher> ArrayTuesday = new ArrayList<>();
            ArrayList<CoupleTeacher> ArrayWednesday = new ArrayList<>();
            ArrayList<CoupleTeacher> ArrayThursday = new ArrayList<>();
            ArrayList<CoupleTeacher> ArrayFriday = new ArrayList<>();
            ArrayList<CoupleTeacher> ArraySaturday = new ArrayList<>();

            ArrayList<CoupleTeacher> ArrayCouple = ArrayTeacher.get(i).GetArrayCoupleTeacher();

            for(int index = 0; index < ArrayCouple.size(); index++){

                int IdDay = ArrayCouple.get(index).GetIdDay();

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

                    switch (ArrayMonday.get(index2).GetNumberCouple()) {
                        case (1): {
                            CoupleOneMonday = CoupleOneMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetGroupName() + " " + ArrayMonday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (2): {
                            CoupleTwoMonday = CoupleTwoMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetGroupName() + " " + ArrayMonday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (3): {
                            CoupleThreeMonday = CoupleThreeMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetGroupName() + " " + ArrayMonday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (4): {
                            CoupleFourMonday = CoupleFourMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetGroupName() + " " + ArrayMonday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (5): {
                            CoupleFiveMonday = CoupleFiveMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetGroupName() + " " + ArrayMonday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (6): {
                            CoupleSixMonday = CoupleSixMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetGroupName() + " " + ArrayMonday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (7): {
                            CoupleSevenMonday = CoupleSevenMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetGroupName() + " " + ArrayMonday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                    }
                }

                Cell cellOneCoupleMonday = rowOne.createCell(numberTeacher);
                cellOneCoupleMonday.setCellValue(CoupleOneMonday);
                cellOneCoupleMonday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
                rowOne.setHeightInPoints(HeightPoints);

                Cell cellTwoCoupleMonday = rowTwo.createCell(numberTeacher);
                cellTwoCoupleMonday.setCellValue(CoupleTwoMonday);
                cellTwoCoupleMonday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
                rowTwo.setHeightInPoints(HeightPoints);

                Cell cellThreeCoupleMonday = rowThree.createCell(numberTeacher);
                cellThreeCoupleMonday.setCellValue(CoupleThreeMonday);
                cellThreeCoupleMonday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
                rowThree.setHeightInPoints(HeightPoints);

                Cell cellFourCoupleMonday = rowFour.createCell(numberTeacher);
                cellFourCoupleMonday.setCellValue(CoupleFourMonday);
                cellFourCoupleMonday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
                rowFour.setHeightInPoints(HeightPoints);

                Cell cellFiveCoupleMonday = rowFive.createCell(numberTeacher);
                cellFiveCoupleMonday.setCellValue(CoupleFiveMonday);
                cellFiveCoupleMonday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
                rowFive.setHeightInPoints(HeightPoints);

                Cell cellSixCoupleMonday = rowSix.createCell(numberTeacher);
                cellSixCoupleMonday.setCellValue(CoupleSixMonday);
                cellSixCoupleMonday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
                rowSix.setHeightInPoints(HeightPoints);

                Cell cellSevenCoupleMonday = rowSeven.createCell(numberTeacher);
                cellSevenCoupleMonday.setCellValue(CoupleSevenMonday);
                cellSevenCoupleMonday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
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

                    switch (ArrayTuesday.get(index2).GetNumberCouple()) {
                        case (1): {
                            CoupleOneTuesday = CoupleOneTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetTypeWeek() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetGroupName() + " " + ArrayTuesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (2): {
                            CoupleTwoTuesday = CoupleTwoTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetTypeWeek() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetGroupName() + " " + ArrayTuesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (3): {
                            CoupleThreeTuesday = CoupleThreeTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetTypeWeek() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetGroupName() + " " + ArrayTuesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (4): {
                            CoupleFourTuesday = CoupleFourTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetTypeWeek() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetGroupName() + " " + ArrayTuesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (5): {
                            CoupleFiveTuesday = CoupleFiveTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetTypeWeek() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetGroupName() + " " + ArrayTuesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (6): {
                            CoupleSixTuesday = CoupleSixTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetTypeWeek() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetGroupName() + " " + ArrayTuesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (7): {
                            CoupleSevenTuesday = CoupleSevenTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetTypeWeek() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetGroupName() + " " + ArrayTuesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                    }
                }

                Cell cellOneCoupleTuesday = rowEight.createCell(numberTeacher);
                cellOneCoupleTuesday.setCellValue(CoupleOneTuesday);
                cellOneCoupleTuesday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
                rowEight.setHeightInPoints(HeightPoints);

                Cell cellTwoCoupleTuesday = rowNine.createCell(numberTeacher);
                cellTwoCoupleTuesday.setCellValue(CoupleTwoTuesday);
                cellTwoCoupleTuesday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
                rowNine.setHeightInPoints(HeightPoints);

                Cell cellThreeCoupleTuesday = rowTen.createCell(numberTeacher);
                cellThreeCoupleTuesday.setCellValue(CoupleThreeTuesday);
                cellThreeCoupleTuesday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
                rowTen.setHeightInPoints(HeightPoints);

                Cell cellFourCoupleTuesday = rowEleven.createCell(numberTeacher);
                cellFourCoupleTuesday.setCellValue(CoupleFourTuesday);
                cellFourCoupleTuesday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
                rowEleven.setHeightInPoints(HeightPoints);

                Cell cellFiveCoupleTuesday = rowTwelve.createCell(numberTeacher);
                cellFiveCoupleTuesday.setCellValue(CoupleFiveTuesday);
                cellFiveCoupleTuesday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
                rowTwelve.setHeightInPoints(HeightPoints);

                Cell cellSixCoupleTuesday = rowThirteen.createCell(numberTeacher);
                cellSixCoupleTuesday.setCellValue(CoupleSixTuesday);
                cellSixCoupleTuesday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
                rowThirteen.setHeightInPoints(HeightPoints);

                Cell cellSevenCoupleTuesday = rowFourteen.createCell(numberTeacher);
                cellSevenCoupleTuesday.setCellValue(CoupleSevenTuesday);
                cellSevenCoupleTuesday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
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

                    switch (ArrayWednesday.get(index2).GetNumberCouple()) {
                        case (1): {
                            CoupleOneWednesday = CoupleOneWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetTypeWeek() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetGroupName() + " " + ArrayWednesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (2): {
                            CoupleTwoWednesday = CoupleTwoWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetTypeWeek() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetGroupName() + " " + ArrayWednesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (3): {
                            CoupleThreeWednesday = CoupleThreeWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetTypeWeek() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetGroupName() + " " + ArrayWednesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (4): {
                            CoupleFourWednesday = CoupleFourWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetTypeWeek() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetGroupName() + " " + ArrayWednesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (5): {
                            CoupleFiveWednesday = CoupleFiveWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetTypeWeek() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetGroupName() + " " + ArrayWednesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (6): {
                            CoupleSixWednesday = CoupleSixWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetTypeWeek() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetGroupName() + " " + ArrayWednesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                        case (7): {
                            CoupleSevenWednesday = CoupleSevenWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetTypeWeek() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetGroupName() + " " + ArrayWednesday.get(index2).GetAud() + "\n" + "\n";
                            break;
                        }
                    }
                }

                Cell cellOneCoupleWednesday = rowfifteen.createCell(numberTeacher);
                cellOneCoupleWednesday.setCellValue(CoupleOneWednesday);
                cellOneCoupleWednesday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
                rowfifteen.setHeightInPoints(HeightPoints);

                Cell cellTwoCoupleWednesday = rowSixteen.createCell(numberTeacher);
                cellTwoCoupleWednesday.setCellValue(CoupleTwoWednesday);
                cellTwoCoupleWednesday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
                rowSixteen.setHeightInPoints(HeightPoints);

                Cell cellThreeCoupleWednesday = rowSeventeen.createCell(numberTeacher);
                cellThreeCoupleWednesday.setCellValue(CoupleThreeWednesday);
                cellThreeCoupleWednesday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
                rowSeventeen.setHeightInPoints(HeightPoints);

                Cell cellFourCoupleWednesday = rowEighteen.createCell(numberTeacher);
                cellFourCoupleWednesday.setCellValue(CoupleFourWednesday);
                cellFourCoupleWednesday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
                rowEighteen.setHeightInPoints(HeightPoints);

                Cell cellFiveCoupleWednesday = rowNineteen.createCell(numberTeacher);
                cellFiveCoupleWednesday.setCellValue(CoupleFiveWednesday);
                cellFiveCoupleWednesday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
                rowNineteen.setHeightInPoints(HeightPoints);

                Cell cellSixCoupleWednesday = rowTwenty.createCell(numberTeacher);
                cellSixCoupleWednesday.setCellValue(CoupleSixWednesday);
                cellSixCoupleWednesday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
                rowTwenty.setHeightInPoints(HeightPoints);

                Cell cellSevenCoupleWednesday = rowTwentyOne.createCell(numberTeacher);
                cellSevenCoupleWednesday.setCellValue(CoupleSevenWednesday);
                cellSevenCoupleWednesday.setCellStyle(cellStyle);
                Teachers.setColumnWidth(numberTeacher, ColumnWidth);
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

                switch (ArrayThursday.get(index2).GetNumberCouple()) {
                    case (1): {
                        CoupleOneThursday = CoupleOneThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetTypeWeek() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetGroupName() + " " + ArrayThursday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (2): {
                        CoupleTwoThursday = CoupleTwoThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetTypeWeek() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetGroupName() + " " + ArrayThursday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (3): {
                        CoupleThreeThursday = CoupleThreeThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetTypeWeek() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetGroupName() + " " + ArrayThursday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (4): {
                        CoupleFourThursday = CoupleFourThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetTypeWeek() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetGroupName() + " " + ArrayThursday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (5): {
                        CoupleFiveThursday = CoupleFiveThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetTypeWeek() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetGroupName() + " " + ArrayThursday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (6): {
                        CoupleSixThursday = CoupleSixThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetTypeWeek() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetGroupName() + " " + ArrayThursday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (7): {
                        CoupleSevenThursday = CoupleSevenThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetTypeWeek() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetGroupName() + " " + ArrayThursday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                }
            }

            Cell cellOneCoupleThursday = rowTwentyTwo.createCell(numberTeacher);
            cellOneCoupleThursday.setCellValue(CoupleOneThursday);
            cellOneCoupleThursday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowTwentyTwo.setHeightInPoints(HeightPoints);

            Cell cellTwoCoupleThursday = rowTwentyThree.createCell(numberTeacher);
            cellTwoCoupleThursday.setCellValue(CoupleTwoThursday);
            cellTwoCoupleThursday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowTwentyThree.setHeightInPoints(HeightPoints);

            Cell cellThreeCoupleThursday = rowTwentyFour.createCell(numberTeacher);
            cellThreeCoupleThursday.setCellValue(CoupleThreeThursday);
            cellThreeCoupleThursday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowTwentyFour.setHeightInPoints(HeightPoints);

            Cell cellFourCoupleThursday = rowTwentyFive.createCell(numberTeacher);
            cellFourCoupleThursday.setCellValue(CoupleFourThursday);
            cellFourCoupleThursday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowTwentyFive.setHeightInPoints(HeightPoints);

            Cell cellFiveCoupleThursday = rowTwentySix.createCell(numberTeacher);
            cellFiveCoupleThursday.setCellValue(CoupleFiveThursday);
            cellFiveCoupleThursday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowTwentySix.setHeightInPoints(HeightPoints);

            Cell cellSixCoupleThursday = rowTwentySeven.createCell(numberTeacher);
            cellSixCoupleThursday.setCellValue(CoupleSixThursday);
            cellSixCoupleThursday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowTwentySeven.setHeightInPoints(HeightPoints);

            Cell cellSevenCoupleThursday = rowTwentyEight.createCell(numberTeacher);
            cellSevenCoupleThursday.setCellValue(CoupleSevenThursday);
            cellSevenCoupleThursday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
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

                switch (ArrayFriday.get(index2).GetNumberCouple()) {
                    case (1): {
                        CoupleOneFriday = CoupleOneFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetTypeWeek() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetGroupName() + " " + ArrayFriday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (2): {
                        CoupleTwoFriday = CoupleTwoFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetTypeWeek() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetGroupName() + " " + ArrayFriday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (3): {
                        CoupleThreeFriday = CoupleThreeFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetTypeWeek() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetGroupName() + " " + ArrayFriday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (4): {
                        CoupleFourFriday = CoupleFourFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetTypeWeek() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetGroupName() + " " + ArrayFriday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (5): {
                        CoupleFiveFriday = CoupleFiveFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetTypeWeek() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetGroupName() + " " + ArrayFriday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }case (6): {
                        CoupleSixFriday = CoupleSixFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetTypeWeek() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetGroupName() + " " + ArrayFriday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (7): {
                        CoupleSevenFriday = CoupleSevenFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetTypeWeek() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetGroupName() + " " + ArrayFriday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                }
            }

            Cell cellOneCoupleFriday = rowTwentyNine.createCell(numberTeacher);
            cellOneCoupleFriday.setCellValue(CoupleOneFriday);
            cellOneCoupleFriday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowTwentyNine.setHeightInPoints(HeightPoints);

            Cell cellTwoCoupleFriday = rowThirty.createCell(numberTeacher);
            cellTwoCoupleFriday.setCellValue(CoupleTwoFriday);
            cellTwoCoupleFriday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowThirty.setHeightInPoints(HeightPoints);

            Cell cellThreeCoupleFriday = rowThirtyOne.createCell(numberTeacher);
            cellThreeCoupleFriday.setCellValue(CoupleThreeFriday);
            cellThreeCoupleFriday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowThirtyOne.setHeightInPoints(HeightPoints);

            Cell cellFourCoupleFriday = rowThirtyTwo.createCell(numberTeacher);
            cellFourCoupleFriday.setCellValue(CoupleFourFriday);
            cellFourCoupleFriday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowThirtyTwo.setHeightInPoints(HeightPoints);

            Cell cellFiveCoupleFriday = rowThirtyThree.createCell(numberTeacher);
            cellFiveCoupleFriday.setCellValue(CoupleFiveFriday);
            cellFiveCoupleFriday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowThirtyThree.setHeightInPoints(HeightPoints);

            Cell cellSixCoupleFriday = rowThirtyFour.createCell(numberTeacher);
            cellSixCoupleFriday.setCellValue(CoupleSixFriday);
            cellSixCoupleFriday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowThirtyFour.setHeightInPoints(HeightPoints);

            Cell cellSevenCoupleFriday = rowThirtyFive.createCell(numberTeacher);
            cellSevenCoupleFriday.setCellValue(CoupleSevenFriday);
            cellSevenCoupleFriday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
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

                switch (ArraySaturday.get(index2).GetNumberCouple()) {
                    case (1): {
                        CoupleOneSaturday = CoupleOneSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetTypeWeek() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetGroupName() + " " + ArraySaturday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (2): {
                        CoupleTwoSaturday = CoupleTwoSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetTypeWeek() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetGroupName() + " " + ArraySaturday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (3): {
                        CoupleThreeSaturday = CoupleThreeSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetTypeWeek() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetGroupName() + " " + ArraySaturday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (4): {
                        CoupleFourSaturday = CoupleFourSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetTypeWeek() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetGroupName() + " " + ArraySaturday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (5): {
                        CoupleFiveSaturday = CoupleFiveSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetTypeWeek() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetGroupName() + " " + ArraySaturday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (6): {
                        CoupleSixSaturday = CoupleSixSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetTypeWeek() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetGroupName() + " " + ArraySaturday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                    case (7): {
                        CoupleSevenSaturday = CoupleSevenSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetTypeWeek() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetGroupName() + " " + ArraySaturday.get(index2).GetAud() + "\n" + "\n";
                        break;
                    }
                }
            }

            Cell cellOneCoupleSaturday = rowThirtySix.createCell(numberTeacher);
            cellOneCoupleSaturday.setCellValue(CoupleOneSaturday);
            cellOneCoupleSaturday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowThirtySix.setHeightInPoints(HeightPoints);

            Cell cellTwoCoupleSaturday = rowThirtySeven.createCell(numberTeacher);
            cellTwoCoupleSaturday.setCellValue(CoupleTwoSaturday);
            cellTwoCoupleSaturday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowThirtySeven.setHeightInPoints(HeightPoints);

            Cell cellThreeCoupleSaturday = rowThirtyEight.createCell(numberTeacher);
            cellThreeCoupleSaturday.setCellValue(CoupleThreeSaturday);
            cellThreeCoupleSaturday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowThirtyEight.setHeightInPoints(HeightPoints);

            Cell cellFourCoupleSaturday = rowThirtyNine.createCell(numberTeacher);
            cellFourCoupleSaturday.setCellValue(CoupleFourSaturday);
            cellFourCoupleSaturday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowThirtyNine.setHeightInPoints(HeightPoints);

            Cell cellFiveCoupleSaturday = rowForty.createCell(numberTeacher);
            cellFiveCoupleSaturday.setCellValue(CoupleFiveSaturday);
            cellFiveCoupleSaturday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowForty.setHeightInPoints(HeightPoints);

            Cell cellSixCoupleSaturday = rowFortyOne.createCell(numberTeacher);
            cellSixCoupleSaturday.setCellValue(CoupleSixSaturday);
            cellSixCoupleSaturday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowFortyOne.setHeightInPoints(HeightPoints);

            Cell cellSevenCoupleSaturday = rowFortyTwo.createCell(numberTeacher);
            cellSevenCoupleSaturday.setCellValue(CoupleSevenSaturday);
            cellSevenCoupleSaturday.setCellStyle(cellStyle);
            Teachers.setColumnWidth(numberTeacher, ColumnWidth);
            rowFortyTwo.setHeightInPoints(HeightPoints);

            cell = rowZero.createCell(numberTeacher);
            cell.setCellValue(TeacherName);
            numberTeacher++;

        }

        FileOutputStream fileOutputStream = new FileOutputStream("AllTeachersExelDoc");

        workbookTeachers.write(fileOutputStream);
        workbookTeachers.close();

    }

}
