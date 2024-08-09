package timetablekrylov.timetablekrylovgr;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class CreatorTableExelGroup {

    ArrayList<Cell> ArrayCellGroupOne = new ArrayList<>();

    private Workbook workbookOneElementGroup = new HSSFWorkbook();

    private Sheet Group = workbookOneElementGroup.createSheet("Группа");

    private Row rowGroupZero = Group.createRow(0);
    private Cell cellDayWeek = rowGroupZero.createCell(0);

    private Cell cellGroupMonday = rowGroupZero.createCell(1);
    private Cell cellGroupTuesday = rowGroupZero.createCell(2);
    private Cell cellGroupWednesday = rowGroupZero.createCell(3);
    private Cell cellGroupThursday = rowGroupZero.createCell(4);
    private Cell cellGroupFriday = rowGroupZero.createCell(5);
    private Cell cellGroupSaturday = rowGroupZero.createCell(6);

    private Row rowGroupOne = Group.createRow(1);
    private Cell cellOneCouple = rowGroupOne.createCell(0);

    private Row rowGroupTwo = Group.createRow(2);
    private Cell cellTwoCouple = rowGroupTwo.createCell(0);

    private Row rowGroupThree = Group.createRow(3);
    private Cell cellThreeCouple = rowGroupThree.createCell(0);

    private Row rowGroupFour = Group.createRow(4);
    private Cell cellFourCouple = rowGroupFour.createCell(0);

    private Row rowGroupFive = Group.createRow(5);
    private Cell cellFiveCouple = rowGroupFive.createCell(0);

    private Row rowGroupSix = Group.createRow(6);
    private Cell cellSixCouple = rowGroupSix.createCell(0);

    private Row rowGroupSeven = Group.createRow(7);
    private Cell cellSevenCouple = rowGroupSeven.createCell(0);


    public void CreatorTimeTableOneGroup(Group group) throws IOException {

        int HeightPoints = 250;
        int ColumnWidth = 10000;

        CellStyle cellStyle = workbookOneElementGroup.createCellStyle();
        cellStyle.setWrapText(true);

        ArrayList<CoupleGroup> ArrayMonday = new ArrayList<>();
        ArrayList<CoupleGroup> ArrayTuesday = new ArrayList<>();
        ArrayList<CoupleGroup> ArrayWednesday = new ArrayList<>();
        ArrayList<CoupleGroup> ArrayThursday = new ArrayList<>();
        ArrayList<CoupleGroup> ArrayFriday = new ArrayList<>();
        ArrayList<CoupleGroup> ArraySaturday = new ArrayList<>();


        ArrayList<CoupleGroup> ArrayCouple = group.GetArrayCouples();

        cellDayWeek.setCellValue("День недели" + "\n" + "Номер пары");
        cellDayWeek.setCellStyle(cellStyle);
        CellUtil.setAlignment(cellDayWeek,HorizontalAlignment.CENTER);
        Group.setColumnWidth(0,8000);
        rowGroupZero.setHeightInPoints(30);

        cellOneCouple.setCellValue("1 пара");
        cellTwoCouple.setCellValue("2 пара");
        cellThreeCouple.setCellValue("3 пара");
        cellFourCouple.setCellValue("4 пара");
        cellFiveCouple.setCellValue("5 пара");
        cellSixCouple.setCellValue("6 пара");
        cellSevenCouple.setCellValue("7 пара");

        cellGroupMonday.setCellValue("Понедельник");
        Group.setColumnWidth(1,8000);

        cellGroupTuesday.setCellValue("Вторник");
        Group.setColumnWidth(2,8000);

        cellGroupWednesday.setCellValue("Среда");
        Group.setColumnWidth(3,8000);

        cellGroupThursday.setCellValue("Четверг");
        Group.setColumnWidth(4,8000);

        cellGroupFriday.setCellValue("Пятница");
        Group.setColumnWidth(5,8000);

        cellGroupSaturday.setCellValue("Суббота");
        Group.setColumnWidth(6,8000);

        for(int i = 0; i < ArrayCouple.size(); i++){

            int IdDay = ArrayCouple.get(i).GetIDDay();

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

        // Cell Monday

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
                    CoupleOneMonday = CoupleOneMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetTypeWeek() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetTeacherName() + " " + ArrayMonday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (2): {
                    CoupleTwoMonday = CoupleTwoMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetTypeWeek() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetTeacherName() + " " + ArrayMonday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (3): {
                    CoupleThreeMonday = CoupleThreeMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetTypeWeek() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetTeacherName() + " " + ArrayMonday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (4): {
                    CoupleFourMonday = CoupleFourMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetTypeWeek() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetTeacherName() + " " + ArrayMonday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (5): {
                    CoupleFiveMonday = CoupleFiveMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetTypeWeek() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetTeacherName() + " " + ArrayMonday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (6): {
                    CoupleSixMonday = CoupleSixMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetTypeWeek() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetTeacherName() + " " + ArrayMonday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (7): {
                    CoupleSevenMonday = CoupleSevenMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetTypeWeek() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetTeacherName() + " " + ArrayMonday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
            }
        }


        Cell cellOneCoupleMonday = rowGroupOne.createCell(1);
        cellOneCoupleMonday.setCellValue(CoupleOneMonday);
        cellOneCoupleMonday.setCellStyle(cellStyle);
        Group.setColumnWidth(1, ColumnWidth);
        rowGroupOne.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleMonday = rowGroupTwo.createCell(1);
        cellTwoCoupleMonday.setCellValue(CoupleTwoMonday);
        cellTwoCoupleMonday.setCellStyle(cellStyle);
        Group.setColumnWidth(1, ColumnWidth);
        rowGroupTwo.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleMonday = rowGroupThree.createCell(1);
        cellThreeCoupleMonday.setCellValue(CoupleThreeMonday);
        cellThreeCoupleMonday.setCellStyle(cellStyle);
        Group.setColumnWidth(1, ColumnWidth);
        rowGroupThree.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleMonday = rowGroupFour.createCell(1);
        cellFourCoupleMonday.setCellValue(CoupleFourMonday);
        cellFourCoupleMonday.setCellStyle(cellStyle);
        Group.setColumnWidth(1, ColumnWidth);
        rowGroupFour.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleMonday = rowGroupFive.createCell(1);
        cellFiveCoupleMonday.setCellValue(CoupleFiveMonday);
        cellFiveCoupleMonday.setCellStyle(cellStyle);
        Group.setColumnWidth(1, ColumnWidth);
        rowGroupFive.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleMonday = rowGroupSix.createCell(1);
        cellSixCoupleMonday.setCellValue(CoupleSixMonday);
        cellSixCoupleMonday.setCellStyle(cellStyle);
        Group.setColumnWidth(1, ColumnWidth);
        rowGroupSix.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleMonday = rowGroupSeven.createCell(1);
        cellSevenCoupleMonday.setCellValue(CoupleSevenMonday);
        cellSevenCoupleMonday.setCellStyle(cellStyle);
        Group.setColumnWidth(1, ColumnWidth);
        rowGroupSeven.setHeightInPoints(HeightPoints);

        // Cell Tuesday

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
                    CoupleOneTuesday = CoupleOneTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetTypeWeek() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetTeacherName() + " " + ArrayTuesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (2): {
                    CoupleTwoTuesday = CoupleTwoTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetTypeWeek() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetTeacherName() + " " + ArrayTuesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (3): {
                    CoupleThreeTuesday = CoupleThreeTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetTypeWeek() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetTeacherName() + " " + ArrayTuesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (4): {
                    CoupleFourTuesday = CoupleFourTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetTypeWeek() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetTeacherName() + " " + ArrayTuesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (5): {
                    CoupleFiveTuesday = CoupleFiveTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetTypeWeek() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetTeacherName() + " " + ArrayTuesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (6): {
                    CoupleSixTuesday = CoupleSixTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetTypeWeek() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetTeacherName() + " " + ArrayTuesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (7): {
                    CoupleSevenTuesday = CoupleSevenTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetTypeWeek() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetTeacherName() + " " + ArrayTuesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
            }
        }

        Cell cellOneCoupleTuesday = rowGroupOne.createCell(2);
        cellOneCoupleTuesday.setCellValue(CoupleOneTuesday);
        cellOneCoupleTuesday.setCellStyle(cellStyle);
        Group.setColumnWidth(2, ColumnWidth);
        rowGroupOne.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleTuesday = rowGroupTwo.createCell(2);
        cellTwoCoupleTuesday.setCellValue(CoupleTwoTuesday);
        cellTwoCoupleTuesday.setCellStyle(cellStyle);
        Group.setColumnWidth(2, ColumnWidth);
        rowGroupTwo.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleTuesday = rowGroupThree.createCell(2);
        cellThreeCoupleTuesday.setCellValue(CoupleThreeTuesday);
        cellThreeCoupleTuesday.setCellStyle(cellStyle);
        Group.setColumnWidth(2, ColumnWidth);
        rowGroupThree.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleTuesday = rowGroupFour.createCell(2);
        cellFourCoupleTuesday.setCellValue(CoupleFourTuesday);
        cellFourCoupleTuesday.setCellStyle(cellStyle);
        Group.setColumnWidth(2, ColumnWidth);
        rowGroupFour.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleTuesday = rowGroupFive.createCell(2);
        cellFiveCoupleTuesday.setCellValue(CoupleFiveTuesday);
        cellFiveCoupleTuesday.setCellStyle(cellStyle);
        Group.setColumnWidth(2, ColumnWidth);
        rowGroupFive.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleTuesday = rowGroupSix.createCell(2);
        cellSixCoupleTuesday.setCellValue(CoupleSixTuesday);
        cellSixCoupleTuesday.setCellStyle(cellStyle);
        Group.setColumnWidth(2, ColumnWidth);
        rowGroupSix.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleTuesday = rowGroupSeven.createCell(2);
        cellSevenCoupleTuesday.setCellValue(CoupleSevenTuesday);
        cellSevenCoupleTuesday.setCellStyle(cellStyle);
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
                    CoupleOneWednesday = CoupleOneWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetTypeWeek() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetTeacherName() + " " + ArrayWednesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (2): {
                    CoupleTwoWednesday = CoupleTwoWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetTypeWeek() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetTeacherName() + " " + ArrayWednesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (3): {
                    CoupleThreeWednesday = CoupleThreeWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetTypeWeek() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetTeacherName() + " " + ArrayWednesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (4): {
                    CoupleFourWednesday = CoupleFourWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetTypeWeek() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetTeacherName() + " " + ArrayWednesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (5): {
                    CoupleFiveWednesday = CoupleFiveWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetTypeWeek() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetTeacherName() + " " + ArrayWednesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (6): {
                    CoupleSixWednesday = CoupleSixWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetTypeWeek() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetTeacherName() + " " + ArrayWednesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (7): {
                    CoupleSevenWednesday = CoupleSevenWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetTypeWeek() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetTeacherName() + " " + ArrayWednesday.get(i).GetAud() + "\n" + "\n";
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
                    CoupleOneThursday = CoupleOneThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetTypeWeek() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetTeacherName() + " " + ArrayThursday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (2): {
                    CoupleTwoThursday = CoupleTwoThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetTypeWeek() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetTeacherName() + " " + ArrayThursday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (3): {
                    CoupleThreeThursday = CoupleThreeThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetTypeWeek() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetTeacherName() + " " + ArrayThursday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (4): {
                    CoupleFourThursday = CoupleFourThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetTypeWeek() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetTeacherName() + " " + ArrayThursday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (5): {
                    CoupleFiveThursday = CoupleFiveThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetTypeWeek() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetTeacherName() + " " + ArrayThursday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (6): {
                    CoupleSixThursday = CoupleSixThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetTypeWeek() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetTeacherName() + " " + ArrayThursday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (7): {
                    CoupleSevenThursday = CoupleSevenThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetTypeWeek() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetTeacherName() + " " + ArrayThursday.get(i).GetAud() + "\n" + "\n";
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
                    CoupleOneFriday = CoupleOneFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetTypeWeek() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetTeacherName() + " " + ArrayFriday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (2): {
                    CoupleTwoFriday = CoupleTwoFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetTypeWeek() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetTeacherName() + " " + ArrayFriday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (3): {
                    CoupleThreeFriday = CoupleThreeFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetTypeWeek() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetTeacherName() + " " + ArrayFriday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (4): {
                    CoupleFourFriday = CoupleFourFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetTypeWeek() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetTeacherName() + " " + ArrayFriday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (5): {
                    CoupleFiveFriday = CoupleFiveFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetTypeWeek() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetTeacherName() + " " + ArrayFriday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (6): {
                    CoupleSixFriday = CoupleSixFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetTypeWeek() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetTeacherName() + " " + ArrayFriday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (7): {
                    CoupleSevenFriday = CoupleSevenFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetTypeWeek() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetTeacherName() + " " + ArrayFriday.get(i).GetAud() + "\n" + "\n";
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
                    CoupleOneSaturday = CoupleOneSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetTypeWeek() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetTeacherName() + " " + ArraySaturday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (2): {
                    CoupleTwoSaturday = CoupleTwoSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetTypeWeek() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetTeacherName() + " " + ArraySaturday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (3): {
                    CoupleThreeSaturday = CoupleThreeSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetTypeWeek() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetTeacherName() + " " + ArraySaturday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (4): {
                    CoupleFourSaturday = CoupleFourSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetTypeWeek() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetTeacherName() + " " + ArraySaturday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (5): {
                    CoupleFiveSaturday = CoupleFiveSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetTypeWeek() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetTeacherName() + " " + ArraySaturday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (6): {
                    CoupleSixSaturday = CoupleSixSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetTypeWeek() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetTeacherName() + " " + ArraySaturday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (7): {
                    CoupleSevenSaturday = CoupleSevenSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetTypeWeek() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetTeacherName() + " " + ArraySaturday.get(i).GetAud() + "\n" + "\n";
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


        FileOutputStream fileOutputStream = new FileOutputStream("OneGroupExelDoc");

        workbookOneElementGroup.write(fileOutputStream);
        fileOutputStream.close();
    }

}
