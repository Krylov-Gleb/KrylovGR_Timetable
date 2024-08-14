package timetablekrylov.timetablekrylovgr;

import javafx.css.PseudoClass;
import javafx.scene.control.CheckBox;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class CreatorTableExelClassrooms {

    int HeightPoints = 250;
    int ColumnWidth = 10000;

    private Workbook workbookClassroom = new HSSFWorkbook();

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
                case ("410/2"):{
                    Array410.add(ArrayCouple.get(i));
                    break;
                }
                case("411/2"):{
                    Array411.add(ArrayCouple.get(i));
                    break;
                }
                case("413/2"):{
                    Array413.add(ArrayCouple.get(i));
                    break;
                }
                case ("416/2"):{
                    Array416.add(ArrayCouple.get(i));
                    break;
                }
                case("417/2"):{
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

                    ArrayList<CoupleGroup> ArrayMonday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayTuesday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayWednesday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayThursday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayFriday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArraySaturday = new ArrayList<>();

                    for (int index = 0; index < Array410.size(); index++) {

                        int IdDay = Array410.get(index).GetIDDay();

                        switch (IdDay) {
                            case (1): {
                                ArrayMonday.add(Array410.get(index));
                                break;
                            }
                            case (2): {
                                ArrayTuesday.add(Array410.get(index));
                                break;
                            }
                            case (3): {
                                ArrayWednesday.add(Array410.get(index));
                                break;
                            }
                            case (4): {
                                ArrayThursday.add(Array410.get(index));
                                break;
                            }
                            case (5): {
                                ArrayFriday.add(Array410.get(index));
                                break;
                            }
                            case (6): {
                                ArraySaturday.add(Array410.get(index));
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
                                CoupleOneMonday = CoupleOneMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoMonday = CoupleTwoMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeMonday = CoupleThreeMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourMonday = CoupleFourMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveMonday = CoupleFiveMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixMonday = CoupleSixMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenMonday = CoupleSevenMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleMonday = rowOne.createCell(ClassroomNumber);
                    cellOneCoupleMonday.setCellValue(CoupleOneMonday);
                    cellOneCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowOne.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleMonday = rowTwo.createCell(ClassroomNumber);
                    cellTwoCoupleMonday.setCellValue(CoupleTwoMonday);
                    cellTwoCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwo.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleMonday = rowThree.createCell(ClassroomNumber);
                    cellThreeCoupleMonday.setCellValue(CoupleThreeMonday);
                    cellThreeCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThree.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleMonday = rowFour.createCell(ClassroomNumber);
                    cellFourCoupleMonday.setCellValue(CoupleFourMonday);
                    cellFourCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFour.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleMonday = rowFive.createCell(ClassroomNumber);
                    cellFiveCoupleMonday.setCellValue(CoupleFiveMonday);
                    cellFiveCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFive.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleMonday = rowSix.createCell(ClassroomNumber);
                    cellSixCoupleMonday.setCellValue(CoupleSixMonday);
                    cellSixCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowSix.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleMonday = rowSeven.createCell(ClassroomNumber);
                    cellSevenCoupleMonday.setCellValue(CoupleSevenMonday);
                    cellSevenCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneTuesday = CoupleOneTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoTuesday = CoupleTwoTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeTuesday = CoupleThreeTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourTuesday = CoupleFourTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveTuesday = CoupleFiveTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixTuesday = CoupleSixTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenTuesday = CoupleSevenTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleTuesday = rowEight.createCell(ClassroomNumber);
                    cellOneCoupleTuesday.setCellValue(CoupleOneTuesday);
                    cellOneCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowEight.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleTuesday = rowNine.createCell(ClassroomNumber);
                    cellTwoCoupleTuesday.setCellValue(CoupleTwoTuesday);
                    cellTwoCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowNine.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleTuesday = rowTen.createCell(ClassroomNumber);
                    cellThreeCoupleTuesday.setCellValue(CoupleThreeTuesday);
                    cellThreeCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTen.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleTuesday = rowEleven.createCell(ClassroomNumber);
                    cellFourCoupleTuesday.setCellValue(CoupleFourTuesday);
                    cellFourCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowEleven.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleTuesday = rowTwelve.createCell(ClassroomNumber);
                    cellFiveCoupleTuesday.setCellValue(CoupleFiveTuesday);
                    cellFiveCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwelve.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleTuesday = rowThirteen.createCell(ClassroomNumber);
                    cellSixCoupleTuesday.setCellValue(CoupleSixTuesday);
                    cellSixCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirteen.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleTuesday = rowFourteen.createCell(ClassroomNumber);
                    cellSevenCoupleTuesday.setCellValue(CoupleSevenTuesday);
                    cellSevenCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneWednesday = CoupleOneWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoWednesday = CoupleTwoWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeWednesday = CoupleThreeWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourWednesday = CoupleFourWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveWednesday = CoupleFiveWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixWednesday = CoupleSixWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenWednesday = CoupleSevenWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                        Cell cellOneCoupleWednesday = rowfifteen.createCell(ClassroomNumber);
                    cellOneCoupleWednesday.setCellValue(CoupleOneWednesday);
                    cellOneCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowfifteen.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleWednesday = rowSixteen.createCell(ClassroomNumber);
                    cellTwoCoupleWednesday.setCellValue(CoupleTwoWednesday);
                    cellTwoCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowSixteen.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleWednesday = rowSeventeen.createCell(ClassroomNumber);
                    cellThreeCoupleWednesday.setCellValue(CoupleThreeWednesday);
                    cellThreeCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowSeventeen.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleWednesday = rowEighteen.createCell(ClassroomNumber);
                    cellFourCoupleWednesday.setCellValue(CoupleFourWednesday);
                    cellFourCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowEighteen.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleWednesday = rowNineteen.createCell(ClassroomNumber);
                    cellFiveCoupleWednesday.setCellValue(CoupleFiveWednesday);
                    cellFiveCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowNineteen.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleWednesday = rowTwenty.createCell(ClassroomNumber);
                    cellSixCoupleWednesday.setCellValue(CoupleSixWednesday);
                    cellSixCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwenty.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleWednesday = rowTwentyOne.createCell(ClassroomNumber);
                    cellSevenCoupleWednesday.setCellValue(CoupleSevenWednesday);
                    cellSevenCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyOne.setHeightInPoints(HeightPoints);

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
                                CoupleOneThursday = CoupleOneThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoThursday = CoupleTwoThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeThursday = CoupleThreeThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourThursday = CoupleFourThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveThursday = CoupleFiveThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixThursday = CoupleSixThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenThursday = CoupleSevenThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                        Cell cellOneCoupleThursday = rowTwentyTwo.createCell(ClassroomNumber);
                    cellOneCoupleThursday.setCellValue(CoupleOneThursday);
                    cellOneCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyTwo.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleThursday = rowTwentyThree.createCell(ClassroomNumber);
                    cellTwoCoupleThursday.setCellValue(CoupleTwoThursday);
                    cellTwoCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyThree.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleThursday = rowTwentyFour.createCell(ClassroomNumber);
                    cellThreeCoupleThursday.setCellValue(CoupleThreeThursday);
                    cellThreeCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyFour.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleThursday = rowTwentyFive.createCell(ClassroomNumber);
                    cellFourCoupleThursday.setCellValue(CoupleFourThursday);
                    cellFourCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyFive.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleThursday = rowTwentySix.createCell(ClassroomNumber);
                    cellFiveCoupleThursday.setCellValue(CoupleFiveThursday);
                    cellFiveCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentySix.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleThursday = rowTwentySeven.createCell(ClassroomNumber);
                    cellSixCoupleThursday.setCellValue(CoupleSixThursday);
                    cellSixCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentySeven.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleThursday = rowTwentyEight.createCell(ClassroomNumber);
                    cellSevenCoupleThursday.setCellValue(CoupleSevenThursday);
                    cellSevenCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneFriday = CoupleOneFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoFriday = CoupleTwoFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeFriday = CoupleThreeFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourFriday = CoupleFourFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveFriday = CoupleFiveFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }case (6): {
                                CoupleSixFriday = CoupleSixFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenFriday = CoupleSevenFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleFriday = rowTwentyNine.createCell(ClassroomNumber);
                    cellOneCoupleFriday.setCellValue(CoupleOneFriday);
                    cellOneCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyNine.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleFriday = rowThirty.createCell(ClassroomNumber);
                    cellTwoCoupleFriday.setCellValue(CoupleTwoFriday);
                    cellTwoCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirty.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleFriday = rowThirtyOne.createCell(ClassroomNumber);
                    cellThreeCoupleFriday.setCellValue(CoupleThreeFriday);
                    cellThreeCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyOne.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleFriday = rowThirtyTwo.createCell(ClassroomNumber);
                    cellFourCoupleFriday.setCellValue(CoupleFourFriday);
                    cellFourCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyTwo.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleFriday = rowThirtyThree.createCell(ClassroomNumber);
                    cellFiveCoupleFriday.setCellValue(CoupleFiveFriday);
                    cellFiveCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyThree.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleFriday = rowThirtyFour.createCell(ClassroomNumber);
                    cellSixCoupleFriday.setCellValue(CoupleSixFriday);
                    cellSixCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyFour.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleFriday = rowThirtyFive.createCell(ClassroomNumber);
                    cellSevenCoupleFriday.setCellValue(CoupleSevenFriday);
                    cellSevenCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneSaturday = CoupleOneSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoSaturday = CoupleTwoSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeSaturday = CoupleThreeSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourSaturday = CoupleFourSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveSaturday = CoupleFiveSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixSaturday = CoupleSixSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenSaturday = CoupleSevenSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleSaturday = rowThirtySix.createCell(ClassroomNumber);
                    cellOneCoupleSaturday.setCellValue(CoupleOneSaturday);
                    cellOneCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtySix.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleSaturday = rowThirtySeven.createCell(ClassroomNumber);
                    cellTwoCoupleSaturday.setCellValue(CoupleTwoSaturday);
                    cellTwoCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtySeven.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleSaturday = rowThirtyEight.createCell(ClassroomNumber);
                    cellThreeCoupleSaturday.setCellValue(CoupleThreeSaturday);
                    cellThreeCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyEight.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleSaturday = rowThirtyNine.createCell(ClassroomNumber);
                    cellFourCoupleSaturday.setCellValue(CoupleFourSaturday);
                    cellFourCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyNine.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleSaturday = rowForty.createCell(ClassroomNumber);
                    cellFiveCoupleSaturday.setCellValue(CoupleFiveSaturday);
                    cellFiveCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowForty.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleSaturday = rowFortyOne.createCell(ClassroomNumber);
                    cellSixCoupleSaturday.setCellValue(CoupleSixSaturday);
                    cellSixCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFortyOne.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleSaturday = rowFortyTwo.createCell(ClassroomNumber);
                    cellSevenCoupleSaturday.setCellValue(CoupleSevenSaturday);
                    cellSevenCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFortyTwo.setHeightInPoints(HeightPoints);

                    // -------------------------------------------------------------------------------------------------
                }

                if(ArrayClassroomCheckBox.get(i).getText().equals("411/2")) {

                    ArrayList<CoupleGroup> ArrayMonday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayTuesday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayWednesday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayThursday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayFriday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArraySaturday = new ArrayList<>();

                    for (int index = 0; index < Array411.size(); index++) {

                        int IdDay = Array411.get(index).GetIDDay();

                        switch (IdDay) {
                            case (1): {
                                ArrayMonday.add(Array411.get(index));
                                break;
                            }
                            case (2): {
                                ArrayTuesday.add(Array411.get(index));
                                break;
                            }
                            case (3): {
                                ArrayWednesday.add(Array411.get(index));
                                break;
                            }
                            case (4): {
                                ArrayThursday.add(Array411.get(index));
                                break;
                            }
                            case (5): {
                                ArrayFriday.add(Array411.get(index));
                                break;
                            }
                            case (6): {
                                ArraySaturday.add(Array411.get(index));
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
                                CoupleOneMonday = CoupleOneMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoMonday = CoupleTwoMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeMonday = CoupleThreeMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourMonday = CoupleFourMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveMonday = CoupleFiveMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixMonday = CoupleSixMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenMonday = CoupleSevenMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleMonday = rowOne.createCell(ClassroomNumber);
                    cellOneCoupleMonday.setCellValue(CoupleOneMonday);
                    cellOneCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowOne.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleMonday = rowTwo.createCell(ClassroomNumber);
                    cellTwoCoupleMonday.setCellValue(CoupleTwoMonday);
                    cellTwoCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwo.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleMonday = rowThree.createCell(ClassroomNumber);
                    cellThreeCoupleMonday.setCellValue(CoupleThreeMonday);
                    cellThreeCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThree.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleMonday = rowFour.createCell(ClassroomNumber);
                    cellFourCoupleMonday.setCellValue(CoupleFourMonday);
                    cellFourCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFour.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleMonday = rowFive.createCell(ClassroomNumber);
                    cellFiveCoupleMonday.setCellValue(CoupleFiveMonday);
                    cellFiveCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFive.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleMonday = rowSix.createCell(ClassroomNumber);
                    cellSixCoupleMonday.setCellValue(CoupleSixMonday);
                    cellSixCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowSix.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleMonday = rowSeven.createCell(ClassroomNumber);
                    cellSevenCoupleMonday.setCellValue(CoupleSevenMonday);
                    cellSevenCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneTuesday = CoupleOneTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoTuesday = CoupleTwoTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeTuesday = CoupleThreeTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourTuesday = CoupleFourTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveTuesday = CoupleFiveTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixTuesday = CoupleSixTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenTuesday = CoupleSevenTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleTuesday = rowEight.createCell(ClassroomNumber);
                    cellOneCoupleTuesday.setCellValue(CoupleOneTuesday);
                    cellOneCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowEight.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleTuesday = rowNine.createCell(ClassroomNumber);
                    cellTwoCoupleTuesday.setCellValue(CoupleTwoTuesday);
                    cellTwoCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowNine.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleTuesday = rowTen.createCell(ClassroomNumber);
                    cellThreeCoupleTuesday.setCellValue(CoupleThreeTuesday);
                    cellThreeCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTen.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleTuesday = rowEleven.createCell(ClassroomNumber);
                    cellFourCoupleTuesday.setCellValue(CoupleFourTuesday);
                    cellFourCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowEleven.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleTuesday = rowTwelve.createCell(ClassroomNumber);
                    cellFiveCoupleTuesday.setCellValue(CoupleFiveTuesday);
                    cellFiveCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwelve.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleTuesday = rowThirteen.createCell(ClassroomNumber);
                    cellSixCoupleTuesday.setCellValue(CoupleSixTuesday);
                    cellSixCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirteen.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleTuesday = rowFourteen.createCell(ClassroomNumber);
                    cellSevenCoupleTuesday.setCellValue(CoupleSevenTuesday);
                    cellSevenCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneWednesday = CoupleOneWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoWednesday = CoupleTwoWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeWednesday = CoupleThreeWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourWednesday = CoupleFourWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveWednesday = CoupleFiveWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixWednesday = CoupleSixWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenWednesday = CoupleSevenWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleWednesday = rowfifteen.createCell(ClassroomNumber);
                    cellOneCoupleWednesday.setCellValue(CoupleOneWednesday);
                    cellOneCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowfifteen.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleWednesday = rowSixteen.createCell(ClassroomNumber);
                    cellTwoCoupleWednesday.setCellValue(CoupleTwoWednesday);
                    cellTwoCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowSixteen.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleWednesday = rowSeventeen.createCell(ClassroomNumber);
                    cellThreeCoupleWednesday.setCellValue(CoupleThreeWednesday);
                    cellThreeCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowSeventeen.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleWednesday = rowEighteen.createCell(ClassroomNumber);
                    cellFourCoupleWednesday.setCellValue(CoupleFourWednesday);
                    cellFourCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowEighteen.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleWednesday = rowNineteen.createCell(ClassroomNumber);
                    cellFiveCoupleWednesday.setCellValue(CoupleFiveWednesday);
                    cellFiveCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowNineteen.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleWednesday = rowTwenty.createCell(ClassroomNumber);
                    cellSixCoupleWednesday.setCellValue(CoupleSixWednesday);
                    cellSixCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwenty.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleWednesday = rowTwentyOne.createCell(ClassroomNumber);
                    cellSevenCoupleWednesday.setCellValue(CoupleSevenWednesday);
                    cellSevenCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyOne.setHeightInPoints(HeightPoints);

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
                                CoupleOneThursday = CoupleOneThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoThursday = CoupleTwoThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeThursday = CoupleThreeThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourThursday = CoupleFourThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveThursday = CoupleFiveThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixThursday = CoupleSixThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenThursday = CoupleSevenThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleThursday = rowTwentyTwo.createCell(ClassroomNumber);
                    cellOneCoupleThursday.setCellValue(CoupleOneThursday);
                    cellOneCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyTwo.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleThursday = rowTwentyThree.createCell(ClassroomNumber);
                    cellTwoCoupleThursday.setCellValue(CoupleTwoThursday);
                    cellTwoCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyThree.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleThursday = rowTwentyFour.createCell(ClassroomNumber);
                    cellThreeCoupleThursday.setCellValue(CoupleThreeThursday);
                    cellThreeCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyFour.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleThursday = rowTwentyFive.createCell(ClassroomNumber);
                    cellFourCoupleThursday.setCellValue(CoupleFourThursday);
                    cellFourCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyFive.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleThursday = rowTwentySix.createCell(ClassroomNumber);
                    cellFiveCoupleThursday.setCellValue(CoupleFiveThursday);
                    cellFiveCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentySix.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleThursday = rowTwentySeven.createCell(ClassroomNumber);
                    cellSixCoupleThursday.setCellValue(CoupleSixThursday);
                    cellSixCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentySeven.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleThursday = rowTwentyEight.createCell(ClassroomNumber);
                    cellSevenCoupleThursday.setCellValue(CoupleSevenThursday);
                    cellSevenCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneFriday = CoupleOneFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoFriday = CoupleTwoFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeFriday = CoupleThreeFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourFriday = CoupleFourFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveFriday = CoupleFiveFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }case (6): {
                                CoupleSixFriday = CoupleSixFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenFriday = CoupleSevenFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleFriday = rowTwentyNine.createCell(ClassroomNumber);
                    cellOneCoupleFriday.setCellValue(CoupleOneFriday);
                    cellOneCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyNine.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleFriday = rowThirty.createCell(ClassroomNumber);
                    cellTwoCoupleFriday.setCellValue(CoupleTwoFriday);
                    cellTwoCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirty.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleFriday = rowThirtyOne.createCell(ClassroomNumber);
                    cellThreeCoupleFriday.setCellValue(CoupleThreeFriday);
                    cellThreeCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyOne.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleFriday = rowThirtyTwo.createCell(ClassroomNumber);
                    cellFourCoupleFriday.setCellValue(CoupleFourFriday);
                    cellFourCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyTwo.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleFriday = rowThirtyThree.createCell(ClassroomNumber);
                    cellFiveCoupleFriday.setCellValue(CoupleFiveFriday);
                    cellFiveCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyThree.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleFriday = rowThirtyFour.createCell(ClassroomNumber);
                    cellSixCoupleFriday.setCellValue(CoupleSixFriday);
                    cellSixCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyFour.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleFriday = rowThirtyFive.createCell(ClassroomNumber);
                    cellSevenCoupleFriday.setCellValue(CoupleSevenFriday);
                    cellSevenCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneSaturday = CoupleOneSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoSaturday = CoupleTwoSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeSaturday = CoupleThreeSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourSaturday = CoupleFourSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveSaturday = CoupleFiveSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixSaturday = CoupleSixSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenSaturday = CoupleSevenSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleSaturday = rowThirtySix.createCell(ClassroomNumber);
                    cellOneCoupleSaturday.setCellValue(CoupleOneSaturday);
                    cellOneCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtySix.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleSaturday = rowThirtySeven.createCell(ClassroomNumber);
                    cellTwoCoupleSaturday.setCellValue(CoupleTwoSaturday);
                    cellTwoCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtySeven.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleSaturday = rowThirtyEight.createCell(ClassroomNumber);
                    cellThreeCoupleSaturday.setCellValue(CoupleThreeSaturday);
                    cellThreeCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyEight.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleSaturday = rowThirtyNine.createCell(ClassroomNumber);
                    cellFourCoupleSaturday.setCellValue(CoupleFourSaturday);
                    cellFourCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyNine.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleSaturday = rowForty.createCell(ClassroomNumber);
                    cellFiveCoupleSaturday.setCellValue(CoupleFiveSaturday);
                    cellFiveCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowForty.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleSaturday = rowFortyOne.createCell(ClassroomNumber);
                    cellSixCoupleSaturday.setCellValue(CoupleSixSaturday);
                    cellSixCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFortyOne.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleSaturday = rowFortyTwo.createCell(ClassroomNumber);
                    cellSevenCoupleSaturday.setCellValue(CoupleSevenSaturday);
                    cellSevenCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFortyTwo.setHeightInPoints(HeightPoints);

                }

                if(ArrayClassroomCheckBox.get(i).getText().equals("413/2")) {

                    ArrayList<CoupleGroup> ArrayMonday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayTuesday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayWednesday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayThursday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayFriday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArraySaturday = new ArrayList<>();

                    for (int index = 0; index < Array413.size(); index++) {

                        int IdDay = Array413.get(index).GetIDDay();

                        switch (IdDay) {
                            case (1): {
                                ArrayMonday.add(Array413.get(index));
                                break;
                            }
                            case (2): {
                                ArrayTuesday.add(Array413.get(index));
                                break;
                            }
                            case (3): {
                                ArrayWednesday.add(Array413.get(index));
                                break;
                            }
                            case (4): {
                                ArrayThursday.add(Array413.get(index));
                                break;
                            }
                            case (5): {
                                ArrayFriday.add(Array413.get(index));
                                break;
                            }
                            case (6): {
                                ArraySaturday.add(Array413.get(index));
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
                                CoupleOneMonday = CoupleOneMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoMonday = CoupleTwoMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeMonday = CoupleThreeMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourMonday = CoupleFourMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveMonday = CoupleFiveMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixMonday = CoupleSixMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenMonday = CoupleSevenMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleMonday = rowOne.createCell(ClassroomNumber);
                    cellOneCoupleMonday.setCellValue(CoupleOneMonday);
                    cellOneCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowOne.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleMonday = rowTwo.createCell(ClassroomNumber);
                    cellTwoCoupleMonday.setCellValue(CoupleTwoMonday);
                    cellTwoCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwo.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleMonday = rowThree.createCell(ClassroomNumber);
                    cellThreeCoupleMonday.setCellValue(CoupleThreeMonday);
                    cellThreeCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThree.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleMonday = rowFour.createCell(ClassroomNumber);
                    cellFourCoupleMonday.setCellValue(CoupleFourMonday);
                    cellFourCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFour.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleMonday = rowFive.createCell(ClassroomNumber);
                    cellFiveCoupleMonday.setCellValue(CoupleFiveMonday);
                    cellFiveCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFive.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleMonday = rowSix.createCell(ClassroomNumber);
                    cellSixCoupleMonday.setCellValue(CoupleSixMonday);
                    cellSixCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowSix.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleMonday = rowSeven.createCell(ClassroomNumber);
                    cellSevenCoupleMonday.setCellValue(CoupleSevenMonday);
                    cellSevenCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneTuesday = CoupleOneTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoTuesday = CoupleTwoTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeTuesday = CoupleThreeTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourTuesday = CoupleFourTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveTuesday = CoupleFiveTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixTuesday = CoupleSixTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenTuesday = CoupleSevenTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleTuesday = rowEight.createCell(ClassroomNumber);
                    cellOneCoupleTuesday.setCellValue(CoupleOneTuesday);
                    cellOneCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowEight.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleTuesday = rowNine.createCell(ClassroomNumber);
                    cellTwoCoupleTuesday.setCellValue(CoupleTwoTuesday);
                    cellTwoCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowNine.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleTuesday = rowTen.createCell(ClassroomNumber);
                    cellThreeCoupleTuesday.setCellValue(CoupleThreeTuesday);
                    cellThreeCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTen.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleTuesday = rowEleven.createCell(ClassroomNumber);
                    cellFourCoupleTuesday.setCellValue(CoupleFourTuesday);
                    cellFourCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowEleven.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleTuesday = rowTwelve.createCell(ClassroomNumber);
                    cellFiveCoupleTuesday.setCellValue(CoupleFiveTuesday);
                    cellFiveCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwelve.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleTuesday = rowThirteen.createCell(ClassroomNumber);
                    cellSixCoupleTuesday.setCellValue(CoupleSixTuesday);
                    cellSixCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirteen.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleTuesday = rowFourteen.createCell(ClassroomNumber);
                    cellSevenCoupleTuesday.setCellValue(CoupleSevenTuesday);
                    cellSevenCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneWednesday = CoupleOneWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoWednesday = CoupleTwoWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeWednesday = CoupleThreeWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourWednesday = CoupleFourWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveWednesday = CoupleFiveWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixWednesday = CoupleSixWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenWednesday = CoupleSevenWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleWednesday = rowfifteen.createCell(ClassroomNumber);
                    cellOneCoupleWednesday.setCellValue(CoupleOneWednesday);
                    cellOneCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowfifteen.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleWednesday = rowSixteen.createCell(ClassroomNumber);
                    cellTwoCoupleWednesday.setCellValue(CoupleTwoWednesday);
                    cellTwoCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowSixteen.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleWednesday = rowSeventeen.createCell(ClassroomNumber);
                    cellThreeCoupleWednesday.setCellValue(CoupleThreeWednesday);
                    cellThreeCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowSeventeen.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleWednesday = rowEighteen.createCell(ClassroomNumber);
                    cellFourCoupleWednesday.setCellValue(CoupleFourWednesday);
                    cellFourCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowEighteen.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleWednesday = rowNineteen.createCell(ClassroomNumber);
                    cellFiveCoupleWednesday.setCellValue(CoupleFiveWednesday);
                    cellFiveCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowNineteen.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleWednesday = rowTwenty.createCell(ClassroomNumber);
                    cellSixCoupleWednesday.setCellValue(CoupleSixWednesday);
                    cellSixCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwenty.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleWednesday = rowTwentyOne.createCell(ClassroomNumber);
                    cellSevenCoupleWednesday.setCellValue(CoupleSevenWednesday);
                    cellSevenCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyOne.setHeightInPoints(HeightPoints);

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
                                CoupleOneThursday = CoupleOneThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoThursday = CoupleTwoThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeThursday = CoupleThreeThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourThursday = CoupleFourThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveThursday = CoupleFiveThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixThursday = CoupleSixThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenThursday = CoupleSevenThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleThursday = rowTwentyTwo.createCell(ClassroomNumber);
                    cellOneCoupleThursday.setCellValue(CoupleOneThursday);
                    cellOneCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyTwo.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleThursday = rowTwentyThree.createCell(ClassroomNumber);
                    cellTwoCoupleThursday.setCellValue(CoupleTwoThursday);
                    cellTwoCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyThree.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleThursday = rowTwentyFour.createCell(ClassroomNumber);
                    cellThreeCoupleThursday.setCellValue(CoupleThreeThursday);
                    cellThreeCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyFour.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleThursday = rowTwentyFive.createCell(ClassroomNumber);
                    cellFourCoupleThursday.setCellValue(CoupleFourThursday);
                    cellFourCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyFive.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleThursday = rowTwentySix.createCell(ClassroomNumber);
                    cellFiveCoupleThursday.setCellValue(CoupleFiveThursday);
                    cellFiveCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentySix.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleThursday = rowTwentySeven.createCell(ClassroomNumber);
                    cellSixCoupleThursday.setCellValue(CoupleSixThursday);
                    cellSixCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentySeven.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleThursday = rowTwentyEight.createCell(ClassroomNumber);
                    cellSevenCoupleThursday.setCellValue(CoupleSevenThursday);
                    cellSevenCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneFriday = CoupleOneFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoFriday = CoupleTwoFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeFriday = CoupleThreeFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourFriday = CoupleFourFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveFriday = CoupleFiveFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }case (6): {
                                CoupleSixFriday = CoupleSixFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenFriday = CoupleSevenFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleFriday = rowTwentyNine.createCell(ClassroomNumber);
                    cellOneCoupleFriday.setCellValue(CoupleOneFriday);
                    cellOneCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyNine.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleFriday = rowThirty.createCell(ClassroomNumber);
                    cellTwoCoupleFriday.setCellValue(CoupleTwoFriday);
                    cellTwoCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirty.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleFriday = rowThirtyOne.createCell(ClassroomNumber);
                    cellThreeCoupleFriday.setCellValue(CoupleThreeFriday);
                    cellThreeCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyOne.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleFriday = rowThirtyTwo.createCell(ClassroomNumber);
                    cellFourCoupleFriday.setCellValue(CoupleFourFriday);
                    cellFourCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyTwo.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleFriday = rowThirtyThree.createCell(ClassroomNumber);
                    cellFiveCoupleFriday.setCellValue(CoupleFiveFriday);
                    cellFiveCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyThree.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleFriday = rowThirtyFour.createCell(ClassroomNumber);
                    cellSixCoupleFriday.setCellValue(CoupleSixFriday);
                    cellSixCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyFour.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleFriday = rowThirtyFive.createCell(ClassroomNumber);
                    cellSevenCoupleFriday.setCellValue(CoupleSevenFriday);
                    cellSevenCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneSaturday = CoupleOneSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoSaturday = CoupleTwoSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeSaturday = CoupleThreeSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourSaturday = CoupleFourSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveSaturday = CoupleFiveSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixSaturday = CoupleSixSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenSaturday = CoupleSevenSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleSaturday = rowThirtySix.createCell(ClassroomNumber);
                    cellOneCoupleSaturday.setCellValue(CoupleOneSaturday);
                    cellOneCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtySix.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleSaturday = rowThirtySeven.createCell(ClassroomNumber);
                    cellTwoCoupleSaturday.setCellValue(CoupleTwoSaturday);
                    cellTwoCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtySeven.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleSaturday = rowThirtyEight.createCell(ClassroomNumber);
                    cellThreeCoupleSaturday.setCellValue(CoupleThreeSaturday);
                    cellThreeCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyEight.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleSaturday = rowThirtyNine.createCell(ClassroomNumber);
                    cellFourCoupleSaturday.setCellValue(CoupleFourSaturday);
                    cellFourCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyNine.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleSaturday = rowForty.createCell(ClassroomNumber);
                    cellFiveCoupleSaturday.setCellValue(CoupleFiveSaturday);
                    cellFiveCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowForty.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleSaturday = rowFortyOne.createCell(ClassroomNumber);
                    cellSixCoupleSaturday.setCellValue(CoupleSixSaturday);
                    cellSixCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFortyOne.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleSaturday = rowFortyTwo.createCell(ClassroomNumber);
                    cellSevenCoupleSaturday.setCellValue(CoupleSevenSaturday);
                    cellSevenCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFortyTwo.setHeightInPoints(HeightPoints);

                }

                if(ArrayClassroomCheckBox.get(i).getText().equals("416/2")) {

                    ArrayList<CoupleGroup> ArrayMonday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayTuesday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayWednesday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayThursday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayFriday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArraySaturday = new ArrayList<>();

                    for (int index = 0; index < Array416.size(); index++) {

                        int IdDay = Array416.get(index).GetIDDay();

                        switch (IdDay) {
                            case (1): {
                                ArrayMonday.add(Array416.get(index));
                                break;
                            }
                            case (2): {
                                ArrayTuesday.add(Array416.get(index));
                                break;
                            }
                            case (3): {
                                ArrayWednesday.add(Array416.get(index));
                                break;
                            }
                            case (4): {
                                ArrayThursday.add(Array416.get(index));
                                break;
                            }
                            case (5): {
                                ArrayFriday.add(Array416.get(index));
                                break;
                            }
                            case (6): {
                                ArraySaturday.add(Array416.get(index));
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
                                CoupleOneMonday = CoupleOneMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoMonday = CoupleTwoMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeMonday = CoupleThreeMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourMonday = CoupleFourMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveMonday = CoupleFiveMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixMonday = CoupleSixMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenMonday = CoupleSevenMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleMonday = rowOne.createCell(ClassroomNumber);
                    cellOneCoupleMonday.setCellValue(CoupleOneMonday);
                    cellOneCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowOne.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleMonday = rowTwo.createCell(ClassroomNumber);
                    cellTwoCoupleMonday.setCellValue(CoupleTwoMonday);
                    cellTwoCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwo.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleMonday = rowThree.createCell(ClassroomNumber);
                    cellThreeCoupleMonday.setCellValue(CoupleThreeMonday);
                    cellThreeCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThree.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleMonday = rowFour.createCell(ClassroomNumber);
                    cellFourCoupleMonday.setCellValue(CoupleFourMonday);
                    cellFourCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFour.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleMonday = rowFive.createCell(ClassroomNumber);
                    cellFiveCoupleMonday.setCellValue(CoupleFiveMonday);
                    cellFiveCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFive.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleMonday = rowSix.createCell(ClassroomNumber);
                    cellSixCoupleMonday.setCellValue(CoupleSixMonday);
                    cellSixCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowSix.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleMonday = rowSeven.createCell(ClassroomNumber);
                    cellSevenCoupleMonday.setCellValue(CoupleSevenMonday);
                    cellSevenCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneTuesday = CoupleOneTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoTuesday = CoupleTwoTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeTuesday = CoupleThreeTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourTuesday = CoupleFourTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveTuesday = CoupleFiveTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixTuesday = CoupleSixTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenTuesday = CoupleSevenTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleTuesday = rowEight.createCell(ClassroomNumber);
                    cellOneCoupleTuesday.setCellValue(CoupleOneTuesday);
                    cellOneCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowEight.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleTuesday = rowNine.createCell(ClassroomNumber);
                    cellTwoCoupleTuesday.setCellValue(CoupleTwoTuesday);
                    cellTwoCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowNine.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleTuesday = rowTen.createCell(ClassroomNumber);
                    cellThreeCoupleTuesday.setCellValue(CoupleThreeTuesday);
                    cellThreeCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTen.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleTuesday = rowEleven.createCell(ClassroomNumber);
                    cellFourCoupleTuesday.setCellValue(CoupleFourTuesday);
                    cellFourCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowEleven.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleTuesday = rowTwelve.createCell(ClassroomNumber);
                    cellFiveCoupleTuesday.setCellValue(CoupleFiveTuesday);
                    cellFiveCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwelve.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleTuesday = rowThirteen.createCell(ClassroomNumber);
                    cellSixCoupleTuesday.setCellValue(CoupleSixTuesday);
                    cellSixCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirteen.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleTuesday = rowFourteen.createCell(ClassroomNumber);
                    cellSevenCoupleTuesday.setCellValue(CoupleSevenTuesday);
                    cellSevenCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneWednesday = CoupleOneWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoWednesday = CoupleTwoWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeWednesday = CoupleThreeWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourWednesday = CoupleFourWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveWednesday = CoupleFiveWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixWednesday = CoupleSixWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenWednesday = CoupleSevenWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleWednesday = rowfifteen.createCell(ClassroomNumber);
                    cellOneCoupleWednesday.setCellValue(CoupleOneWednesday);
                    cellOneCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowfifteen.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleWednesday = rowSixteen.createCell(ClassroomNumber);
                    cellTwoCoupleWednesday.setCellValue(CoupleTwoWednesday);
                    cellTwoCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowSixteen.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleWednesday = rowSeventeen.createCell(ClassroomNumber);
                    cellThreeCoupleWednesday.setCellValue(CoupleThreeWednesday);
                    cellThreeCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowSeventeen.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleWednesday = rowEighteen.createCell(ClassroomNumber);
                    cellFourCoupleWednesday.setCellValue(CoupleFourWednesday);
                    cellFourCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowEighteen.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleWednesday = rowNineteen.createCell(ClassroomNumber);
                    cellFiveCoupleWednesday.setCellValue(CoupleFiveWednesday);
                    cellFiveCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowNineteen.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleWednesday = rowTwenty.createCell(ClassroomNumber);
                    cellSixCoupleWednesday.setCellValue(CoupleSixWednesday);
                    cellSixCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwenty.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleWednesday = rowTwentyOne.createCell(ClassroomNumber);
                    cellSevenCoupleWednesday.setCellValue(CoupleSevenWednesday);
                    cellSevenCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyOne.setHeightInPoints(HeightPoints);

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
                                CoupleOneThursday = CoupleOneThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoThursday = CoupleTwoThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeThursday = CoupleThreeThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourThursday = CoupleFourThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveThursday = CoupleFiveThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixThursday = CoupleSixThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenThursday = CoupleSevenThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleThursday = rowTwentyTwo.createCell(ClassroomNumber);
                    cellOneCoupleThursday.setCellValue(CoupleOneThursday);
                    cellOneCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyTwo.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleThursday = rowTwentyThree.createCell(ClassroomNumber);
                    cellTwoCoupleThursday.setCellValue(CoupleTwoThursday);
                    cellTwoCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyThree.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleThursday = rowTwentyFour.createCell(ClassroomNumber);
                    cellThreeCoupleThursday.setCellValue(CoupleThreeThursday);
                    cellThreeCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyFour.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleThursday = rowTwentyFive.createCell(ClassroomNumber);
                    cellFourCoupleThursday.setCellValue(CoupleFourThursday);
                    cellFourCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyFive.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleThursday = rowTwentySix.createCell(ClassroomNumber);
                    cellFiveCoupleThursday.setCellValue(CoupleFiveThursday);
                    cellFiveCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentySix.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleThursday = rowTwentySeven.createCell(ClassroomNumber);
                    cellSixCoupleThursday.setCellValue(CoupleSixThursday);
                    cellSixCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentySeven.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleThursday = rowTwentyEight.createCell(ClassroomNumber);
                    cellSevenCoupleThursday.setCellValue(CoupleSevenThursday);
                    cellSevenCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneFriday = CoupleOneFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoFriday = CoupleTwoFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeFriday = CoupleThreeFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourFriday = CoupleFourFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveFriday = CoupleFiveFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }case (6): {
                                CoupleSixFriday = CoupleSixFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenFriday = CoupleSevenFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleFriday = rowTwentyNine.createCell(ClassroomNumber);
                    cellOneCoupleFriday.setCellValue(CoupleOneFriday);
                    cellOneCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyNine.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleFriday = rowThirty.createCell(ClassroomNumber);
                    cellTwoCoupleFriday.setCellValue(CoupleTwoFriday);
                    cellTwoCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirty.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleFriday = rowThirtyOne.createCell(ClassroomNumber);
                    cellThreeCoupleFriday.setCellValue(CoupleThreeFriday);
                    cellThreeCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyOne.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleFriday = rowThirtyTwo.createCell(ClassroomNumber);
                    cellFourCoupleFriday.setCellValue(CoupleFourFriday);
                    cellFourCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyTwo.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleFriday = rowThirtyThree.createCell(ClassroomNumber);
                    cellFiveCoupleFriday.setCellValue(CoupleFiveFriday);
                    cellFiveCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyThree.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleFriday = rowThirtyFour.createCell(ClassroomNumber);
                    cellSixCoupleFriday.setCellValue(CoupleSixFriday);
                    cellSixCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyFour.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleFriday = rowThirtyFive.createCell(ClassroomNumber);
                    cellSevenCoupleFriday.setCellValue(CoupleSevenFriday);
                    cellSevenCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneSaturday = CoupleOneSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoSaturday = CoupleTwoSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeSaturday = CoupleThreeSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourSaturday = CoupleFourSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveSaturday = CoupleFiveSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixSaturday = CoupleSixSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenSaturday = CoupleSevenSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleSaturday = rowThirtySix.createCell(ClassroomNumber);
                    cellOneCoupleSaturday.setCellValue(CoupleOneSaturday);
                    cellOneCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtySix.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleSaturday = rowThirtySeven.createCell(ClassroomNumber);
                    cellTwoCoupleSaturday.setCellValue(CoupleTwoSaturday);
                    cellTwoCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtySeven.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleSaturday = rowThirtyEight.createCell(ClassroomNumber);
                    cellThreeCoupleSaturday.setCellValue(CoupleThreeSaturday);
                    cellThreeCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyEight.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleSaturday = rowThirtyNine.createCell(ClassroomNumber);
                    cellFourCoupleSaturday.setCellValue(CoupleFourSaturday);
                    cellFourCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyNine.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleSaturday = rowForty.createCell(ClassroomNumber);
                    cellFiveCoupleSaturday.setCellValue(CoupleFiveSaturday);
                    cellFiveCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowForty.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleSaturday = rowFortyOne.createCell(ClassroomNumber);
                    cellSixCoupleSaturday.setCellValue(CoupleSixSaturday);
                    cellSixCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFortyOne.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleSaturday = rowFortyTwo.createCell(ClassroomNumber);
                    cellSevenCoupleSaturday.setCellValue(CoupleSevenSaturday);
                    cellSevenCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFortyTwo.setHeightInPoints(HeightPoints);

                }

                if(ArrayClassroomCheckBox.get(i).getText().equals("417/2")) {

                    ArrayList<CoupleGroup> ArrayMonday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayTuesday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayWednesday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayThursday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArrayFriday = new ArrayList<>();
                    ArrayList<CoupleGroup> ArraySaturday = new ArrayList<>();

                    for (int index = 0; index < Array417.size(); index++) {

                        int IdDay = Array417.get(index).GetIDDay();

                        switch (IdDay) {
                            case (1): {
                                ArrayMonday.add(Array417.get(index));
                                break;
                            }
                            case (2): {
                                ArrayTuesday.add(Array417.get(index));
                                break;
                            }
                            case (3): {
                                ArrayWednesday.add(Array417.get(index));
                                break;
                            }
                            case (4): {
                                ArrayThursday.add(Array417.get(index));
                                break;
                            }
                            case (5): {
                                ArrayFriday.add(Array417.get(index));
                                break;
                            }
                            case (6): {
                                ArraySaturday.add(Array417.get(index));
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
                                CoupleOneMonday = CoupleOneMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoMonday = CoupleTwoMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeMonday = CoupleThreeMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourMonday = CoupleFourMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveMonday = CoupleFiveMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixMonday = CoupleSixMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenMonday = CoupleSevenMonday + ArrayMonday.get(index2).GetDiscipline() + " (" + ArrayMonday.get(index2).GetCoupleType() + ")\n" + ArrayMonday.get(index2).GetNumberWeek() + " " + ArrayMonday.get(index2).GetTeacherName() + " " + ArrayMonday.get(index2).GetAud() + " (" + ArrayMonday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleMonday = rowOne.createCell(ClassroomNumber);
                    cellOneCoupleMonday.setCellValue(CoupleOneMonday);
                    cellOneCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowOne.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleMonday = rowTwo.createCell(ClassroomNumber);
                    cellTwoCoupleMonday.setCellValue(CoupleTwoMonday);
                    cellTwoCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwo.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleMonday = rowThree.createCell(ClassroomNumber);
                    cellThreeCoupleMonday.setCellValue(CoupleThreeMonday);
                    cellThreeCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThree.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleMonday = rowFour.createCell(ClassroomNumber);
                    cellFourCoupleMonday.setCellValue(CoupleFourMonday);
                    cellFourCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFour.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleMonday = rowFive.createCell(ClassroomNumber);
                    cellFiveCoupleMonday.setCellValue(CoupleFiveMonday);
                    cellFiveCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFive.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleMonday = rowSix.createCell(ClassroomNumber);
                    cellSixCoupleMonday.setCellValue(CoupleSixMonday);
                    cellSixCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowSix.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleMonday = rowSeven.createCell(ClassroomNumber);
                    cellSevenCoupleMonday.setCellValue(CoupleSevenMonday);
                    cellSevenCoupleMonday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneTuesday = CoupleOneTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoTuesday = CoupleTwoTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeTuesday = CoupleThreeTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourTuesday = CoupleFourTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveTuesday = CoupleFiveTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixTuesday = CoupleSixTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenTuesday = CoupleSevenTuesday + ArrayTuesday.get(index2).GetDiscipline() + " (" + ArrayTuesday.get(index2).GetCoupleType() + ")\n" + ArrayTuesday.get(index2).GetNumberWeek() + " " + ArrayTuesday.get(index2).GetTeacherName() + " " + ArrayTuesday.get(index2).GetAud() + " (" +  ArrayTuesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleTuesday = rowEight.createCell(ClassroomNumber);
                    cellOneCoupleTuesday.setCellValue(CoupleOneTuesday);
                    cellOneCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowEight.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleTuesday = rowNine.createCell(ClassroomNumber);
                    cellTwoCoupleTuesday.setCellValue(CoupleTwoTuesday);
                    cellTwoCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowNine.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleTuesday = rowTen.createCell(ClassroomNumber);
                    cellThreeCoupleTuesday.setCellValue(CoupleThreeTuesday);
                    cellThreeCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTen.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleTuesday = rowEleven.createCell(ClassroomNumber);
                    cellFourCoupleTuesday.setCellValue(CoupleFourTuesday);
                    cellFourCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowEleven.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleTuesday = rowTwelve.createCell(ClassroomNumber);
                    cellFiveCoupleTuesday.setCellValue(CoupleFiveTuesday);
                    cellFiveCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwelve.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleTuesday = rowThirteen.createCell(ClassroomNumber);
                    cellSixCoupleTuesday.setCellValue(CoupleSixTuesday);
                    cellSixCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirteen.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleTuesday = rowFourteen.createCell(ClassroomNumber);
                    cellSevenCoupleTuesday.setCellValue(CoupleSevenTuesday);
                    cellSevenCoupleTuesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneWednesday = CoupleOneWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoWednesday = CoupleTwoWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeWednesday = CoupleThreeWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourWednesday = CoupleFourWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveWednesday = CoupleFiveWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixWednesday = CoupleSixWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenWednesday = CoupleSevenWednesday + ArrayWednesday.get(index2).GetDiscipline() + " (" + ArrayWednesday.get(index2).GetCoupleType() + ")\n" + ArrayWednesday.get(index2).GetNumberWeek() + " " + ArrayWednesday.get(index2).GetTeacherName() + " " + ArrayWednesday.get(index2).GetAud() + " (" +  ArrayWednesday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleWednesday = rowfifteen.createCell(ClassroomNumber);
                    cellOneCoupleWednesday.setCellValue(CoupleOneWednesday);
                    cellOneCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowfifteen.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleWednesday = rowSixteen.createCell(ClassroomNumber);
                    cellTwoCoupleWednesday.setCellValue(CoupleTwoWednesday);
                    cellTwoCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowSixteen.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleWednesday = rowSeventeen.createCell(ClassroomNumber);
                    cellThreeCoupleWednesday.setCellValue(CoupleThreeWednesday);
                    cellThreeCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowSeventeen.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleWednesday = rowEighteen.createCell(ClassroomNumber);
                    cellFourCoupleWednesday.setCellValue(CoupleFourWednesday);
                    cellFourCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowEighteen.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleWednesday = rowNineteen.createCell(ClassroomNumber);
                    cellFiveCoupleWednesday.setCellValue(CoupleFiveWednesday);
                    cellFiveCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowNineteen.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleWednesday = rowTwenty.createCell(ClassroomNumber);
                    cellSixCoupleWednesday.setCellValue(CoupleSixWednesday);
                    cellSixCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwenty.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleWednesday = rowTwentyOne.createCell(ClassroomNumber);
                    cellSevenCoupleWednesday.setCellValue(CoupleSevenWednesday);
                    cellSevenCoupleWednesday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyOne.setHeightInPoints(HeightPoints);

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
                                CoupleOneThursday = CoupleOneThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoThursday = CoupleTwoThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeThursday = CoupleThreeThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourThursday = CoupleFourThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveThursday = CoupleFiveThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixThursday = CoupleSixThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenThursday = CoupleSevenThursday + ArrayThursday.get(index2).GetDiscipline() + " (" + ArrayThursday.get(index2).GetCoupleType() + ")\n" + ArrayThursday.get(index2).GetNumberWeek() + " " + ArrayThursday.get(index2).GetTeacherName() + " " + ArrayThursday.get(index2).GetAud() + " (" +  ArrayThursday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleThursday = rowTwentyTwo.createCell(ClassroomNumber);
                    cellOneCoupleThursday.setCellValue(CoupleOneThursday);
                    cellOneCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyTwo.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleThursday = rowTwentyThree.createCell(ClassroomNumber);
                    cellTwoCoupleThursday.setCellValue(CoupleTwoThursday);
                    cellTwoCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyThree.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleThursday = rowTwentyFour.createCell(ClassroomNumber);
                    cellThreeCoupleThursday.setCellValue(CoupleThreeThursday);
                    cellThreeCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyFour.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleThursday = rowTwentyFive.createCell(ClassroomNumber);
                    cellFourCoupleThursday.setCellValue(CoupleFourThursday);
                    cellFourCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyFive.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleThursday = rowTwentySix.createCell(ClassroomNumber);
                    cellFiveCoupleThursday.setCellValue(CoupleFiveThursday);
                    cellFiveCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentySix.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleThursday = rowTwentySeven.createCell(ClassroomNumber);
                    cellSixCoupleThursday.setCellValue(CoupleSixThursday);
                    cellSixCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentySeven.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleThursday = rowTwentyEight.createCell(ClassroomNumber);
                    cellSevenCoupleThursday.setCellValue(CoupleSevenThursday);
                    cellSevenCoupleThursday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneFriday = CoupleOneFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoFriday = CoupleTwoFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeFriday = CoupleThreeFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourFriday = CoupleFourFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveFriday = CoupleFiveFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }case (6): {
                                CoupleSixFriday = CoupleSixFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenFriday = CoupleSevenFriday + ArrayFriday.get(index2).GetDiscipline() + " (" + ArrayFriday.get(index2).GetCoupleType() + ")\n" + ArrayFriday.get(index2).GetNumberWeek() + " " + ArrayFriday.get(index2).GetTeacherName() + " " + ArrayFriday.get(index2).GetAud() + " (" +  ArrayFriday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleFriday = rowTwentyNine.createCell(ClassroomNumber);
                    cellOneCoupleFriday.setCellValue(CoupleOneFriday);
                    cellOneCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowTwentyNine.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleFriday = rowThirty.createCell(ClassroomNumber);
                    cellTwoCoupleFriday.setCellValue(CoupleTwoFriday);
                    cellTwoCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirty.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleFriday = rowThirtyOne.createCell(ClassroomNumber);
                    cellThreeCoupleFriday.setCellValue(CoupleThreeFriday);
                    cellThreeCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyOne.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleFriday = rowThirtyTwo.createCell(ClassroomNumber);
                    cellFourCoupleFriday.setCellValue(CoupleFourFriday);
                    cellFourCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyTwo.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleFriday = rowThirtyThree.createCell(ClassroomNumber);
                    cellFiveCoupleFriday.setCellValue(CoupleFiveFriday);
                    cellFiveCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyThree.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleFriday = rowThirtyFour.createCell(ClassroomNumber);
                    cellSixCoupleFriday.setCellValue(CoupleSixFriday);
                    cellSixCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyFour.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleFriday = rowThirtyFive.createCell(ClassroomNumber);
                    cellSevenCoupleFriday.setCellValue(CoupleSevenFriday);
                    cellSevenCoupleFriday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
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
                                CoupleOneSaturday = CoupleOneSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (2): {
                                CoupleTwoSaturday = CoupleTwoSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (3): {
                                CoupleThreeSaturday = CoupleThreeSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (4): {
                                CoupleFourSaturday = CoupleFourSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (5): {
                                CoupleFiveSaturday = CoupleFiveSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (6): {
                                CoupleSixSaturday = CoupleSixSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                            case (7): {
                                CoupleSevenSaturday = CoupleSevenSaturday + ArraySaturday.get(index2).GetDiscipline() + " (" + ArraySaturday.get(index2).GetCoupleType() + ")\n" + ArraySaturday.get(index2).GetNumberWeek() + " " + ArraySaturday.get(index2).GetTeacherName() + " " + ArraySaturday.get(index2).GetAud() + " (" +  ArraySaturday.get(index2).GetTypeWeek() + ")" + "\n" + "\n";
                                break;
                            }
                        }
                    }

                    Cell cellOneCoupleSaturday = rowThirtySix.createCell(ClassroomNumber);
                    cellOneCoupleSaturday.setCellValue(CoupleOneSaturday);
                    cellOneCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtySix.setHeightInPoints(HeightPoints);

                    Cell cellTwoCoupleSaturday = rowThirtySeven.createCell(ClassroomNumber);
                    cellTwoCoupleSaturday.setCellValue(CoupleTwoSaturday);
                    cellTwoCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtySeven.setHeightInPoints(HeightPoints);

                    Cell cellThreeCoupleSaturday = rowThirtyEight.createCell(ClassroomNumber);
                    cellThreeCoupleSaturday.setCellValue(CoupleThreeSaturday);
                    cellThreeCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyEight.setHeightInPoints(HeightPoints);

                    Cell cellFourCoupleSaturday = rowThirtyNine.createCell(ClassroomNumber);
                    cellFourCoupleSaturday.setCellValue(CoupleFourSaturday);
                    cellFourCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowThirtyNine.setHeightInPoints(HeightPoints);

                    Cell cellFiveCoupleSaturday = rowForty.createCell(ClassroomNumber);
                    cellFiveCoupleSaturday.setCellValue(CoupleFiveSaturday);
                    cellFiveCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowForty.setHeightInPoints(HeightPoints);

                    Cell cellSixCoupleSaturday = rowFortyOne.createCell(ClassroomNumber);
                    cellSixCoupleSaturday.setCellValue(CoupleSixSaturday);
                    cellSixCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFortyOne.setHeightInPoints(HeightPoints);

                    Cell cellSevenCoupleSaturday = rowFortyTwo.createCell(ClassroomNumber);
                    cellSevenCoupleSaturday.setCellValue(CoupleSevenSaturday);
                    cellSevenCoupleSaturday.setCellStyle(cellStyle);
                    Classroom.setColumnWidth(ClassroomNumber, ColumnWidth);
                    rowFortyTwo.setHeightInPoints(HeightPoints);

                }
            }
            ClassroomNumber++;
        }

        FileOutputStream fileOutputStream = new FileOutputStream("AllClassroomExelDoc");

        workbookClassroom.write(fileOutputStream);
        fileOutputStream.close();

    }

}
