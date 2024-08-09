package timetablekrylov.timetablekrylovgr;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class CreatorTableExelTeacher {

    private Workbook workbookTeacher = new HSSFWorkbook();

    private Sheet Teacher = workbookTeacher.createSheet("Преподаватель");

    private Row rowTeacherZero = Teacher.createRow(0);
    private Cell cellTeacherDayWeekAndCoupleNumber = rowTeacherZero.createCell(0);

    private Cell Monday = rowTeacherZero.createCell(1);
    private Cell Tuesday = rowTeacherZero.createCell(2);
    private Cell Wednesday = rowTeacherZero.createCell(3);
    private Cell Thursday = rowTeacherZero.createCell(4);
    private Cell Friday = rowTeacherZero.createCell(5);
    private Cell Saturday = rowTeacherZero.createCell(6);

    private Row StrOneCouple = Teacher.createRow(1);
    private Cell OneCouple = StrOneCouple.createCell(0);

    private Row StrTwoCouple = Teacher.createRow(2);
    private Cell TwoCouple = StrTwoCouple.createCell(0);

    private Row StrThreeCouple = Teacher.createRow(3);
    private Cell ThreeCouple = StrThreeCouple.createCell(0);

    private Row StrFourCouple = Teacher.createRow(4);
    private Cell FourCouple = StrFourCouple.createCell(0);

    private Row StrFiveCouple = Teacher.createRow(5);
    private Cell FiveCouple = StrFiveCouple.createCell(0);

    private Row StrSixCouple = Teacher.createRow(6);
    private Cell SixCouple = StrSixCouple.createCell(0);

    private Row StrSevenCouple = Teacher.createRow(7);
    private Cell SevenCouple = StrSevenCouple.createCell(0);

    public void CreatorTimeTableTeacherOne(Teacher teacher) throws IOException {

        String CoupleCell = "";
        int HeightPoints = 250;
        int ColumnWidth = 10000;

        CellStyle cellStyle = workbookTeacher.createCellStyle();
        cellStyle.setWrapText(true);

        ArrayList<CoupleTeacher> ArrayMonday = new ArrayList<>();
        ArrayList<CoupleTeacher> ArrayTuesday = new ArrayList<>();
        ArrayList<CoupleTeacher> ArrayWednesday = new ArrayList<>();
        ArrayList<CoupleTeacher> ArrayThursday = new ArrayList<>();
        ArrayList<CoupleTeacher> ArrayFriday = new ArrayList<>();
        ArrayList<CoupleTeacher> ArraySaturday = new ArrayList<>();

        ArrayList<CoupleTeacher> ArrayCouple = teacher.GetArrayCoupleTeacher();

        cellTeacherDayWeekAndCoupleNumber.setCellStyle(cellStyle);
        cellTeacherDayWeekAndCoupleNumber.setCellValue("День недели" + "\n" + "Номер пары");
        Teacher.setColumnWidth(0,8000);
        rowTeacherZero.setHeightInPoints(30);

        OneCouple.setCellValue("1 пара");
        TwoCouple.setCellValue("2 пара");
        ThreeCouple.setCellValue("3 пара");
        FourCouple.setCellValue("4 пара");
        FiveCouple.setCellValue("5 пара");
        SixCouple.setCellValue("6 пара");
        SevenCouple.setCellValue("7 пара");

        Monday.setCellValue("Понедельник");
        Teacher.setColumnWidth(1,8000);

        Tuesday.setCellValue("Вторник");
        Teacher.setColumnWidth(2,8000);

        Wednesday.setCellValue("Среде");
        Teacher.setColumnWidth(3,8000);

        Thursday.setCellValue("Четверг");
        Teacher.setColumnWidth(4,8000);

        Friday.setCellValue("Пятница");
        Teacher.setColumnWidth(5,8000);

        Saturday.setCellValue("Суббота");
        Teacher.setColumnWidth(6,8000);

        for(int i = 0; i < ArrayCouple.size(); i++){

            int IdDay = ArrayCouple.get(i).GetIdDay();

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

        // Couple Monday

        String CoupleOneMonday = "";
        String CoupleTwoMonday = "";
        String CoupleThreeMonday = "";
        String CoupleFourMonday = "";
        String CoupleFiveMonday = "";
        String CoupleSixMonday = "";
        String CoupleSevenMonday = "";

        for(int i = 0; i < ArrayMonday.size(); i++){

            switch (ArrayMonday.get(i).GetNumberCouple()) {
                case (1): {
                    CoupleOneMonday = CoupleOneMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetType() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetGroupName() + " " + ArrayMonday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (2): {
                    CoupleTwoMonday = CoupleTwoMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetType() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetGroupName() + " " + ArrayMonday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (3): {
                    CoupleThreeMonday = CoupleThreeMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetType() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetGroupName() + " " + ArrayMonday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (4): {
                    CoupleFourMonday = CoupleFourMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetType() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetGroupName() + " " + ArrayMonday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (5): {
                    CoupleFiveMonday = CoupleFiveMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetType() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetGroupName() + " " + ArrayMonday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (6): {
                    CoupleSixMonday = CoupleSixMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetType() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetGroupName() + " " + ArrayMonday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (7): {
                    CoupleSevenMonday = CoupleSevenMonday + ArrayMonday.get(i).GetDiscipline() + " (" + ArrayMonday.get(i).GetType() + ")\n" + ArrayMonday.get(i).GetNumberWeek() + " " + ArrayMonday.get(i).GetGroupName() + " " + ArrayMonday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
            }
        }


        Cell cellOneCoupleMonday = StrOneCouple.createCell(1);
        cellOneCoupleMonday.setCellValue(CoupleOneMonday);
        cellOneCoupleMonday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(1, ColumnWidth);
        StrOneCouple.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleMonday = StrTwoCouple.createCell(1);
        cellTwoCoupleMonday.setCellValue(CoupleTwoMonday);
        cellTwoCoupleMonday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(1, ColumnWidth);
        StrTwoCouple.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleMonday = StrThreeCouple.createCell(1);
        cellThreeCoupleMonday.setCellValue(CoupleThreeMonday);
        cellThreeCoupleMonday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(1, ColumnWidth);
        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleMonday = StrFourCouple.createCell(1);
        cellFourCoupleMonday.setCellValue(CoupleFourMonday);
        cellFourCoupleMonday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(1, ColumnWidth);
        StrFourCouple.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleMonday = StrFiveCouple.createCell(1);
        cellFiveCoupleMonday.setCellValue(CoupleFiveMonday);
        cellFiveCoupleMonday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(1, ColumnWidth);
        StrFiveCouple.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleMonday = StrSixCouple.createCell(1);
        cellSixCoupleMonday.setCellValue(CoupleSixMonday);
        cellSixCoupleMonday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(1, ColumnWidth);
        StrSixCouple.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleMonday = StrSevenCouple.createCell(1);
        cellSevenCoupleMonday.setCellValue(CoupleSevenMonday);
        cellSevenCoupleMonday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(1, ColumnWidth);
        StrSevenCouple.setHeightInPoints(HeightPoints);

        // Couple Tuesday

        String CoupleOneTuesday = "";
        String CoupleTwoTuesday = "";
        String CoupleThreeTuesday = "";
        String CoupleFourTuesday = "";
        String CoupleFiveTuesday = "";
        String CoupleSixTuesday = "";
        String CoupleSevenTuesday = "";

        for(int i = 0; i < ArrayTuesday.size(); i++){

            switch (ArrayTuesday.get(i).GetNumberCouple()) {
                case (1): {
                    CoupleOneTuesday = CoupleOneTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetType() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetGroupName() + " " + ArrayTuesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (2): {
                    CoupleTwoTuesday = CoupleTwoTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetType() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetGroupName() + " " + ArrayTuesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (3): {
                    CoupleThreeTuesday = CoupleThreeTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetType() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetGroupName() + " " + ArrayTuesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (4): {
                    CoupleFourTuesday = CoupleFourTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetType() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetGroupName() + " " + ArrayTuesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (5): {
                    CoupleFiveTuesday = CoupleFiveTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetType() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetGroupName() + " " + ArrayTuesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (6): {
                    CoupleSixTuesday = CoupleSixTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetType() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetGroupName() + " " + ArrayTuesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (7): {
                    CoupleSevenTuesday = CoupleSevenTuesday + ArrayTuesday.get(i).GetDiscipline() + " (" + ArrayTuesday.get(i).GetType() + ")\n" + ArrayTuesday.get(i).GetNumberWeek() + " " + ArrayTuesday.get(i).GetGroupName() + " " + ArrayTuesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
            }
        }

        Cell cellOneCoupleTuesday = StrOneCouple.createCell(2);
        cellOneCoupleTuesday.setCellValue(CoupleOneTuesday);
        cellOneCoupleTuesday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(2, ColumnWidth);
        StrOneCouple.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleTuesday = StrTwoCouple.createCell(2);
        cellTwoCoupleTuesday.setCellValue(CoupleTwoTuesday);
        cellTwoCoupleTuesday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(2, ColumnWidth);
        StrTwoCouple.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleTuesday = StrThreeCouple.createCell(2);
        cellThreeCoupleTuesday.setCellValue(CoupleThreeTuesday);
        cellThreeCoupleTuesday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(2, ColumnWidth);
        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleTuesday = StrFourCouple.createCell(2);
        cellFourCoupleTuesday.setCellValue(CoupleFourTuesday);
        cellFourCoupleTuesday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(2, ColumnWidth);
        StrFourCouple.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleTuesday = StrFiveCouple.createCell(2);
        cellFiveCoupleTuesday.setCellValue(CoupleFiveTuesday);
        cellFiveCoupleTuesday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(2, ColumnWidth);
        StrFiveCouple.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleTuesday = StrSixCouple.createCell(2);
        cellSixCoupleTuesday.setCellValue(CoupleSixTuesday);
        cellSixCoupleTuesday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(2, ColumnWidth);
        StrSixCouple.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleTuesday = StrSevenCouple.createCell(2);
        cellSevenCoupleTuesday.setCellValue(CoupleSevenTuesday);
        cellSevenCoupleTuesday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(2, ColumnWidth);
        StrSevenCouple.setHeightInPoints(HeightPoints);


        // Couple Wednesday

        String CoupleOneWednesday = "";
        String CoupleTwoWednesday = "";
        String CoupleThreeWednesday = "";
        String CoupleFourWednesday = "";
        String CoupleFiveWednesday = "";
        String CoupleSixWednesday = "";
        String CoupleSevenWednesday = "";

        for(int i = 0; i < ArrayWednesday.size(); i++){

            switch (ArrayWednesday.get(i).GetNumberCouple()) {
                case (1): {
                    CoupleOneWednesday = CoupleOneWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetType() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetGroupName() + " " + ArrayWednesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (2): {
                    CoupleTwoWednesday = CoupleTwoWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetType() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetGroupName() + " " + ArrayWednesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (3): {
                    CoupleThreeWednesday = CoupleThreeWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetType() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetGroupName() + " " + ArrayWednesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (4): {
                    CoupleFourWednesday = CoupleFourWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetType() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetGroupName() + " " + ArrayWednesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (5): {
                    CoupleFiveWednesday = CoupleFiveWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetType() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetGroupName() + " " + ArrayWednesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (6): {
                    CoupleSixWednesday = CoupleSixWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetType() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetGroupName() + " " + ArrayWednesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (7): {
                    CoupleSevenWednesday = CoupleSevenWednesday + ArrayWednesday.get(i).GetDiscipline() + " (" + ArrayWednesday.get(i).GetType() + ")\n" + ArrayWednesday.get(i).GetNumberWeek() + " " + ArrayWednesday.get(i).GetGroupName() + " " + ArrayWednesday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
            }
        }

        Cell cellOneCoupleWednesday = StrOneCouple.createCell(3);
        cellOneCoupleWednesday.setCellValue(CoupleOneWednesday);
        cellOneCoupleWednesday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(3, ColumnWidth);
        StrOneCouple.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleWednesday = StrTwoCouple.createCell(3);
        cellTwoCoupleWednesday.setCellValue(CoupleTwoWednesday);
        cellTwoCoupleWednesday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(3, ColumnWidth);
        StrTwoCouple.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleWednesday = StrThreeCouple.createCell(3);
        cellThreeCoupleWednesday.setCellValue(CoupleThreeWednesday);
        cellThreeCoupleWednesday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(3, ColumnWidth);
        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleWednesday = StrFourCouple.createCell(3);
        cellFourCoupleWednesday.setCellValue(CoupleFourWednesday);
        cellFourCoupleWednesday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(3, ColumnWidth);
        StrFourCouple.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleWednesday = StrFiveCouple.createCell(3);
        cellFiveCoupleWednesday.setCellValue(CoupleFiveWednesday);
        cellFiveCoupleWednesday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(3, ColumnWidth);
        StrFiveCouple.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleWednesday = StrSixCouple.createCell(3);
        cellSixCoupleWednesday.setCellValue(CoupleSixWednesday);
        cellSixCoupleWednesday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(3, ColumnWidth);
        StrSixCouple.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleWednesday = StrSevenCouple.createCell(3);
        cellSevenCoupleWednesday.setCellValue(CoupleSevenWednesday);
        cellSevenCoupleWednesday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(3, ColumnWidth);
        StrSevenCouple.setHeightInPoints(HeightPoints);

        // Couple Thursday

        String CoupleOneThursday = "";
        String CoupleTwoThursday = "";
        String CoupleThreeThursday = "";
        String CoupleFourThursday= "";
        String CoupleFiveThursday = "";
        String CoupleSixThursday = "";
        String CoupleSevenThursday = "";

        for(int i = 0; i < ArrayThursday.size(); i++){

            switch (ArrayThursday.get(i).GetNumberCouple()) {
                case (1): {
                    CoupleOneThursday = CoupleOneThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetType() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetGroupName() + " " + ArrayThursday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (2): {
                    CoupleTwoThursday = CoupleTwoThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetType() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetGroupName() + " " + ArrayThursday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (3): {
                    CoupleThreeThursday = CoupleThreeThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetType() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetGroupName() + " " + ArrayThursday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (4): {
                    CoupleFourThursday = CoupleFourThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetType() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetGroupName() + " " + ArrayThursday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (5): {
                    CoupleFiveThursday = CoupleFiveThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetType() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetGroupName() + " " + ArrayThursday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (6): {
                    CoupleSixThursday = CoupleSixThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetType() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetGroupName() + " " + ArrayThursday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (7): {
                    CoupleSevenThursday = CoupleSevenThursday + ArrayThursday.get(i).GetDiscipline() + " (" + ArrayThursday.get(i).GetType() + ")\n" + ArrayThursday.get(i).GetNumberWeek() + " " + ArrayThursday.get(i).GetGroupName() + " " + ArrayThursday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
            }
        }

        Cell cellOneCoupleThursday = StrOneCouple.createCell(4);
        cellOneCoupleThursday.setCellValue(CoupleOneThursday);
        cellOneCoupleThursday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(4, ColumnWidth);
        StrOneCouple.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleThursday = StrTwoCouple.createCell(4);
        cellTwoCoupleThursday.setCellValue(CoupleTwoThursday);
        cellTwoCoupleThursday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(4, ColumnWidth);
        StrTwoCouple.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleThursday = StrThreeCouple.createCell(4);
        cellThreeCoupleThursday.setCellValue(CoupleThreeThursday);
        cellThreeCoupleThursday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(4, ColumnWidth);
        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleThursday = StrFourCouple.createCell(4);
        cellFourCoupleThursday.setCellValue(CoupleFourThursday);
        cellFourCoupleThursday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(4, ColumnWidth);
        StrFourCouple.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleThursday = StrFiveCouple.createCell(4);
        cellFiveCoupleThursday.setCellValue(CoupleFiveThursday);
        cellFiveCoupleThursday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(4, ColumnWidth);
        StrFiveCouple.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleThursday = StrSixCouple.createCell(4);
        cellSixCoupleThursday.setCellValue(CoupleSixThursday);
        cellSixCoupleThursday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(4, ColumnWidth);
        StrSixCouple.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleThursday = StrSevenCouple.createCell(4);
        cellSevenCoupleThursday.setCellValue(CoupleSevenThursday);
        cellSevenCoupleThursday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(4, ColumnWidth);
        StrSevenCouple.setHeightInPoints(HeightPoints);

        // Couple Friday

        String CoupleOneFriday = "";
        String CoupleTwoFriday = "";
        String CoupleThreeFriday = "";
        String CoupleFourFriday = "";
        String CoupleFiveFriday = "";
        String CoupleSixFriday = "";
        String CoupleSevenFriday = "";

        for(int i = 0; i < ArrayFriday.size(); i++){

            switch (ArrayFriday.get(i).GetNumberCouple()) {
                case (1): {
                    CoupleOneFriday = CoupleOneFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetType() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetGroupName() + " " + ArrayFriday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (2): {
                    CoupleTwoFriday = CoupleTwoFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetType() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetGroupName() + " " + ArrayFriday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (3): {
                    CoupleThreeFriday = CoupleThreeFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetType() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetGroupName() + " " + ArrayFriday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (4): {
                    CoupleFourFriday = CoupleFourFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetType() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetGroupName() + " " + ArrayFriday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (5): {
                    CoupleFiveFriday = CoupleFiveFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetType() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetGroupName() + " " + ArrayFriday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (6): {
                    CoupleSixFriday = CoupleSixFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetType() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetGroupName() + " " + ArrayFriday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (7): {
                    CoupleSevenFriday = CoupleSevenFriday + ArrayFriday.get(i).GetDiscipline() + " (" + ArrayFriday.get(i).GetType() + ")\n" + ArrayFriday.get(i).GetNumberWeek() + " " + ArrayFriday.get(i).GetGroupName() + " " + ArrayFriday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
            }
        }

        Cell cellOneCoupleFriday = StrOneCouple.createCell(5);
        cellOneCoupleFriday.setCellValue(CoupleOneFriday);
        cellOneCoupleFriday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(5, ColumnWidth);
        StrOneCouple.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleFriday = StrTwoCouple.createCell(5);
        cellTwoCoupleFriday.setCellValue(CoupleTwoFriday);
        cellTwoCoupleFriday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(5, ColumnWidth);
        StrTwoCouple.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleFriday = StrThreeCouple.createCell(5);
        cellThreeCoupleFriday.setCellValue(CoupleThreeFriday);
        cellThreeCoupleFriday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(5, ColumnWidth);
        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleFriday = StrFourCouple.createCell(5);
        cellFourCoupleFriday.setCellValue(CoupleFourFriday);
        cellFourCoupleFriday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(5, ColumnWidth);
        StrFourCouple.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleFriday = StrFiveCouple.createCell(5);
        cellFiveCoupleFriday.setCellValue(CoupleFiveFriday);
        cellFiveCoupleFriday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(5, ColumnWidth);
        StrFiveCouple.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleFriday = StrSixCouple.createCell(5);
        cellSixCoupleFriday.setCellValue(CoupleSixFriday);
        cellSixCoupleFriday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(5, ColumnWidth);
        StrSixCouple.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleFriday = StrSevenCouple.createCell(5);
        cellSevenCoupleFriday.setCellValue(CoupleSevenFriday);
        cellSevenCoupleFriday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(5, ColumnWidth);
        StrSevenCouple.setHeightInPoints(HeightPoints);

        // Couple Saturday

        String CoupleOneSaturday = "";
        String CoupleTwoSaturday = "";
        String CoupleThreeSaturday = "";
        String CoupleFourSaturday = "";
        String CoupleFiveSaturday = "";
        String CoupleSixSaturday = "";
        String CoupleSevenSaturday = "";

        for(int i = 0; i < ArraySaturday.size(); i++){

            switch (ArraySaturday.get(i).GetNumberCouple()) {
                case (1): {
                    CoupleOneSaturday = CoupleOneSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetType() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetGroupName() + " " + ArraySaturday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (2): {
                    CoupleTwoSaturday = CoupleTwoSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetType() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetGroupName() + " " + ArraySaturday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (3): {
                    CoupleThreeSaturday = CoupleThreeSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetType() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetGroupName() + " " + ArraySaturday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (4): {
                    CoupleFourSaturday = CoupleFourSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetType() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetGroupName() + " " + ArraySaturday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (5): {
                    CoupleFiveSaturday = CoupleFiveSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetType() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetGroupName() + " " + ArraySaturday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (6): {
                    CoupleSixSaturday = CoupleSixSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetType() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetGroupName() + " " + ArraySaturday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
                case (7): {
                    CoupleSevenSaturday = CoupleSevenSaturday + ArraySaturday.get(i).GetDiscipline() + " (" + ArraySaturday.get(i).GetType() + ")\n" + ArraySaturday.get(i).GetNumberWeek() + " " + ArraySaturday.get(i).GetGroupName() + " " + ArraySaturday.get(i).GetAud() + "\n" + "\n";
                    break;
                }
            }
        }

        Cell cellOneCoupleSaturday = StrOneCouple.createCell(6);
        cellOneCoupleSaturday.setCellValue(CoupleOneSaturday);
        cellOneCoupleSaturday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(6, ColumnWidth);
        StrOneCouple.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleSaturday = StrTwoCouple.createCell(6);
        cellTwoCoupleSaturday.setCellValue(CoupleTwoSaturday);
        cellTwoCoupleSaturday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(6, ColumnWidth);
        StrTwoCouple.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleSaturday = StrThreeCouple.createCell(6);
        cellThreeCoupleSaturday.setCellValue(CoupleThreeSaturday);
        cellThreeCoupleSaturday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(6, ColumnWidth);
        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleSaturday = StrFourCouple.createCell(6);
        cellFourCoupleSaturday.setCellValue(CoupleFourSaturday);
        cellFourCoupleSaturday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(6, ColumnWidth);
        StrFourCouple.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleSaturday = StrFiveCouple.createCell(6);
        cellFiveCoupleSaturday.setCellValue(CoupleFiveSaturday);
        cellFiveCoupleSaturday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(6, ColumnWidth);
        StrFiveCouple.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleSaturday = StrSixCouple.createCell(6);
        cellSixCoupleSaturday.setCellValue(CoupleSixSaturday);
        cellSixCoupleSaturday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(6, ColumnWidth);
        StrSixCouple.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleSaturday = StrSevenCouple.createCell(6);
        cellSevenCoupleSaturday.setCellValue(CoupleSevenSaturday);
        cellSevenCoupleSaturday.setCellStyle(cellStyle);
        Teacher.setColumnWidth(6, ColumnWidth);
        StrSevenCouple.setHeightInPoints(HeightPoints);

        FileOutputStream fileOutputStream = new FileOutputStream("OneTeacherExelDoc");

        workbookTeacher.write(fileOutputStream);

        fileOutputStream.close();

    }

}
