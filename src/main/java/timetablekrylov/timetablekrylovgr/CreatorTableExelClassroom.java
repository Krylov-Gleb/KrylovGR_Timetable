package timetablekrylov.timetablekrylovgr;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.channels.ClosedSelectorException;
import java.util.ArrayList;

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

    public CreatorTableExelClassroom() throws FileNotFoundException {
    }

    public void CreatorTimeTableClassroomOne(ArrayList<CoupleGroup> Array) throws IOException {

        int HeightPoints = 250;
        int ColumnWidth = 10000;

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


        Cell cellOneCoupleMonday = StrOneCouple.createCell(1);
        cellOneCoupleMonday.setCellValue(CoupleOneMonday);
        cellOneCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(1, ColumnWidth);
        StrOneCouple.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleMonday = StrTwoCouple.createCell(1);
        cellTwoCoupleMonday.setCellValue(CoupleTwoMonday);
        cellTwoCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(1, ColumnWidth);
        StrTwoCouple.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleMonday = StrThreeCouple.createCell(1);
        cellThreeCoupleMonday.setCellValue(CoupleThreeMonday);
        cellThreeCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(1, ColumnWidth);
        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleMonday = StrFourCouple.createCell(1);
        cellFourCoupleMonday.setCellValue(CoupleFourMonday);
        cellFourCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(1, ColumnWidth);
        StrFourCouple.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleMonday = StrFiveCouple.createCell(1);
        cellFiveCoupleMonday.setCellValue(CoupleFiveMonday);
        cellFiveCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(1, ColumnWidth);
        StrFiveCouple.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleMonday = StrSixCouple.createCell(1);
        cellSixCoupleMonday.setCellValue(CoupleSixMonday);
        cellSixCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(1, ColumnWidth);
        StrSixCouple.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleMonday = StrSevenCouple.createCell(1);
        cellSevenCoupleMonday.setCellValue(CoupleSevenMonday);
        cellSevenCoupleMonday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(1, ColumnWidth);
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

        Cell cellOneCoupleTuesday = StrOneCouple.createCell(2);
        cellOneCoupleTuesday.setCellValue(CoupleOneTuesday);
        cellOneCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(2, ColumnWidth);
        StrOneCouple.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleTuesday = StrTwoCouple.createCell(2);
        cellTwoCoupleTuesday.setCellValue(CoupleTwoTuesday);
        cellTwoCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(2, ColumnWidth);
        StrTwoCouple.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleTuesday = StrThreeCouple.createCell(2);
        cellThreeCoupleTuesday.setCellValue(CoupleThreeTuesday);
        cellThreeCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(2, ColumnWidth);
        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleTuesday = StrFourCouple.createCell(2);
        cellFourCoupleTuesday.setCellValue(CoupleFourTuesday);
        cellFourCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(2, ColumnWidth);
        StrFourCouple.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleTuesday = StrFiveCouple.createCell(2);
        cellFiveCoupleTuesday.setCellValue(CoupleFiveTuesday);
        cellFiveCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(2, ColumnWidth);
        StrFiveCouple.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleTuesday = StrSixCouple.createCell(2);
        cellSixCoupleTuesday.setCellValue(CoupleSixTuesday);
        cellSixCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(2, ColumnWidth);
        StrSixCouple.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleTuesday = StrSevenCouple.createCell(2);
        cellSevenCoupleTuesday.setCellValue(CoupleSevenTuesday);
        cellSevenCoupleTuesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(2, ColumnWidth);
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

        Cell cellOneCoupleWednesday = StrOneCouple.createCell(3);
        cellOneCoupleWednesday.setCellValue(CoupleOneWednesday);
        cellOneCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(3, ColumnWidth);
        StrOneCouple.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleWednesday = StrTwoCouple.createCell(3);
        cellTwoCoupleWednesday.setCellValue(CoupleTwoWednesday);
        cellTwoCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(3, ColumnWidth);
        StrTwoCouple.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleWednesday = StrThreeCouple.createCell(3);
        cellThreeCoupleWednesday.setCellValue(CoupleThreeWednesday);
        cellThreeCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(3, ColumnWidth);
        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleWednesday = StrFourCouple.createCell(3);
        cellFourCoupleWednesday.setCellValue(CoupleFourWednesday);
        cellFourCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(3, ColumnWidth);
        StrFourCouple.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleWednesday = StrFiveCouple.createCell(3);
        cellFiveCoupleWednesday.setCellValue(CoupleFiveWednesday);
        cellFiveCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(3, ColumnWidth);
        StrFiveCouple.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleWednesday = StrSixCouple.createCell(3);
        cellSixCoupleWednesday.setCellValue(CoupleSixWednesday);
        cellSixCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(3, ColumnWidth);
        StrSixCouple.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleWednesday = StrSevenCouple.createCell(3);
        cellSevenCoupleWednesday.setCellValue(CoupleSevenWednesday);
        cellSevenCoupleWednesday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(3, ColumnWidth);
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

        Cell cellOneCoupleThursday = StrOneCouple.createCell(4);
        cellOneCoupleThursday.setCellValue(CoupleOneThursday);
        cellOneCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(4, ColumnWidth);
        StrOneCouple.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleThursday = StrTwoCouple.createCell(4);
        cellTwoCoupleThursday.setCellValue(CoupleTwoThursday);
        cellTwoCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(4, ColumnWidth);
        StrTwoCouple.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleThursday = StrThreeCouple.createCell(4);
        cellThreeCoupleThursday.setCellValue(CoupleThreeThursday);
        cellThreeCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(4, ColumnWidth);
        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleThursday = StrFourCouple.createCell(4);
        cellFourCoupleThursday.setCellValue(CoupleFourThursday);
        cellFourCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(4, ColumnWidth);
        StrFourCouple.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleThursday = StrFiveCouple.createCell(4);
        cellFiveCoupleThursday.setCellValue(CoupleFiveThursday);
        cellFiveCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(4, ColumnWidth);
        StrFiveCouple.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleThursday = StrSixCouple.createCell(4);
        cellSixCoupleThursday.setCellValue(CoupleSixThursday);
        cellSixCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(4, ColumnWidth);
        StrSixCouple.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleThursday = StrSevenCouple.createCell(4);
        cellSevenCoupleThursday.setCellValue(CoupleSevenThursday);
        cellSevenCoupleThursday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(4, ColumnWidth);
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

        Cell cellOneCoupleFriday = StrOneCouple.createCell(5);
        cellOneCoupleFriday.setCellValue(CoupleOneFriday);
        cellOneCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(5, ColumnWidth);
        StrOneCouple.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleFriday = StrTwoCouple.createCell(5);
        cellTwoCoupleFriday.setCellValue(CoupleTwoFriday);
        cellTwoCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(5, ColumnWidth);
        StrTwoCouple.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleFriday = StrThreeCouple.createCell(5);
        cellThreeCoupleFriday.setCellValue(CoupleThreeFriday);
        cellThreeCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(5, ColumnWidth);
        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleFriday = StrFourCouple.createCell(5);
        cellFourCoupleFriday.setCellValue(CoupleFourFriday);
        cellFourCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(5, ColumnWidth);
        StrFourCouple.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleFriday = StrFiveCouple.createCell(5);
        cellFiveCoupleFriday.setCellValue(CoupleFiveFriday);
        cellFiveCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(5, ColumnWidth);
        StrFiveCouple.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleFriday = StrSixCouple.createCell(5);
        cellSixCoupleFriday.setCellValue(CoupleSixFriday);
        cellSixCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(5, ColumnWidth);
        StrSixCouple.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleFriday = StrSevenCouple.createCell(5);
        cellSevenCoupleFriday.setCellValue(CoupleSevenFriday);
        cellSevenCoupleFriday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(5, ColumnWidth);
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

        Cell cellOneCoupleSaturday = StrOneCouple.createCell(6);
        cellOneCoupleSaturday.setCellValue(CoupleOneSaturday);
        cellOneCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(6, ColumnWidth);
        StrOneCouple.setHeightInPoints(HeightPoints);

        Cell cellTwoCoupleSaturday = StrTwoCouple.createCell(6);
        cellTwoCoupleSaturday.setCellValue(CoupleTwoSaturday);
        cellTwoCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(6, ColumnWidth);
        StrTwoCouple.setHeightInPoints(HeightPoints);

        Cell cellThreeCoupleSaturday = StrThreeCouple.createCell(6);
        cellThreeCoupleSaturday.setCellValue(CoupleThreeSaturday);
        cellThreeCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(6, ColumnWidth);
        StrThreeCouple.setHeightInPoints(HeightPoints);

        Cell cellFourCoupleSaturday = StrFourCouple.createCell(6);
        cellFourCoupleSaturday.setCellValue(CoupleFourSaturday);
        cellFourCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(6, ColumnWidth);
        StrFourCouple.setHeightInPoints(HeightPoints);

        Cell cellFiveCoupleSaturday = StrFiveCouple.createCell(6);
        cellFiveCoupleSaturday.setCellValue(CoupleFiveSaturday);
        cellFiveCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(6, ColumnWidth);
        StrFiveCouple.setHeightInPoints(HeightPoints);

        Cell cellSixCoupleSaturday = StrSixCouple.createCell(6);
        cellSixCoupleSaturday.setCellValue(CoupleSixSaturday);
        cellSixCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(6, ColumnWidth);
        StrSixCouple.setHeightInPoints(HeightPoints);

        Cell cellSevenCoupleSaturday = StrSevenCouple.createCell(6);
        cellSevenCoupleSaturday.setCellValue(CoupleSevenSaturday);
        cellSevenCoupleSaturday.setCellStyle(cellStyle);
        Classroom.setColumnWidth(6, ColumnWidth);
        StrSevenCouple.setHeightInPoints(HeightPoints);


        FileOutputStream fileOutputStream = new FileOutputStream("OneClassroomExelDoc");

        workbookOneClassroom.write(fileOutputStream);
        fileOutputStream.close();

    }
}
