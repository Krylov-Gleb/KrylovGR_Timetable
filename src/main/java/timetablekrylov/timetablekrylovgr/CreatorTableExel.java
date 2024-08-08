package timetablekrylov.timetablekrylovgr;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class CreatorTableExel {

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

        CellStyle cellStyle = workbookOneElementGroup.createCellStyle();
        cellStyle.setWrapText(true);

        ArrayList<CoupleGroup> ArrayCoupleOne = new ArrayList<>();
        ArrayList<CoupleGroup> ArrayCoupleTwo = new ArrayList<>();
        ArrayList<CoupleGroup> ArrayCoupleThree = new ArrayList<>();
        ArrayList<CoupleGroup> ArrayCoupleFour = new ArrayList<>();
        ArrayList<CoupleGroup> ArrayCoupleFive = new ArrayList<>();
        ArrayList<CoupleGroup> ArrayCoupleSix = new ArrayList<>();


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
                    ArrayCoupleOne.add(ArrayCouple.get(i));
                    break;
                }
                case (2):{
                    ArrayCoupleTwo.add(ArrayCouple.get(i));
                    break;
                }
                case (3):{
                    ArrayCoupleThree.add(ArrayCouple.get(i));
                    break;
                }
                case (4):{
                    ArrayCoupleFour.add(ArrayCouple.get(i));
                    break;
                }
                case (5):{
                    ArrayCoupleFive.add(ArrayCouple.get(i));
                    break;
                }
                case (6):{
                    ArrayCoupleSix.add(ArrayCouple.get(i));
                    break;
                }
            }
        }

        for(int i = 0; i < ArrayCoupleOne.size(); i++) {

            Cell cell;

            String CoupleCell = "";

            CoupleCell = ArrayCoupleOne.get(i).GetDiscipline() + " (" + ArrayCoupleOne.get(i).GetCoupleType() + ")\n" + ArrayCoupleOne.get(i).GetNumberWeek() + " " + ArrayCoupleOne.get(i).GetTeacherName() + " " + ArrayCoupleOne.get(i).GetAud();

            switch (ArrayCoupleOne.get(i).GetCoupleNumber()){
                case(1):{
                    cell = rowGroupOne.createCell(ArrayCoupleOne.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleOne.get(i).GetIDDay(), 10000);
                    rowGroupOne.setHeightInPoints(90);
                    break;
                }
                case(2):{
                    cell = rowGroupTwo.createCell(ArrayCoupleOne.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleOne.get(i).GetIDDay(), 10000);
                    rowGroupTwo.setHeightInPoints(90);
                    break;
                }
                case(3):{
                    cell = rowGroupThree.createCell(ArrayCoupleOne.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleOne.get(i).GetIDDay(), 10000);
                    rowGroupThree.setHeightInPoints(90);
                    break;
                }
                case(4):{
                    cell = rowGroupFour.createCell(ArrayCoupleOne.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleOne.get(i).GetIDDay(), 10000);
                    rowGroupFour.setHeightInPoints(90);
                    break;
                }
                case(5):{
                    cell = rowGroupFive.createCell(ArrayCoupleOne.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleOne.get(i).GetIDDay(), 10000);
                    rowGroupFive.setHeightInPoints(90);
                    break;
                }
                case(6):{
                    cell = rowGroupSix.createCell(ArrayCoupleOne.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleOne.get(i).GetIDDay(), 10000);
                    rowGroupSix.setHeightInPoints(90);
                    break;
                }
                case(7):{
                    cell = rowGroupSeven.createCell(ArrayCoupleOne.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleOne.get(i).GetIDDay(), 10000);
                    rowGroupSeven.setHeightInPoints(90);
                    break;
                }
            }
        }

        for(int i = 0; i < ArrayCoupleOne.size(); i++) {

            Cell cell;

            String CoupleCell = "";

            if (i != ArrayCoupleOne.size()-1) {
                if (ArrayCoupleOne.get(i).GetCoupleNumber() == ArrayCoupleOne.get(i + 1).GetCoupleNumber()) {
                    CoupleCell = ArrayCoupleOne.get(i).GetDiscipline() + " (" + ArrayCoupleOne.get(i).GetCoupleType() + ")\n" + ArrayCoupleOne.get(i).GetNumberWeek() + " " + ArrayCoupleOne.get(i).GetTeacherName() + " " + ArrayCoupleOne.get(i).GetAud() + "\n" + "\n" + ArrayCoupleOne.get(i + 1).GetDiscipline() + " (" + ArrayCoupleOne.get(i + 1).GetCoupleType() + ")\n" + ArrayCoupleOne.get(i + 1).GetNumberWeek() + " " + ArrayCoupleOne.get(i + 1).GetTeacherName() + " " + ArrayCoupleOne.get(i + 1).GetAud();

                    switch (ArrayCoupleOne.get(i).GetCoupleNumber()){
                        case(1):{
                            cell = rowGroupOne.createCell(ArrayCoupleOne.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleOne.get(i).GetIDDay(), 10000);
                            rowGroupOne.setHeightInPoints(90);
                            break;
                        }
                        case(2):{
                            cell = rowGroupTwo.createCell(ArrayCoupleOne.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleOne.get(i).GetIDDay(), 10000);
                            rowGroupTwo.setHeightInPoints(90);
                            break;
                        }
                        case(3):{
                            cell = rowGroupThree.createCell(ArrayCoupleOne.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleOne.get(i).GetIDDay(), 10000);
                            rowGroupThree.setHeightInPoints(90);
                            break;
                        }
                        case(4):{
                            cell = rowGroupFour.createCell(ArrayCoupleOne.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleOne.get(i).GetIDDay(), 10000);
                            rowGroupFour.setHeightInPoints(90);
                            break;
                        }
                        case(5):{
                            cell = rowGroupFive.createCell(ArrayCoupleOne.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleOne.get(i).GetIDDay(), 10000);
                            rowGroupFive.setHeightInPoints(90);
                            break;
                        }
                        case(6):{
                            cell = rowGroupSix.createCell(ArrayCoupleOne.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleOne.get(i).GetIDDay(), 10000);
                            rowGroupSix.setHeightInPoints(90);
                            break;
                        }
                        case(7):{
                            cell = rowGroupSeven.createCell(ArrayCoupleOne.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleOne.get(i).GetIDDay(), 10000);
                            rowGroupSeven.setHeightInPoints(90);
                            break;
                        }
                    }
                }
            }
        }

        // CellTwo

        for(int i = 0; i < ArrayCoupleTwo.size(); i++) {

            Cell cell;

            String CoupleCell = "";

            CoupleCell = ArrayCoupleTwo.get(i).GetDiscipline() + " (" + ArrayCoupleTwo.get(i).GetCoupleType() + ")\n" + ArrayCoupleTwo.get(i).GetNumberWeek() + " " + ArrayCoupleTwo.get(i).GetTeacherName() + " " + ArrayCoupleTwo.get(i).GetAud();

            switch (ArrayCoupleTwo.get(i).GetCoupleNumber()){
                case(1):{
                    cell = rowGroupOne.createCell(ArrayCoupleTwo.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleTwo.get(i).GetIDDay(), 10000);
                    rowGroupOne.setHeightInPoints(90);
                    break;
                }
                case(2):{
                    cell = rowGroupTwo.createCell(ArrayCoupleTwo.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleTwo.get(i).GetIDDay(), 10000);
                    rowGroupTwo.setHeightInPoints(90);
                    break;
                }
                case(3):{
                    cell = rowGroupThree.createCell(ArrayCoupleTwo.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleTwo.get(i).GetIDDay(), 10000);
                    rowGroupThree.setHeightInPoints(90);
                    break;
                }
                case(4):{
                    cell = rowGroupFour.createCell(ArrayCoupleTwo.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleTwo.get(i).GetIDDay(), 10000);
                    rowGroupFour.setHeightInPoints(90);
                    break;
                }
                case(5):{
                    cell = rowGroupFive.createCell(ArrayCoupleTwo.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleTwo.get(i).GetIDDay(), 10000);
                    rowGroupFive.setHeightInPoints(90);
                    break;
                }
                case(6):{
                    cell = rowGroupSix.createCell(ArrayCoupleTwo.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleTwo.get(i).GetIDDay(), 10000);
                    rowGroupSix.setHeightInPoints(90);
                    break;
                }
                case(7):{
                    cell = rowGroupSeven.createCell(ArrayCoupleTwo.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleTwo.get(i).GetIDDay(), 10000);
                    rowGroupSeven.setHeightInPoints(90);
                    break;
                }
            }
        }

        for(int i = 0; i < ArrayCoupleTwo.size(); i++) {

            Cell cell;

            String CoupleCell = "";

            if (i != ArrayCoupleTwo.size()-1) {
                if (ArrayCoupleTwo.get(i).GetCoupleNumber() == ArrayCoupleTwo.get(i + 1).GetCoupleNumber()) {
                    CoupleCell = ArrayCoupleTwo.get(i).GetDiscipline() + " (" + ArrayCoupleTwo.get(i).GetCoupleType() + ")\n" + ArrayCoupleTwo.get(i).GetNumberWeek() + " " + ArrayCoupleTwo.get(i).GetTeacherName() + " " + ArrayCoupleTwo.get(i).GetAud() + "\n" + "\n" + ArrayCoupleTwo.get(i + 1).GetDiscipline() + " (" + ArrayCoupleTwo.get(i + 1).GetCoupleType() + ")\n" + ArrayCoupleTwo.get(i + 1).GetNumberWeek() + " " + ArrayCoupleTwo.get(i + 1).GetTeacherName() + " " + ArrayCoupleTwo.get(i + 1).GetAud();

                    switch (ArrayCoupleTwo.get(i).GetCoupleNumber()){
                        case(1):{
                            cell = rowGroupOne.createCell(ArrayCoupleTwo.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleTwo.get(i).GetIDDay(), 10000);
                            rowGroupOne.setHeightInPoints(90);
                            break;
                        }
                        case(2):{
                            cell = rowGroupTwo.createCell(ArrayCoupleTwo.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleTwo.get(i).GetIDDay(), 10000);
                            rowGroupTwo.setHeightInPoints(90);
                            break;
                        }
                        case(3):{
                            cell = rowGroupThree.createCell(ArrayCoupleTwo.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleTwo.get(i).GetIDDay(), 10000);
                            rowGroupThree.setHeightInPoints(90);
                            break;
                        }
                        case(4):{
                            cell = rowGroupFour.createCell(ArrayCoupleTwo.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleTwo.get(i).GetIDDay(), 10000);
                            rowGroupFour.setHeightInPoints(90);
                            break;
                        }
                        case(5):{
                            cell = rowGroupFive.createCell(ArrayCoupleTwo.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleTwo.get(i).GetIDDay(), 10000);
                            rowGroupFive.setHeightInPoints(90);
                            break;
                        }
                        case(6):{
                            cell = rowGroupSix.createCell(ArrayCoupleTwo.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleTwo.get(i).GetIDDay(), 10000);
                            rowGroupSix.setHeightInPoints(90);
                            break;
                        }
                        case(7):{
                            cell = rowGroupSeven.createCell(ArrayCoupleTwo.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleTwo.get(i).GetIDDay(), 10000);
                            rowGroupSeven.setHeightInPoints(90);
                            break;
                        }
                    }
                }
            }
        }

        // CellThree

        for(int i = 0; i < ArrayCoupleThree.size(); i++) {

            Cell cell;

            String CoupleCell = "";

            CoupleCell = ArrayCoupleThree.get(i).GetDiscipline() + " (" + ArrayCoupleThree.get(i).GetCoupleType() + ")\n" + ArrayCoupleThree.get(i).GetNumberWeek() + " " + ArrayCoupleThree.get(i).GetTeacherName() + " " + ArrayCoupleThree.get(i).GetAud();

            switch (ArrayCoupleThree.get(i).GetCoupleNumber()){
                case(1):{
                    cell = rowGroupOne.createCell(ArrayCoupleThree.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleThree.get(i).GetIDDay(), 10000);
                    rowGroupOne.setHeightInPoints(90);
                    break;
                }
                case(2):{
                    cell = rowGroupTwo.createCell(ArrayCoupleThree.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleThree.get(i).GetIDDay(), 10000);
                    rowGroupTwo.setHeightInPoints(90);
                    break;
                }
                case(3):{
                    cell = rowGroupThree.createCell(ArrayCoupleThree.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleThree.get(i).GetIDDay(), 10000);
                    rowGroupThree.setHeightInPoints(90);
                    break;
                }
                case(4):{
                    cell = rowGroupFour.createCell(ArrayCoupleThree.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleThree.get(i).GetIDDay(), 10000);
                    rowGroupFour.setHeightInPoints(90);
                    break;
                }
                case(5):{
                    cell = rowGroupFive.createCell(ArrayCoupleThree.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleThree.get(i).GetIDDay(), 10000);
                    rowGroupFive.setHeightInPoints(90);
                    break;
                }
                case(6):{
                    cell = rowGroupSix.createCell(ArrayCoupleThree.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleThree.get(i).GetIDDay(), 10000);
                    rowGroupSix.setHeightInPoints(90);
                    break;
                }
                case(7):{
                    cell = rowGroupSeven.createCell(ArrayCoupleThree.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleThree.get(i).GetIDDay(), 10000);
                    rowGroupSeven.setHeightInPoints(90);
                    break;
                }
            }
        }

        for(int i = 0; i < ArrayCoupleThree.size(); i++) {

            Cell cell;

            String CoupleCell = "";

            if (i != ArrayCoupleThree.size()-1) {
                if (ArrayCoupleThree.get(i).GetCoupleNumber() == ArrayCoupleThree.get(i + 1).GetCoupleNumber()) {
                    CoupleCell = ArrayCoupleThree.get(i).GetDiscipline() + " (" + ArrayCoupleThree.get(i).GetCoupleType() + ")\n" + ArrayCoupleThree.get(i).GetNumberWeek() + " " + ArrayCoupleThree.get(i).GetTeacherName() + " " + ArrayCoupleThree.get(i).GetAud() + "\n" + "\n" + ArrayCoupleThree.get(i + 1).GetDiscipline() + " (" + ArrayCoupleThree.get(i + 1).GetCoupleType() + ")\n" + ArrayCoupleThree.get(i + 1).GetNumberWeek() + " " + ArrayCoupleThree.get(i + 1).GetTeacherName() + " " + ArrayCoupleThree.get(i + 1).GetAud();

                    switch (ArrayCoupleThree.get(i).GetCoupleNumber()){
                        case(1):{
                            cell = rowGroupOne.createCell(ArrayCoupleThree.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleThree.get(i).GetIDDay(), 10000);
                            rowGroupOne.setHeightInPoints(90);
                            break;
                        }
                        case(2):{
                            cell = rowGroupTwo.createCell(ArrayCoupleThree.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleThree.get(i).GetIDDay(), 10000);
                            rowGroupTwo.setHeightInPoints(90);
                            break;
                        }
                        case(3):{
                            cell = rowGroupThree.createCell(ArrayCoupleThree.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleThree.get(i).GetIDDay(), 10000);
                            rowGroupThree.setHeightInPoints(90);
                            break;
                        }
                        case(4):{
                            cell = rowGroupFour.createCell(ArrayCoupleThree.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleThree.get(i).GetIDDay(), 10000);
                            rowGroupFour.setHeightInPoints(90);
                            break;
                        }
                        case(5):{
                            cell = rowGroupFive.createCell(ArrayCoupleThree.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleThree.get(i).GetIDDay(), 10000);
                            rowGroupFive.setHeightInPoints(90);
                            break;
                        }
                        case(6):{
                            cell = rowGroupSix.createCell(ArrayCoupleThree.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleThree.get(i).GetIDDay(), 10000);
                            rowGroupSix.setHeightInPoints(90);
                            break;
                        }
                        case(7):{
                            cell = rowGroupSeven.createCell(ArrayCoupleThree.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleThree.get(i).GetIDDay(), 10000);
                            rowGroupSeven.setHeightInPoints(90);
                            break;
                        }
                    }
                }
            }
        }

        // CellFour

        for(int i = 0; i < ArrayCoupleFour.size(); i++) {

            Cell cell;

            String CoupleCell = "";

            CoupleCell = ArrayCoupleFour.get(i).GetDiscipline() + " (" + ArrayCoupleFour.get(i).GetCoupleType() + ")\n" + ArrayCoupleFour.get(i).GetNumberWeek() + " " + ArrayCoupleFour.get(i).GetTeacherName() + " " + ArrayCoupleFour.get(i).GetAud();

            switch (ArrayCoupleFour.get(i).GetCoupleNumber()){
                case(1):{
                    cell = rowGroupOne.createCell(ArrayCoupleFour.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleFour.get(i).GetIDDay(), 10000);
                    rowGroupOne.setHeightInPoints(90);
                    break;
                }
                case(2):{
                    cell = rowGroupTwo.createCell(ArrayCoupleFour.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleFour.get(i).GetIDDay(), 10000);
                    rowGroupTwo.setHeightInPoints(90);
                    break;
                }
                case(3):{
                    cell = rowGroupThree.createCell(ArrayCoupleFour.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleFour.get(i).GetIDDay(), 10000);
                    rowGroupThree.setHeightInPoints(90);
                    break;
                }
                case(4):{
                    cell = rowGroupFour.createCell(ArrayCoupleFour.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleFour.get(i).GetIDDay(), 10000);
                    rowGroupFour.setHeightInPoints(90);
                    break;
                }
                case(5):{
                    cell = rowGroupFive.createCell(ArrayCoupleFour.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleFour.get(i).GetIDDay(), 10000);
                    rowGroupFive.setHeightInPoints(90);
                    break;
                }
                case(6):{
                    cell = rowGroupSix.createCell(ArrayCoupleFour.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleFour.get(i).GetIDDay(), 10000);
                    rowGroupSix.setHeightInPoints(90);
                    break;
                }
                case(7):{
                    cell = rowGroupSeven.createCell(ArrayCoupleFour.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleFour.get(i).GetIDDay(), 10000);
                    rowGroupSeven.setHeightInPoints(90);
                    break;
                }
            }
        }

        for(int i = 0; i < ArrayCoupleFour.size(); i++) {

            Cell cell;

            String CoupleCell = "";

            if (i != ArrayCoupleFour.size()-1) {
                if (ArrayCoupleFour.get(i).GetCoupleNumber() == ArrayCoupleFour.get(i + 1).GetCoupleNumber()) {
                    CoupleCell = ArrayCoupleFour.get(i).GetDiscipline() + " (" + ArrayCoupleFour.get(i).GetCoupleType() + ")\n" + ArrayCoupleFour.get(i).GetNumberWeek() + " " + ArrayCoupleFour.get(i).GetTeacherName() + " " + ArrayCoupleFour.get(i).GetAud() + "\n" + "\n" + ArrayCoupleFour.get(i + 1).GetDiscipline() + " (" + ArrayCoupleFour.get(i + 1).GetCoupleType() + ")\n" + ArrayCoupleFour.get(i + 1).GetNumberWeek() + " " + ArrayCoupleFour.get(i + 1).GetTeacherName() + " " + ArrayCoupleFour.get(i + 1).GetAud();

                    switch (ArrayCoupleFour.get(i).GetCoupleNumber()){
                        case(1):{
                            cell = rowGroupOne.createCell(ArrayCoupleFour.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleFour.get(i).GetIDDay(), 10000);
                            rowGroupOne.setHeightInPoints(90);
                            break;
                        }
                        case(2):{
                            cell = rowGroupTwo.createCell(ArrayCoupleFour.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleFour.get(i).GetIDDay(), 10000);
                            rowGroupTwo.setHeightInPoints(90);
                            break;
                        }
                        case(3):{
                            cell = rowGroupThree.createCell(ArrayCoupleFour.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleFour.get(i).GetIDDay(), 10000);
                            rowGroupThree.setHeightInPoints(90);
                            break;
                        }
                        case(4):{
                            cell = rowGroupFour.createCell(ArrayCoupleFour.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleFour.get(i).GetIDDay(), 10000);
                            rowGroupFour.setHeightInPoints(90);
                            break;
                        }
                        case(5):{
                            cell = rowGroupFive.createCell(ArrayCoupleFour.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleFour.get(i).GetIDDay(), 10000);
                            rowGroupFive.setHeightInPoints(90);
                            break;
                        }
                        case(6):{
                            cell = rowGroupSix.createCell(ArrayCoupleFour.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleFour.get(i).GetIDDay(), 10000);
                            rowGroupSix.setHeightInPoints(90);
                            break;
                        }
                        case(7):{
                            cell = rowGroupSeven.createCell(ArrayCoupleFour.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleFour.get(i).GetIDDay(), 10000);
                            rowGroupSeven.setHeightInPoints(90);
                            break;
                        }
                    }
                }
            }
        }

        // CellFive

        for(int i = 0; i < ArrayCoupleFive.size(); i++) {

            Cell cell;

            String CoupleCell = "";

            CoupleCell = ArrayCoupleFive.get(i).GetDiscipline() + " (" + ArrayCoupleFive.get(i).GetCoupleType() + ")\n" + ArrayCoupleFive.get(i).GetNumberWeek() + " " + ArrayCoupleFive.get(i).GetTeacherName() + " " + ArrayCoupleFive.get(i).GetAud();

            switch (ArrayCoupleFive.get(i).GetCoupleNumber()){
                case(1):{
                    cell = rowGroupOne.createCell(ArrayCoupleFive.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleFive.get(i).GetIDDay(), 10000);
                    rowGroupOne.setHeightInPoints(90);
                    break;
                }
                case(2):{
                    cell = rowGroupTwo.createCell(ArrayCoupleFive.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleFive.get(i).GetIDDay(), 10000);
                    rowGroupTwo.setHeightInPoints(90);
                    break;
                }
                case(3):{
                    cell = rowGroupThree.createCell(ArrayCoupleFive.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleFive.get(i).GetIDDay(), 10000);
                    rowGroupThree.setHeightInPoints(90);
                    break;
                }
                case(4):{
                    cell = rowGroupFour.createCell(ArrayCoupleFive.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleFive.get(i).GetIDDay(), 10000);
                    rowGroupFour.setHeightInPoints(90);
                    break;
                }
                case(5):{
                    cell = rowGroupFive.createCell(ArrayCoupleFive.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleFive.get(i).GetIDDay(), 10000);
                    rowGroupFive.setHeightInPoints(90);
                    break;
                }
                case(6):{
                    cell = rowGroupSix.createCell(ArrayCoupleFive.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleFive.get(i).GetIDDay(), 10000);
                    rowGroupSix.setHeightInPoints(90);
                    break;
                }
                case(7):{
                    cell = rowGroupSeven.createCell(ArrayCoupleFive.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleFive.get(i).GetIDDay(), 10000);
                    rowGroupSeven.setHeightInPoints(90);
                    break;
                }
            }
        }

        for(int i = 0; i < ArrayCoupleFive.size(); i++) {

            Cell cell;

            String CoupleCell = "";

            if (i != ArrayCoupleFive.size() - 1) {
                if (ArrayCoupleFive.get(i).GetCoupleNumber() == ArrayCoupleFive.get(i + 1).GetCoupleNumber()) {
                    CoupleCell = ArrayCoupleFive.get(i).GetDiscipline() + " (" + ArrayCoupleFive.get(i).GetCoupleType() + ")\n" + ArrayCoupleFive.get(i).GetNumberWeek() + " " + ArrayCoupleFive.get(i).GetTeacherName() + " " + ArrayCoupleFive.get(i).GetAud() + "\n" + "\n" + ArrayCoupleFive.get(i + 1).GetDiscipline() + " (" + ArrayCoupleFive.get(i + 1).GetCoupleType() + ")\n" + ArrayCoupleFive.get(i + 1).GetNumberWeek() + " " + ArrayCoupleFive.get(i + 1).GetTeacherName() + " " + ArrayCoupleFive.get(i + 1).GetAud();

                    switch (ArrayCoupleFive.get(i).GetCoupleNumber()) {
                        case (1): {
                            cell = rowGroupOne.createCell(ArrayCoupleFive.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleFive.get(i).GetIDDay(), 10000);
                            rowGroupOne.setHeightInPoints(90);
                            break;
                        }
                        case (2): {
                            cell = rowGroupTwo.createCell(ArrayCoupleFive.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleFive.get(i).GetIDDay(), 10000);
                            rowGroupTwo.setHeightInPoints(90);
                            break;
                        }
                        case (3): {
                            cell = rowGroupThree.createCell(ArrayCoupleFive.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleFive.get(i).GetIDDay(), 10000);
                            rowGroupThree.setHeightInPoints(90);
                            break;
                        }
                        case (4): {
                            cell = rowGroupFour.createCell(ArrayCoupleFive.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleFive.get(i).GetIDDay(), 10000);
                            rowGroupFour.setHeightInPoints(90);
                            break;
                        }
                        case (5): {
                            cell = rowGroupFive.createCell(ArrayCoupleFive.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleFive.get(i).GetIDDay(), 10000);
                            rowGroupFive.setHeightInPoints(90);
                            break;
                        }
                        case (6): {
                            cell = rowGroupSix.createCell(ArrayCoupleFive.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleFive.get(i).GetIDDay(), 10000);
                            rowGroupSix.setHeightInPoints(90);
                            break;
                        }
                        case (7): {
                            cell = rowGroupSeven.createCell(ArrayCoupleFive.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleFive.get(i).GetIDDay(), 10000);
                            rowGroupSeven.setHeightInPoints(90);
                            break;
                        }
                    }
                }
            }
        }

        // CellSix

        for(int i = 0; i < ArrayCoupleSix.size(); i++) {

            Cell cell;

            String CoupleCell = "";

            CoupleCell = ArrayCoupleSix.get(i).GetDiscipline() + " (" + ArrayCoupleSix.get(i).GetCoupleType() + ")\n" + ArrayCoupleSix.get(i).GetNumberWeek() + " " + ArrayCoupleSix.get(i).GetTeacherName() + " " + ArrayCoupleSix.get(i).GetAud();

            switch (ArrayCoupleSix.get(i).GetCoupleNumber()){
                case(1):{
                    cell = rowGroupOne.createCell(ArrayCoupleSix.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleSix.get(i).GetIDDay(), 10000);
                    rowGroupOne.setHeightInPoints(90);
                    break;
                }
                case(2):{
                    cell = rowGroupTwo.createCell(ArrayCoupleSix.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleSix.get(i).GetIDDay(), 10000);
                    rowGroupTwo.setHeightInPoints(90);
                    break;
                }
                case(3):{
                    cell = rowGroupThree.createCell(ArrayCoupleSix.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleSix.get(i).GetIDDay(), 10000);
                    rowGroupThree.setHeightInPoints(90);
                    break;
                }
                case(4):{
                    cell = rowGroupFour.createCell(ArrayCoupleSix.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleSix.get(i).GetIDDay(), 10000);
                    rowGroupFour.setHeightInPoints(90);
                    break;
                }
                case(5):{
                    cell = rowGroupFive.createCell(ArrayCoupleSix.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleSix.get(i).GetIDDay(), 10000);
                    rowGroupFive.setHeightInPoints(90);
                    break;
                }
                case(6):{
                    cell = rowGroupSix.createCell(ArrayCoupleSix.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleSix.get(i).GetIDDay(), 10000);
                    rowGroupSix.setHeightInPoints(90);
                    break;
                }
                case(7):{
                    cell = rowGroupSeven.createCell(ArrayCoupleSix.get(i).GetIDDay());
                    cell.setCellValue(CoupleCell);
                    cell.setCellStyle(cellStyle);
                    Group.setColumnWidth(ArrayCoupleSix.get(i).GetIDDay(), 10000);
                    rowGroupSeven.setHeightInPoints(90);
                    break;
                }
            }
        }

        for(int i = 0; i < ArrayCoupleSix.size(); i++) {

            Cell cell;

            String CoupleCell = "";

            if (i != ArrayCoupleSix.size() - 1) {
                if (ArrayCoupleSix.get(i).GetCoupleNumber() == ArrayCoupleSix.get(i + 1).GetCoupleNumber()) {
                    CoupleCell = ArrayCoupleSix.get(i).GetDiscipline() + " (" + ArrayCoupleSix.get(i).GetCoupleType() + ")\n" + ArrayCoupleSix.get(i).GetNumberWeek() + " " + ArrayCoupleSix.get(i).GetTeacherName() + " " + ArrayCoupleSix.get(i).GetAud() + "\n" + "\n" + ArrayCoupleSix.get(i + 1).GetDiscipline() + " (" + ArrayCoupleSix.get(i + 1).GetCoupleType() + ")\n" + ArrayCoupleSix.get(i + 1).GetNumberWeek() + " " + ArrayCoupleSix.get(i + 1).GetTeacherName() + " " + ArrayCoupleSix.get(i + 1).GetAud();

                    switch (ArrayCoupleSix.get(i).GetCoupleNumber()) {
                        case (1): {
                            cell = rowGroupOne.createCell(ArrayCoupleSix.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleSix.get(i).GetIDDay(), 10000);
                            rowGroupOne.setHeightInPoints(90);
                            break;
                        }
                        case (2): {
                            cell = rowGroupTwo.createCell(ArrayCoupleSix.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleSix.get(i).GetIDDay(), 10000);
                            rowGroupTwo.setHeightInPoints(90);
                            break;
                        }
                        case (3): {
                            cell = rowGroupThree.createCell(ArrayCoupleSix.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleSix.get(i).GetIDDay(), 10000);
                            rowGroupThree.setHeightInPoints(90);
                            break;
                        }
                        case (4): {
                            cell = rowGroupFour.createCell(ArrayCoupleSix.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleSix.get(i).GetIDDay(), 10000);
                            rowGroupFour.setHeightInPoints(90);
                            break;
                        }
                        case (5): {
                            cell = rowGroupFive.createCell(ArrayCoupleSix.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleSix.get(i).GetIDDay(), 10000);
                            rowGroupFive.setHeightInPoints(90);
                            break;
                        }
                        case (6): {
                            cell = rowGroupSix.createCell(ArrayCoupleSix.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleSix.get(i).GetIDDay(), 10000);
                            rowGroupSix.setHeightInPoints(90);
                            break;
                        }
                        case (7): {
                            cell = rowGroupSeven.createCell(ArrayCoupleSix.get(i).GetIDDay());
                            cell.setCellValue(CoupleCell);
                            cell.setCellStyle(cellStyle);
                            Group.setColumnWidth(ArrayCoupleSix.get(i).GetIDDay(), 10000);
                            rowGroupSeven.setHeightInPoints(90);
                            break;
                        }
                    }
                }
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream("OneGroupExelDoc");

        workbookOneElementGroup.write(fileOutputStream);
        fileOutputStream.close();
    }

}
