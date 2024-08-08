package timetablekrylov.timetablekrylovgr;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

// Creating a Couple class
public class CoupleGroup {

    // I am creating private fields of the Couple class
    // The field that stores the ID of the day
    private int IDDay;
    // The field that stores the number of the pair
    private int CoupleNumber;
    // The field that stores the type of the couple
    private String CoupleType;
    // The field that stores discipline
    private String Discipline;
    private String TypeWeek;
    private String Aud;
    private String NumberWeek;
    private boolean Zaoch;
    private String TeacherName;

    // Creating a method that creates couple
    // To do this, use the Json string
    public void CreatorCouple(String Json){
        SetIdDay(Json);
        SetCoupleNumber(Json);
        SetTypeCouple(Json);
        SetDiscipline(Json);
        SetTypeWeek(Json);
        SetAud(Json);
        SetNumberWeek(Json);
        SetZaoch(Json);
        SetTeacherName(Json);
    }

    // Method for output couple
    public void GetCouple(){
        System.out.println("День недели = " + IDDay);
        System.out.println("Номер пары = " + CoupleNumber);
        System.out.println("Дисциплина = " + Discipline);
        System.out.println("Тип занятия = " + CoupleType);
        System.out.println("Тип недели = " + TypeWeek);
        System.out.println("Номер аудитории = " + Aud);
        System.out.println("Недели = " + NumberWeek);
        System.out.println("Заочная форма = " + Zaoch);
        System.out.println("Преподаватель = " + TeacherName);
        System.out.println("\n");
    }

    // Method for changing the IdDay value
    private void SetIdDay(String Json){

        // I use regular expressions
        Pattern pattern = Pattern.compile("\"id_day\":\"\\d\"");
        Matcher matcher = pattern.matcher(Json);

        String RegX = "";

        // Getting the value
        while(matcher.find()){
            RegX = matcher.group();
        }

        // I use regular expressions
        Pattern pattern1 = Pattern.compile("\\d");
        Matcher matcher1 = pattern1.matcher(RegX);

        // Getting the value
        // Assigning values to the IDDay field
        while(matcher1.find()){
            IDDay = Integer.parseInt(matcher1.group());
        }

    }

    // Method for changing the CoupleNumber value
    private void SetCoupleNumber(String Json){

        // I use regular expressions
        Pattern pattern = Pattern.compile("\"number_para\":\"\\d\"");
        Matcher matcher = pattern.matcher(Json);

        String RegX = "";

        // Getting the value
        while(matcher.find()){
            RegX = matcher.group();
        }

        // I use regular expressions
        Pattern pattern1 = Pattern.compile("\\d");
        Matcher matcher1 = pattern1.matcher(RegX);

        // Getting the value
        // Assigning values to the CoupleNumber field
        while(matcher1.find()){
            CoupleNumber = Integer.parseInt(matcher1.group());
        }

    }

    // Method for changing the Discipline value
    private void SetDiscipline(String Json){

        // I use regular expressions
        Pattern pattern = Pattern.compile("\"discipline\":\"[A-zА-яё\\.\\-\\, ]+\"");
        Matcher matcher = pattern.matcher(Json);

        String RegX = "";

        // Getting the value
        while(matcher.find()){
            RegX = matcher.group();
        }

        // I use regular expressions
        Pattern pattern1 = Pattern.compile("[А-я\\-\\.\\, ]+");
        Matcher matcher1 = pattern1.matcher(RegX);

        // Getting the value
        // // Assigning values to the Discipline field
        while(matcher1.find()){
            Discipline = matcher1.group();
        }

    }

    // Method for changing the TypeCouple value
    private void SetTypeCouple(String Json){

        // I use regular expressions
        Pattern pattern = Pattern.compile("\"type\":\"[A-zА-яё]+\"");
        Matcher matcher = pattern.matcher(Json);

        String RegX = "";

        // Getting the value
        while(matcher.find()){
            RegX = matcher.group();
        }

        // I use regular expressions
        Pattern pattern1 = Pattern.compile("[А-яё]+");
        Matcher matcher1 = pattern1.matcher(RegX);

        // Getting the value
        // Assigning values to the CoupleType field
        while(matcher1.find()){
            CoupleType = matcher1.group();
        }

    }

    // Method for changing the TypeWeek value
    private void SetTypeWeek(String Json){

        // I use regular expressions
        Pattern pattern = Pattern.compile("\"type_week\":\"[A-zА-яё]+\"");
        Matcher matcher = pattern.matcher(Json);

        String RegX = "";

        // Getting the value
        while(matcher.find()){
            RegX = matcher.group();
        }

        // I use regular expressions
        Pattern pattern1 = Pattern.compile("(all|even|odd)");
        Matcher matcher1 = pattern1.matcher(RegX);

        // Getting the value
        // Assigning values to the TypeWeek field
        while(matcher1.find()){
            TypeWeek = matcher1.group();
        }

    }

    // Method for changing the Aud value
    private void SetAud(String Json){

        // I use regular expressions
        Pattern pattern = Pattern.compile("\"aud\":\"[А-я\\. 0-9\\/\\-]+\"");
        Matcher matcher = pattern.matcher(Json);

        String RegX = "";

        // Getting the value
        while(matcher.find()){
            RegX = matcher.group();
        }

        // I use regular expressions
        Pattern pattern1 = Pattern.compile("([\\d]{1,3}(|\\-)([А-яA-z0-9]|)\\/\\d|Зал)");
        Matcher matcher1 = pattern1.matcher(RegX);

        // Getting the value
        // Assigning values to the Aud field
        while(matcher1.find()){
            Aud = matcher1.group();
        }

    }

    // Method for changing the NumberWeek value
    private void SetNumberWeek(String Json){

        // I use regular expressions
        Pattern pattern = Pattern.compile("\"number_week\":\"[\\d\\-\\/\\,]+\"");
        Matcher matcher = pattern.matcher(Json);

        String RegX = "";

        // Getting the value
        while(matcher.find()){
            RegX = matcher.group();
        }

        // I use regular expressions
        Pattern pattern1 = Pattern.compile("[\\d\\-\\/\\,]+");
        Matcher matcher1 = pattern1.matcher(RegX);

        // Getting the value
        // Assigning values to the NumberWeek field
        while(matcher1.find()){
            NumberWeek = matcher1.group();
        }

    }

    // Method for changing the Zaoch value
    private void SetZaoch(String Json){

        // I use regular expressions
        Pattern pattern = Pattern.compile("\"zaoch\":(true|false)");
        Matcher matcher = pattern.matcher(Json);

        String RegX = "";

        // Getting the value
        while(matcher.find()){
            RegX = matcher.group();
        }

        // I use regular expressions
        Pattern pattern1 = Pattern.compile("(true|false)");
        Matcher matcher1 = pattern1.matcher(RegX);

        // Getting the value
        // Assigning values to the Zaoch field
        while(matcher1.find()){
            Zaoch = Boolean.parseBoolean(matcher1.group());
        }

    }

    // Method for changing the TeacherName value
    private void SetTeacherName(String Json){

        // I use regular expressions
        Pattern pattern = Pattern.compile("\"name\":\"[A-zА-яё. ]+\"");
        Matcher matcher = pattern.matcher(Json);

        String RegX = "";

        // Getting the value
        while(matcher.find()){
            RegX = matcher.group();
        }

        // I use regular expressions
        Pattern pattern1 = Pattern.compile("[А-яё. ]+");
        Matcher matcher1 = pattern1.matcher(RegX);

        // Getting the value
        // Assigning values to the TeacherName field
        while(matcher1.find()){
            TeacherName = matcher1.group();
        }

    }

    public int GetIDDay(){
        return IDDay;
    }

    public int GetCoupleNumber(){
        return CoupleNumber;
    }

    public String GetCoupleType(){
        return CoupleType;
    }

    public String GetDiscipline(){
        return Discipline;
    }

    public String GetTypeWeek(){
        return TypeWeek;
    }

    public String GetAud(){
        return Aud;
    }

    public String GetNumberWeek(){
        return NumberWeek;
    }

    public boolean GetZaoch(){
        return Zaoch;
    }

    public String GetTeacherName(){
        return TeacherName;
    }
}
