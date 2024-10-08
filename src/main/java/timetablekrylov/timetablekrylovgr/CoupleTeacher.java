package timetablekrylov.timetablekrylovgr;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

// Creating a CoupleTeacher class
public class CoupleTeacher {

    // I am creating private fields of the CoupleTeacher class
    // The field that stores the ID of the day
    private int idDay;
    // The field that stores the number of the pair
    private int NumberCouple;
    // The field that stores discipline
    private String Discipline;
    // The field that stores the type of the couple
    private String Type;
    // The field that stores the type of the week
    private String TypeWeek;
    // The field that stores auditorium
    private String Aud;
    private String NumberWeek;
    private String UnderGroup;
    private String GroupName;

    // Creating a method that creates couple
    // To do this, use the Json string
    public void CreatorCouple(String Json){
        SetIdDay(Json);
        SetCoupleNumber(Json);
        SetDiscipline(Json);
        SetType(Json);
        SetTypeWeek(Json);
        SetAud(Json);
        SetNumberWeek(Json);
        SetUnderGroup(Json);
        SetGroupName(Json);
    }

    // Method for changing the idDay value
    private void SetIdDay(String Json){

        // I use regular expressions
        Pattern pattern = Pattern.compile("\"id_day\":\"(\\d|)\"");
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
        // Assigning values to the idDay field
        while(matcher1.find()){
            idDay = Integer.parseInt(matcher1.group());
        }

    }

    // Method for changing the CoupleNumber value
    private void SetCoupleNumber(String Json){

        // I use regular expressions
        Pattern pattern = Pattern.compile("\"number_para\":\"(\\d|)\"");
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
        // Assigning values to the NumberCouple field
        while(matcher1.find()){
            NumberCouple = Integer.parseInt(matcher1.group());
        }

    }

    // Method for changing the Discipline value
    private void SetDiscipline(String Json){

        // I use regular expressions
        Pattern pattern = Pattern.compile("\"discipline\":\"([A-zА-яё0-9 \\~\\`\\!\\@\\#\\№\\$\\;\\%\\^\\:\\&\\?\\*\\(\\)\\-\\_\\+\\=\\.\\,\\}\\{\\[\\]\\|\\/]+|)\"");
        Matcher matcher = pattern.matcher(Json);

        String RegX = "";

        // Getting the value
        while(matcher.find()){
            RegX = matcher.group();
        }

        // I use regular expressions
        Pattern pattern1 = Pattern.compile("\"([A-zА-яё0-9 \\~\\`\\!\\@\\#\\№\\$\\;\\%\\^\\\\&\\?\\*\\(\\)\\-\\+\\=\\.\\,\\}\\{\\[\\]\\|\\/]+|)\"");
        Matcher matcher1 = pattern1.matcher(RegX);

        // Getting the value
        // Assigning values to the Discipline field
        while(matcher1.find()){
            Discipline = matcher1.group();
        }

    }

    // Method for changing the Type value
    private void SetType(String Json){

        // I use regular expressions
        Pattern pattern = Pattern.compile("\"type\":\"([А-яA-zё0-9 \\~\\`\\!\\@\\#\\№\\$\\;\\%\\^\\:\\&\\?\\*\\(\\)\\-\\_\\+\\=\\.\\,\\}\\{\\[\\]\\|\\/]+|)\"");
        Matcher matcher = pattern.matcher(Json);

        String RegX = "";

        // Getting the value
        while(matcher.find()){
            RegX = matcher.group();
        }

        // I use regular expressions
        Pattern pattern1 = Pattern.compile("\"([A-zА-яё0-9 \\~\\`\\!\\@\\#\\№\\$\\;\\%\\^\\\\&\\?\\*\\(\\)\\-\\+\\=\\.\\,\\}\\{\\[\\]\\|\\/]+|)\"");
        Matcher matcher1 = pattern1.matcher(RegX);

        // Getting the value
        // Assigning values to the Type field
        while(matcher1.find()){
            Type = matcher1.group();
        }

    }

    // Method for changing the TypeWeek value
    private void SetTypeWeek(String Json){

        // I use regular expressions
        Pattern pattern = Pattern.compile("\"type_week\":\"([А-яA-zё0-9 \\~\\`\\!\\@\\#\\№\\$\\;\\%\\^\\:\\&\\?\\*\\(\\)\\-\\_\\+\\=\\.\\,\\}\\{\\[\\]\\|\\/]+|)\"");
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
        Pattern pattern = Pattern.compile("\"aud\":\"([А-яA-zё0-9 \\~\\`\\!\\@\\#\\№\\$\\;\\%\\^\\:\\&\\?\\*\\(\\)\\-\\_\\+\\=\\.\\,\\}\\{\\[\\]\\|\\/]+|)\"");
        Matcher matcher = pattern.matcher(Json);

        String RegX = "";

        // Getting the value
        while(matcher.find()){
            RegX = matcher.group();
        }

        // I use regular expressions
        Pattern pattern1 = Pattern.compile("\"([A-zА-яё0-9 \\~\\`\\!\\@\\#\\№\\$\\;\\%\\^\\\\&\\?\\*\\(\\)\\-\\+\\=\\.\\,\\}\\{\\[\\]\\|\\/]+|)\"");
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
        Pattern pattern = Pattern.compile("\"number_week\":\"([0-9 \\~\\`\\!\\@\\#\\№\\$\\;\\%\\^\\:\\&\\?\\*\\(\\)\\-\\_\\+\\=\\.\\,\\}\\{\\[\\]\\|\\/]|)+\"");
        Matcher matcher = pattern.matcher(Json);

        String RegX = "";

        // Getting the value
        while(matcher.find()){
            RegX = matcher.group();
        }

        // I use regular expressions
        Pattern pattern1 = Pattern.compile("\"([A-zА-яё0-9 \\~\\`\\!\\@\\#\\№\\$\\;\\%\\^\\\\&\\?\\*\\(\\)\\-\\+\\=\\.\\,\\}\\{\\[\\]\\|\\/]+|)\"");
        Matcher matcher1 = pattern1.matcher(RegX);

        // Getting the value
        // Assigning values to the NumberWeek field
        while(matcher1.find()){
            NumberWeek = matcher1.group();
        }

    }

    // Method for changing the UnderGroup value
    private void SetUnderGroup(String Json){

        String RegX = "";

        // I use regular expressions
        Pattern pattern = Pattern.compile("\"under_group\":\"([A-zА-яё0-9 \\~\\`\\!\\@\\#\\№\\$\\;\\%\\^\\:\\&\\?\\*\\(\\)\\-\\_\\+\\=\\.\\,\\}\\{\\[\\]\\|\\/]+|)\"");
        Matcher matcher = pattern.matcher(Json);

        // Getting the value
        while(matcher.find()){
            RegX = matcher.group();
        }

        // I use regular expressions
        Pattern pattern1 = Pattern.compile("\"([A-zА-яё0-9 \\~\\`\\!\\@\\#\\№\\$\\;\\%\\^\\\\&\\?\\*\\(\\)\\-\\+\\=\\.\\,\\}\\{\\[\\]\\|\\/]+|)\"");
        Matcher matcher1 = pattern1.matcher(RegX);

        // Getting the value
        // Assigning values to the UnderGroup field
        while(matcher1.find()){
            UnderGroup = matcher1.group();
        }

    }

    // Method for changing the GroupName value
    private void SetGroupName(String Json){

        String RegX = "";

        // I use regular expressions
        Pattern pattern = Pattern.compile("\"group_name\":\"([A-zА-яё0-9 \\~\\`\\!\\@\\#\\№\\$\\;\\%\\^\\:\\&\\?\\*\\(\\)\\-\\_\\+\\=\\.\\,\\}\\{\\[\\]\\|\\/]+|)\"");
        Matcher matcher = pattern.matcher(Json);

        // Getting the value
        while(matcher.find()){
            RegX = matcher.group();
        }

        // I use regular expressions
        Pattern pattern1 = Pattern.compile("\"([A-zА-яё0-9 \\~\\`\\!\\@\\#\\№\\$\\;\\%\\^\\:\\&\\?\\*\\(\\)\\-\\_\\+\\=\\.\\,\\}\\{\\[\\]\\|\\/]+|)\"");
        Matcher matcher1 = pattern1.matcher(RegX);

        // Getting the value
        // Assigning values to the GroupName field
        while (matcher1.find()) {
            GroupName = matcher1.group();
        }

    }

    public int GetIdDay(){
        return idDay;
    }

    public int GetNumberCouple(){
        return NumberCouple;
    }

    public String GetDiscipline(){
        return Discipline;
    }

    public String GetType(){
        return Type;
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

    public String GetUnderGroup(){
        return UnderGroup;
    }

    public String GetGroupName(){
        return GroupName;
    }

}
