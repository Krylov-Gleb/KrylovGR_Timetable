package timetablekrylov.timetablekrylovgr;

import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Group {

    // Creating a field that stores the group id
    private int groupId;

    // Creating a field that stores the name of the group
    private String groupName;

    // Creating an array of pairs in Json format
    private ArrayList<String> ArrayCouplesJson = new ArrayList<>();

    // Creating an array of objects of the couple class
    private ArrayList<CoupleGroup> ArrayCouples = new ArrayList<>();

    private ArrayList<String> ArrayExelCouple = new ArrayList<>();

    // I create a method to set the group ID
    public void SetGroupId(String Json){

        // I use regular expressions
        // (More details in class Teacher)
        Pattern pattern = Pattern.compile("\"id\":\\d+,\"name\":");
        Matcher matcher = pattern.matcher(Json);

        String RegX = "";

        // Getting the value
        while(matcher.find()){
            RegX = matcher.group();
        }

        // I use regular expressions
        Pattern pattern1 = Pattern.compile("\\d+");
        Matcher matcher1 = pattern1.matcher(RegX);

        // Getting the value
        while(matcher1.find()){
            groupId = Integer.parseInt(matcher1.group());
        }

    }

    // I create a method to set the group Name
    public void SetGroupName(String Json){

        // I use regular expressions
        // (More details in class Teacher)
        Pattern pattern = Pattern.compile("id\":\\d+,\"name\":\"[А-яA-z-\\dё]+\"");
        Matcher matcher = pattern.matcher(Json);

        String RegX = "";

        // Getting the value
        while(matcher.find()){
            RegX = matcher.group();
        }

        // I use regular expressions
        Pattern pattern1 = Pattern.compile("[A-zА-яё]+-\\d+");
        Matcher matcher1 = pattern1.matcher(RegX);

        // Getting the value
        while(matcher1.find()){
            groupName = matcher1.group();
        }

    }

    // I am creating a method to create group couple
    public void CreatorCouples(String Json){

        // I use regular expressions
        // (More details in class Teacher)
        Pattern pattern = Pattern.compile("\"id_day\":\"\\d\",\"number_para\":\"\\d\",\"discipline\":\"[A-zА-я \\.\\-\\/]+\",\"type\":\"[А-яA-z]+\",\"type_week\":\"[А-яA-z]+\",\"aud\":\"[А-я\\. 0-9\\/\\-]+\",\"number_week\":\"[0-9\\/\\,\\-]+\",\"comment\":\"(|[A-zА-я\\.\\/\\,\\-])\",\"zaoch\":(true|false),\"name\":\"[A-zА-яё. ]+\"");
        Matcher matcher = pattern.matcher(Json);

        while(matcher.find()){
            ArrayCouplesJson.add(matcher.group());
        }

        for(int i = 0; i < ArrayCouplesJson.size(); i++){
            ArrayCouples.add(new CoupleGroup());
        }

        for(int i = 0; i < ArrayCouples.size(); i++){
            ArrayCouples.get(i).CreatorCouple(ArrayCouplesJson.get(i));
        }

        SetGroupId(Json);
        SetGroupName(Json);

    }

    // I create a method to output data
    public void GetCouples(){
        System.out.print("\n");
        System.out.println("Id группы = " + groupId);
        System.out.println("Группа = " + groupName);
        System.out.print("\n");
        for(int i = 0; i < ArrayCouples.size(); i++){
            ArrayCouples.get(i).GetCouple();
        }
    }

    public ArrayList<CoupleGroup> GetArrayCouples(){
        return ArrayCouples;
    }

    public ArrayList<String> CreatorExelCouple(){

        String Couple = "";

        for(int i = 0; i < ArrayCouples.size(); i++){
            Couple = String.valueOf(ArrayCouples.get(i).GetIDDay()) + ArrayCouples.get(i).GetCoupleNumber() + ArrayCouples.get(i).GetCoupleType() + ArrayCouples.get(i).GetDiscipline() + ArrayCouples.get(i).GetTypeWeek() + ArrayCouples.get(i).GetAud() + ArrayCouples.get(i).GetNumberWeek() + ArrayCouples.get(i).GetZaoch() + ArrayCouples.get(i).GetTeacherName();
            ArrayExelCouple.add(Couple);
            Couple = "";
        }

        return ArrayExelCouple;
    }
}
