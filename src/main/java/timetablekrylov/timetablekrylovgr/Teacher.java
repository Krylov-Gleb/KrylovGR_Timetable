package timetablekrylov.timetablekrylovgr;

import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

// Creating a teacher class
public class Teacher {

    // Creating a variable for the teacher's name
    private String TeacherName;

    // Creating an array of pairs in Json format
    private ArrayList<String> ArrayTeacherCouplesJson = new ArrayList<>();

    // Creating objects of the CoupleTeacher class
    private ArrayList<CoupleTeacher> ArrayTeacherCouples = new ArrayList<>();

    // I am creating a private method to change the parameter of the Name (Teacher) field
    private void SetNameTeacher(String Json){

        // I use regular expressions
        // Regular expression ("\"name\":\"[ёA-zА-я ]+\"")
        Pattern pattern = Pattern.compile("\"name\":\"[ёA-zА-я ]+\"");

        // I use an expression on a Json string
        Matcher matcher = pattern.matcher(Json);

        // Creating a variable to store the result
        String RegX = "";

        // I assign the resulting value
        while(matcher.find()){
            RegX = matcher.group();
        }

        // I use regular expressions
        // Getting rid of the excess
        // Regular expression ("\"[А-яё ]+\"")
        Pattern pattern1 = Pattern.compile("\"[А-яё ]+\"");
        // I use an expression on a RegX string
        Matcher matcher1 = pattern1.matcher(RegX);

        // I assign the resulting value to the TeacherName field
        while(matcher1.find()){
            TeacherName = matcher1.group();
        }

    }

    // A method for creating a schedule of pairs
    public void CreatorCouples(String Json){

        // Using a regular expression, I get teacher classes in Json format.
        Pattern pattern = Pattern.compile("\"id_day\":\"\\d\",\"number_para\":\"\\d\",\"discipline\":\"[A-zА-я \\.\\-\\/]+\",\"type\":\"[А-яA-z]+\",\"type_week\":\"[А-яA-z]+\",\"aud\":\"[А-я\\. 0-9\\/\\-]+\",\"number_week\":\"([0-9\\/\\,\\-]+|)\",\"comment\":\"(|[A-zА-я\\.\\/\\,\\-])\",\"zaoch\":(true|false),\"name\":\"([A-zА-яё. ]|)\",\"under_group\":\"(|([0-9п\\/г \\,]|)+)\",\"under_group_1\":\"(|([0-9п\\/г \\,]|)+)\",\"under_group_2\":\"(|([0-9п\\/г \\,]|)+)\",\"group_name\":\"(|[А-я\\-\\d\\, ]+)\"");

        // I use an expression on a Json string
        Matcher matcher = pattern.matcher(Json);

        // We get the pairs and write them to an array (Json format).
        while(matcher.find()){
            ArrayTeacherCouplesJson.add(matcher.group());
        }

        // Depending on the number of elements, we create CoupleTeacher objects and write them to an array.
        for(int i = 0; i < ArrayTeacherCouplesJson.size(); i++){
            ArrayTeacherCouples.add(new CoupleTeacher());
        }

        // Setting values for objects of the CoupleTeacher class
        for(int i = 0; i < ArrayTeacherCouples.size(); i++){
            ArrayTeacherCouples.get(i).CreatorCouple(ArrayTeacherCouplesJson.get(i));
        }

        // We set the name of the teacher
        SetNameTeacher(Json);

    }

    // The method for getting the teacher's schedule
    public void GetCoupleTeacher(){
        System.out.print("\n");

        // Output TeacherName
        System.out.println(TeacherName);
        System.out.println("\n");

        // Output values from the ArrayTeacherCouples array (Pairs)
        for(int i = 0; i < ArrayTeacherCouples.size(); i++){
            ArrayTeacherCouples.get(i).GetCouple();
        }
    }

    public ArrayList<CoupleTeacher> GetArrayCoupleTeacher(){
        return ArrayTeacherCouples;
    }


}
