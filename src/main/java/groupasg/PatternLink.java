package groupasg;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class PatternLink {

    public static boolean isMatch(String pattern, String link) {

        Pattern p = Pattern.compile(pattern);   // the pattern to search for
        Matcher m = p.matcher(link);

        if (m.find())
            return true;
        else
            return false;
    }

    public static void main(String[] args) {
        String githubLink = "https://github.com/STIW3054-A191/Assignments/issues/1";
        String myString = "STIW3054-A191";
        Boolean result = new PatternLink().isMatch(myString, githubLink);
        System.out.println(result);
    }
}