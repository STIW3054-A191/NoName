package groupasg;

import java.io.IOException;
import java.util.ArrayList;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class ListStudents {
    public static String webPage = "https://github.com/STIW3054-A191/Assignments/wiki/List_of_Student";
    public static void main(String[] args) throws IOException {
        String n = "", m = "";
        Student student = new Student( );
        ArrayList<Student> list = new ArrayList<Student>( );
        int count = 0;
        Element html = Jsoup.connect(webPage).get( ).getElementById("wiki-body");
        Elements e = html.select("table");
        Elements body = e.select("tbody");
        Elements rowItems = body.select("tr");
        int st = rowItems.size( );
        Elements td = rowItems.select("td");
        for (int j = 0; j < td.size( ); j++) {
            if (td.get(j).text( ).startsWith(String.valueOf(2)) && td.get(j).text( ).length( ) == 6) {
                // System.out.println(j +" " +td.get(j).text( ));
                m = td.get(j).text( );
                student.setMatricno(m);
            }
            if (td.get(j).text( ).toUpperCase( ).charAt(0) >= 'A' && td.get(j).text( ).toUpperCase( ).charAt(0) <= 'Z') {
                // System.out.println(td.get(j).text( ));
                n = td.get(j).text( );
                student.setName(n);
                student.createAccount(m, n);
            }
        }
        student.StudentList();
    }
}