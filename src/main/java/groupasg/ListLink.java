package groupasg;

import java.util.ArrayList;

public class ListLink {
    String link;
    String matric;
    int count = 1;
    ArrayList<ListLink> list = new ArrayList<ListLink>( );


    ListLink(String matric, String link) {
        this.matric = matric;
        this.link = link;
    }

    ListLink() {
    }

    public String getLink() {
        return link;
    }

    public void setLink(String link) {
        this.link = link;
    }

    public String getMatric() {
        return matric;
    }

    public void setMatric(String matric) {
        this.matric = matric;
    }

    public void keepLink(String matric, String link){
        ListLink s = new ListLink(matric, link);
        list.add(s);
    }
    public void StudentLink() {
        for (ListLink std : list) {
            System.out.println(std.getMatric( ) + " " + std.getLink( ));
        }
    }
}
