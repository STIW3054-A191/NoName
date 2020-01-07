package groupasg;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class SubMain extends Thread
{
    private final String githubLink = "https://github.com/STIW3054-A191/Assignments/issues/1";
    private final String myString = "STIW3054-A191-A";
    private Document doc;
    private PatternLink list;
    private int total=1;
    PatternLink pl = new PatternLink( );
    ListLink ll = new ListLink();
    private int start;

    synchronized public void run()
    {
            try {
                doc = Jsoup.connect(githubLink).get( );
                Elements links = doc.select("a[href]");
                for (Element link : links) {
                    String theLink = link.attr("href");
                    String[] arr = theLink.split("/");
                    String matric = "";
                    if (pl.isMatch(myString, theLink) == true) {
                        for ( String ss : arr) {
                            if(ss.startsWith("2")){
                                matric = ss.substring(0,6);

                            }}
                        RepoDw repoDw = new RepoDw(theLink);
                        repoDw.RepoDownload(theLink);
                        total = repoDw.getCount();
                        // System.out.println(total+" "+ matric +" " +theLink);
                        //groupasg.ListLink l = new groupasg.ListLink(theLink, matric);
                        ll.keepLink(matric, theLink);
                        //total++;
                    }
                }
                Thread.sleep(1000);
                //ll.StudentLink();
               //System.out.println(total);

            } catch (Exception e) {
                e.printStackTrace( );
            }


    }

}
