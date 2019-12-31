import org.eclipse.jgit.api.Git;
import org.eclipse.jgit.api.errors.GitAPIException;

import java.nio.file.Paths;

public class RepoDw {
    String link;
    RepoDw(String link){
        this.link=link;
    }

    public String getLink() {
        return link;
    }

    public void setLink(String link) {
        this.link = link;
    }

    void RepoDownload(String link){
       // String repoUrl = "https://github.com/SilentHlive/255227-STIW3054-A191-A2.git";
        String foldern = link.replace(".git","");
        String[] arr = foldern.split("/");
        String fname = "";
        for ( String ss : arr) {
            if(ss.startsWith("2")){
                fname = ss;
            }}
        System.out.println(fname);
        String cloneDirectoryPath = "C:\\Users\\HALIMAH\\IdeaProjects\\GroupAsg\\src\\repo\\"+fname; // Ex.in windows c:\\gitProjects\SpringBootMongoDbCRUD\
        try {
            System.out.println("Cloning "+link+" into "+link);
            Git.cloneRepository().setURI(link).setDirectory(Paths.get(cloneDirectoryPath).toFile()).call();
            System.out.println("Completed Cloning");
        } catch (GitAPIException e) {
            System.out.println("Exception occurred while cloning repo");
            e.printStackTrace();
        }
    }
}
