package groupasg;

import java.io.File;
import java.util.concurrent.locks.ReentrantLock;
import java.io.*;
public class Compile extends Thread
{
    private static String filen;
    // create a file that is really a directory
    static final String filename= "pom.xml";
    private static ReentrantLock lock = new ReentrantLock();
    static Compile t1 = new Compile();


    public static void main(String[] args)
    {

        File aDirectory = new File("C:\\Users\\HALIMAH\\IdeaProjects\\GroupAsg\\src\\repo");
        // get a listing of all files in the directory
        String[] filesInDir = aDirectory.list();
        // sort the list of files (optional)
        // Arrays.sort(filesInDir);
        // have everything i need, just print it now
        for ( int i=0; i<filesInDir.length; i++ )
        {

            System.out.println( "file: " + filesInDir[i] );
            String folder = aDirectory+"\\"+filesInDir[i];
            File f = new File(aDirectory + "\\" + filesInDir[i]);
            //t1.start();
            System.out.println(f);
            listFilesForFolder(f, filename);
            new Compile().listFile(folder, filename);
        }


    }

    public static void listFilesForFolder(final File folder, String ext) {
        GenericExtFilter filter = new GenericExtFilter(ext);
        for (final File fileEntry : folder.listFiles()) {
            if (fileEntry.isDirectory()) {
                listFilesForFolder(fileEntry, filename);
            } else {
                String[] list = fileEntry.list(filter);
                //System.out.println(fileEntry.getName());
                String temp = new StringBuffer(String.valueOf(folder)).append(File.separator)
                        .append(fileEntry).toString();
                System.out.println("file : " + temp);
            }
        }
    }


    public void listFile(String folder, String ext) {
        GenericExtFilter filter = new GenericExtFilter(ext);

        File dir = new File(folder);

        if(dir.isDirectory()==false){
            System.out.println("Directory does not exists : " + folder);
            return;
        }

        // list out all the file name and filter by the extension
        String[] list = dir.list(filter);

        if (list.length == 0) {
            System.out.println("no files end with : " + ext);
            return;
        }

        for (String file : list) {
            String temp = new StringBuffer(folder).append(File.separator)
                    .append(file).toString();
            System.out.println("file : " + temp);
        }
    }
    // inner class, generic extension filter
    public static class GenericExtFilter implements FilenameFilter {

        private String ext;

        public GenericExtFilter(String ext) {
            this.ext = ext;
        }

        public boolean accept(File dir, String name) {
            return (name.endsWith(ext));
        }
    }


    public void setFileName(String filen){
        this.filen= filen;
    }
    public static String getFileName(){
        return filen;
    }

}