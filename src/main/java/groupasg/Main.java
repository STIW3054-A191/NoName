package groupasg;

import java.util.concurrent.ScheduledThreadPoolExecutor;

public class Main {
    ScheduledThreadPoolExecutor pool = new ScheduledThreadPoolExecutor(1);
    SubMain tpool = new SubMain();

    public static void main(String[] args)
    {
        new Main();
    }


    Main()
    {
        pool.execute(tpool);
        //pool.execute(tpool);
        pool.shutdown();
    }

}