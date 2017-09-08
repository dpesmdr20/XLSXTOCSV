/**
 * Created by dimanandhar on 9/6/17.
 */
public class PerfomanceTracker {
   static Runtime runtime;
    static long startTime,stopTime;
    private static final long MEGABYTE = 1024L * 1024L;


    public static void init(){
        runtime = Runtime.getRuntime();
        runtime.gc();
        startTime = System.currentTimeMillis();
    }
    public static void stop(){
       stopTime = System.currentTimeMillis();
    }
    public static long getMemoryUsage(){
        return bytesToMegabytes(runtime.totalMemory() - runtime.freeMemory());
    }
    public static long getRunTimePeriod(){
        return (stopTime-startTime)/1000;
    }

    public static long bytesToMegabytes(long bytes) {
        return bytes / MEGABYTE;
    }
}
