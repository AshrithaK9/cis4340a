
public final class CountBoxes implements Runnable {
  private static volatile int counter;
  // ...
  private final Object lock = new Object();
 
  @Override public void run() {
    synchronized (lock) {
      counter++;
      // ...
    }
  }
 
  public static void main(String[] args) {
    for (int i = 0; i < 2; i++) {
    new Thread(new CountBoxes()).start();
    }
  }
}
