public final class NetworkHandler {
  private final ExecutorService executor;
 
  NetworkHandler(int poolSize) {
    this.executor = Executors.newFixedThreadPool(poolSize);
  }
 
  public void startThreads() {
    for (int i = 0; i < 3; i++) {
      executor.execute(new HandleRequest());
    }
  }
 
  public void shutdownPool() {
    executor.shutdown();
  }
 
  public static void main(String[] args)  {
    NetworkHandler nh = new NetworkHandler(3);
    nh.startThreads();
    nh.shutdownPool();
  }
}
