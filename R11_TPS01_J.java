public final class ValidationService {
  private final ExecutorService pool;
 
  public ValidationService(int poolSize) {
    pool = Executors.newFixedThreadPool(poolSize);
  }
 
  public void shutdown() {
    pool.shutdown();
  }
 
  public StringBuilder fieldAggregator(String... inputs)
      throws InterruptedException, ExecutionException {
 
    StringBuilder sb = new StringBuilder();
    // Stores the results
    Future<String>[] results = new Future[inputs.length];
 
    // Submits the tasks to thread pool
    for (int i = 0; i < inputs.length; i++) {
      results[i] = pool.submit(
        new ValidateInput<String>(inputs[i], pool));
    }
 
    for (int i = 0; i < inputs.length; i++) { // Aggregates the results
      sb.append(results[i].get());
    }
    return sb;
  }
}
 
public final class ValidateInput<V> implements Callable<V> {
  private final V input;
  private final ExecutorService pool;
 
  ValidateInput(V input, ExecutorService pool) {
    this.input = input;
    this.pool = pool;
  }
 
  @Override public V call() throws Exception {
    // If validation fails, throw an exception here
    // Subtask
    Future<V> future = pool.submit(new SanitizeInput<V>(input));
    return (V) future.get();
  }
}
 
public final class SanitizeInput<V> implements Callable<V> {
  private final V input;
 
  SanitizeInput(V input) {
    this.input = input;
  }
 
  @Override public V call() throws Exception {
    // Sanitize input and return
    return (V) input;
  }
}
