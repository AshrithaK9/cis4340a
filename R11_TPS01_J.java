public final class ValidationService {
  // ...
  public StringBuilder fieldAggregator(String... inputs)
      throws InterruptedException, ExecutionException {
    // ...
    for (int i = 0; i < inputs.length; i++) {
      // Don't pass-in thread pool 
      results[i] = pool.submit(new ValidateInput<String>(inputs[i]));
    }
    // ...
  }
}
 
// Does not use same thread pool
public final class ValidateInput<V> implements Callable<V> {
  private final V input;
 
  ValidateInput(V input) {
    this.input = input;
  }
 
  @Override public V call() throws Exception {
    // If validation fails, throw an exception here
    return (V) new SanitizeInput().sanitize(input);
  }
}
 
public final class SanitizeInput<V> {  // No longer a Callable task
  public SanitizeInput() {}
 
  public V sanitize(V input) {
    // Sanitize input and return
    return input;
  }
}
