import java.util.regex.Pattern;
import java.util.logging.Logger;

public class R00_IDS03_J {
  private static final Logger logger = Logger.getLogger(R00_IDS03_J.class.getName());
public static void main(String[] args) {
        String username = "user123";
        boolean loginSuccessful = true;
        if (loginSuccessful) {
            logger.severe("User login succeeded for: " + sanitizeUser(username));
        } else {
            logger.severe("User login failed for: " + sanitizeUser(username));
        }
    }

    public static String sanitizeUser(String username) {
        return Pattern.matches("[A-Za-z0-9]+", username) ? username : "unauthorized user";
    }
}
