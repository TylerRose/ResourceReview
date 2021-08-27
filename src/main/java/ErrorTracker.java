
import java.util.ArrayList;

/**
 * The ErrorTracker class maintains a list of errors and their error messages in
 * parallel lists.
 *
 * @author Tyler Rose
 */
public class ErrorTracker {

    private static ErrorTracker tracker = null;
    private ArrayList<String> erroredIDs;
    private ArrayList<String> messages;

    /**
     * Private constructor for Singleton ErrorTracker. Initialize parallel lists
     */
    private ErrorTracker() {
        erroredIDs = new ArrayList<>();
        messages = new ArrayList<>();
    }

    /**
     * Add a Listing ID to the error list without a message
     *
     * @param listingID the ID to add
     */
    public void addError(String listingID) {
        if (!erroredIDs.contains(listingID)) {
            erroredIDs.add(listingID);
            messages.add("");
        }
    }

    /**
     * Add a Listing ID to the error list with a message
     *
     * @param listingID the ID to add
     * @param message the message about the error
     */
    public void addError(String listingID, String message) {
        if (!erroredIDs.contains(listingID)) {
            erroredIDs.add(listingID);
            messages.add(message);
        } else {
            //append new message to existing errored ID
            messages.add(messages.indexOf(erroredIDs.indexOf(listingID)), messages.get(erroredIDs.indexOf(listingID)) + ", " + message);
        }
    }

    /**
     * Check if a given Listing ID had an error
     *
     * @param listingID the ID to check
     * @return true if there was an error, false if not
     */
    public boolean checkForError(String listingID) {
        return erroredIDs.contains((listingID));
    }

    /**
     * Get the number of errors
     *
     * @return the number of errors
     */
    public int getNumOfErrors() {
        return erroredIDs.size();
    }

    /**
     * Print out the list of Listing ID's that had errors
     */
    public void printErrorList() {
        for (int i = 0; i < erroredIDs.size(); i++) {
            if (!messages.get(i).contains("HideError")) {
                System.out.println(erroredIDs.get(i) + "  -  " + messages.get(i));
            }
        }
    }

    /**
     * Get the list of Listing IDs that had errors
     *
     * @return the error list
     */
    public ArrayList<String> getList() {
        return erroredIDs;
    }

    /**
     * Get the ErrorTracker instance
     *
     * @return an ErrorTracker Instance
     */
    public static ErrorTracker getInstance() {
        if (tracker == null) {
            tracker = new ErrorTracker();
            return tracker;
        } else {
            return tracker;
        }
    }
}
