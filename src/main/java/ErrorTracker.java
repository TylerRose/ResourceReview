
import java.util.ArrayList;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author trose
 */
public class ErrorTracker {

    private static ErrorTracker tracker = null;
    private ArrayList<String> erroredIDs;
    private ArrayList<String> messages;

    private ErrorTracker() {
        erroredIDs = new ArrayList<>();
        messages = new ArrayList<>();
    }

    public void addError(String listingID) {
        if (!erroredIDs.contains(listingID)) {
            erroredIDs.add(listingID);
            messages.add("");
        }
    }

    public void addError(String listingID, String message) {
        if (!erroredIDs.contains(listingID)) {
            erroredIDs.add(listingID);
            messages.add(message);
        } else {
            //append new message to existing errored ID
            messages.add(messages.indexOf(erroredIDs.indexOf(listingID)), messages.get(erroredIDs.indexOf(listingID)) + ", "+ message);
        }
    }

    public boolean checkForError(String listingID) {
        return erroredIDs.contains((listingID));
    }

    public int getNumOfErrors() {
        return erroredIDs.size();
    }

    public void printErrorList() {
        for (int i = 0; i < erroredIDs.size(); i++) {
            if (!messages.get(i).contains("HideError")) {
                System.out.println(erroredIDs.get(i) + "  -  " + messages.get(i));
            }
        }
    }

    public ArrayList<String> getList() {
        return erroredIDs;
    }

    public static ErrorTracker getInstance() {
        if (tracker == null) {
            tracker = new ErrorTracker();
            return tracker;
        } else {
            return tracker;
        }
    }
}
