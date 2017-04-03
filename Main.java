package gov.usaid.googleapps.imappop3fixer;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.StringReader;
import java.net.MalformedURLException;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathFactory;

import org.apache.commons.configuration.ConfigurationException;
import org.apache.log4j.Logger;
import org.w3c.dom.Document;
import org.xml.sax.InputSource;

import com.google.api.client.http.HttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.admin.directory.Directory;
import com.google.api.services.admin.directory.model.User;
import com.google.api.services.admin.directory.model.Users;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;
import com.google.gdata.client.GoogleService;
import com.google.gdata.data.appsforyourdomain.generic.GenericEntry;
import com.google.gdata.util.ServiceException;

public class Main {


    /*---------------------------------------------Global Variables---------------------------------------------*/
    //Timer start
    static Long tStart = System.currentTimeMillis();
    static Logger logger = Logger.getLogger(Main.class);
    //Application Name
    static final String appName = "ImapPop3Fixer";
    static HttpTransport HTTP_TRANSPORT = new NetHttpTransport();
    static JacksonFactory JSON_FACTORY = new JacksonFactory();

    //Gather the properties
    static Properties properties = Properties.getInstance();

    //Google API Objects
    static final Sheets sheetsService = Properties.getSheetsService();
    static final GoogleService googleService = Properties.getGoogleService();
    static final Directory directoryService = Properties.getDirectoryService();
    //Sheets used
    static Sheet exclusionSheet = getSheetByName(properties.getExclusionSheetName());
    static Sheet reportSheet = getSheetByName(properties.getReportsheetName());

    //Global variables used
    //Default email settings URL
    static String emailsettingsUrl = "https://apps-apis.google.com/a/feeds/emailsettings/2.0/";

    //List of users that are excluded
    static ArrayList<String> omitList = new ArrayList<>();
    //List of users that will be reported
    static ArrayList<String> failedUsersList = new ArrayList<>();
    //Total Users tested
    static int totalTested;
    //Total USers excluded
    static int totalUsersExcluded;


    /*-------------------------------------------------Main Program-------------------------------------------------*/
    public static void main(String[] args) throws IOException, ConfigurationException {

        /*-----------------------SandBox--------------------*/
        //Only ran if DEV_MODE:TRUE in the imappop3fixer.properties file
        if (properties.isDevMode()) {
            updateSingleUserProtocols("rhenderson", true);
            testAPIS("rhenderson@dev.usaid.gov");
        }
        /*-----------------------SandBox--------------------*/

        //Check Program Mode and display message
        logger.info("Starting " + appName + "\n\n");

         /* Check if testmode is enabled. The test mode is set in the oauth
         * configuration file found at the root level of this server. This is
		 * important because if it is not in test mode it will update the pop
		 * and Imap settings of every email account within the DOMAIN_NAME*/

        if (properties.isTestmode().toLowerCase().equals("true")) {
            logger.info("\n#####################\n" + "##TEST MODE ENABLED##\n" + "#####################");

        }
        // If test mode is not enabled then log that it is live and email admins
        else {
            logger.info("\n<<<<<<<<<<<<----------------->>>>>>>>>>>>\n" + "<<<<<<<<<<<<----------------->>>>>>>>>>>>\n"
                    + "<<<<<<<<<<<<----------------->>>>>>>>>>>>\n" + "<<                                     >>\n"
                    + "<<          " + appName + " is Live!!    >>\n" + "<<                                     >>\n"
                    + "<<<<<<<<<<<<----------------->>>>>>>>>>>>\n" + "<<<<<<<<<<<<----------------->>>>>>>>>>>>\n"
                    + "<<<<<<<<<<<<----------------->>>>>>>>>>>>");
        }


        //Generate the omitList
        omitList = createOmitList();

        //Test and Report all users
        checkAndFixAllUsers();

        //Update Report Sheet
        insertIntoReportSheet();

        //log users that are being reported
        logger.info("\n==================\n======Report======\n==================");
        logger.info("Total users: " + totalTested);
        logger.info("Total exluded: " + totalUsersExcluded);
        logger.info("Total failed: " + failedUsersList.size());
        //End and log Timer
        long tEnd = System.currentTimeMillis();
        long tDelta = tEnd - tStart;
        double elapsedSeconds = tDelta / 1000.00;
        logger.info("Total elasped time: " + elapsedSeconds + " seconds");
        logger.info("\n==================\n====End Report====\n==================\n");

        //Log Debug info
        logger.info("\n%%%%%%%%%%%%%%%\n%%DEBUG STUFF%%\n%%%%%%%%%%%%%%%");
        logger.info("Oauth2 instances called: " + properties.getImp3InstanceCounter() + " times");
        //logger.info("SheetsApi instance called: " + sheetsAPI.getSheetsApiIterator() + " times");
        logger.info("\n%%%%%%%%%%%%%%%\n%%DEBUG STUFF%%\n%%%%%%%%%%%%%%%\n");

        //log program end message
        logger.info("\n###############\n##" + appName + " End##\n###############");
        //Terminate program
        terminate();

    }

    /*-----------------------------------------------MainProgram Methods--------------------------------------------*/
    // Returns list of users from exclude.txt
    static ArrayList<String> createOmitList() {

        logger.info("\n******************************\n**Updating/Creating OmitList**\n******************************");
        // Updates the exclude.txt with list from ExclusionList sheet
        try {
            // Creates new file for exclusionList to reside on the server at
            // root level
            File file = new File(properties.getROOT() + properties.getOmitList());
            FileWriter fileWriter = new FileWriter(file.getAbsoluteFile());
            BufferedWriter bufferedWriter = new BufferedWriter(fileWriter);

            ArrayList<String> users = readAllRowsOfAsheet(exclusionSheet, "A2", "B");
            // Retrieve List of users by iterating through the Exclude sheet
            for (String userEmail : users) {
                //Write the user to the local OmitList
                bufferedWriter.write(userEmail);
                bufferedWriter.newLine();
                //Add the users to the excluded users Arraylist
                omitList.add(userEmail);
            }

            bufferedWriter.close();
            logger.info("Updated exclude.text file.");
        } catch (Exception e) {
            e.printStackTrace();
        }
        logger.info("\n***************************************\n**Finished Updating/Creating OmitList**\n***************************************\n");
        return omitList;

    }

    //Uses the Directory object to Obtain information on a Users Email settings and updates each on accordingly
    static void checkAndFixAllUsers() {

        //Log the current batch number
        int batchCount = 0;

        logger.info("\n*****************\n**RUNNING TESTS**\n*****************");
        //Create an object to store a List of Users from directory
        Directory.Users.List requests = null;

        //Try Set the requests to handle a max amount of users
        try {
            requests = directoryService.users().list().setDomain(properties.getDomain()).setMaxResults(Integer.valueOf(properties.getMaxUsersPerRequest()));
        } catch (IOException e) {
            e.printStackTrace();
        }

        //Create a Users object to hold the users of the current token
        Users currentTokenUsers = null;

        //Request the currentToken User set
        try {
            currentTokenUsers = requests.execute();
        } catch (IOException e) {
            if (e.getLocalizedMessage().contains("Authorized")) {
                logger.info("The account \"" + properties.getAdminEmail() + "\" is not authorized to run this application " +
                        "within the " + properties.getDomain() + " domain.\n" +
                        "Please ensure that all scopes and API's are enabled before running this program. Once" +
                        " completed start begin the setup process to reconfigure this machine to " +
                        "properly run the program.");
            }
            //logger.info("ERROR: Properties file not properly configured. Beginning setup...");
            properties.create();
        }


        do {
            try {
                //increment logBatchCount
                batchCount++;
                User user;
                String userEmail;
                String userName;
                String userFullName;

                //Iterate the the current user set
                logger.info("=========Start batch [" + batchCount + "] ============");
                for (int i = 0; i < currentTokenUsers.getUsers().size(); i++) {
                    user = currentTokenUsers.getUsers().get(i);
                    userEmail = user.getPrimaryEmail();
                    userName = user.getPrimaryEmail().split("@")[0];
                    userFullName = user.getName().getFullName();

                    //If current user is not on omit/Exlusion list
                    if (!omitList.contains(userEmail)) {
                        //logger.info("Checking: " + userEmail);

                        //Check the users protocols
                        if (getSingleUserCurrentProtocolSetting(userEmail) == true) {

                            //if protocols are then enabled log, update, and add to list
                            logger.info("***FAILED***: [" + userEmail + "]");
                            failedUsersList.add(userEmail + ", (" + userFullName + ")");


                            //If in testmode program will not update the users imap/pop settings
                            if (!properties.isTestmode().toLowerCase().equals("true")) {
                                updateSingleUserProtocols(userName, false);
                            } else {
                                logger.info("------------TESTMODE WILL NOT UPDATE: " + userName);
                            }
                        } else {
                            logger.info("Passed: [" + userEmail + "]");
                        }
                    }
                    //else user must be on exclusionList
                    else {
                        logger.info("#######EXCLUDED########: " + userEmail);
                        totalUsersExcluded++;
                    }
                    //Increment totalTested
                    totalTested++;
                }
                //Get next token
                requests.setPageToken(currentTokenUsers.getNextPageToken());

                //Execute requests
                currentTokenUsers = requests.execute();
            } catch (Exception e) {
                //Catch when out of users/tokens
                e.printStackTrace();
                requests.setPageToken(null);
            }
        } while (requests.getPageToken() != null && requests.getPageToken().length() > 0);

        logger.info("\n***************\n***TESTS END***\n***************\n");
    }

    //Return a boolean based on the status of Imap and Pop settings for a particular user
    static boolean getSingleUserCurrentProtocolSetting(String username) throws IOException, ServiceException {
        username = username.substring(0, username.indexOf("@"));
        ArrayList<String> xmlType = new ArrayList<>();
        xmlType.add("imap");
        xmlType.add("pop");

        // Parse XML using XPath to find value for property 'enable'
        String response;
        boolean answer = false;

        for (int i = 0; i < xmlType.size(); i++) {
            // Get the value as GenericEntry
            GenericEntry entry = googleService.getEntry(new URL(emailsettingsUrl + properties.getDomain() + "/" + username + "/" + xmlType.get(i)), GenericEntry.class);

            //Add root node to make it valid xml string
            String rootNode = "<resp>" + entry.getXmlBlob().getBlob() + "</resp>";

            try {
                InputSource source = new InputSource(new StringReader(rootNode));

                DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
                dbf.setValidating(false);
                DocumentBuilder db = dbf.newDocumentBuilder();
                Document document = db.parse(source);

                XPathFactory xpathFactory = XPathFactory.newInstance();
                XPath xpath = xpathFactory.newXPath();

                XPathExpression expr = xpath.compile("//*[name()='apps:property'][@name='enable']/@value");
                response = expr.evaluate(document);

                if (response.equals("true")) {
                    answer = true;
                    break;
                }

            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return answer;
    }

    //Updates a single user's pop3 and imap settings to true(enabled) or false(disabled)
    static void updateSingleUserProtocols(String username, boolean setTo) {

        if (username.contains("@")) {
            username = username.substring(0, username.indexOf("@"));
        }

        //Entry needed for RestUrul
        GenericEntry entry = new GenericEntry();
        entry.addProperty("enable", String.valueOf(setTo));

        /*Try updating the Imap then the Pop. Imap must complete first because using the extra
        properties that the pop settings require will cause it to fail*/
        try {
            //Update Imap
            URL restUrlImap = new URL(emailsettingsUrl + properties.getDomain() + "/" + username + "/imap");
            //execute the update
            googleService.update(restUrlImap, entry);

            //Update Pop
            entry.addProperty("enableFor", "MAIL_FROM_NOW_ON");
            entry.addProperty("action", "KEEP");
            URL restUrlPop = new URL(emailsettingsUrl + properties.getDomain() + "/" + username + "/pop");
            //execute the update
            googleService.update(restUrlPop, entry);

        } catch (MalformedURLException e) {
            e.printStackTrace();
        } catch (ServiceException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        //Log
        logger.info("----------Updated----------" + username + " Imap and Pop3 protocols are now set to " + setTo);
    }

    //ran only if "DEV_MODE:TRUE" is applied in properties file
    static void runDevMode() {
        @SuppressWarnings("unchecked")
        ArrayList<String> methodList = new ArrayList() {{
            add(0, "Main Menu");
            add(1, "createOmitList();");
            add(2, "getSingleUserCurrentProtocolSetting();");
            add(3, "checkAndFixAllUsers();");
            add(4, "updateSingleUserProtocols();");
            add(5, "insertIntoReportSheet();");
            add(6, "Random Methods");
        }};

        String methods = "\nEnter the number of the method to Run:\n";
        for (int i = 0; i < methodList.size(); i++) {
            methods += i + ": " + methodList.get(i) + "\n";
        }

        boolean keepDebugging = true;
        String status;

        while (keepDebugging == true) {
            logger.info(methods);
            try {
                BufferedReader userInput = new BufferedReader(new InputStreamReader(System.in));
                int selection = Integer.parseInt(userInput.readLine());
                String username;
                boolean answer;
                switch (selection) {
                    case 0:
                        break;
                    case 1:
                        createOmitList();
                        break;
                    case 2:
                        logger.info("Enter a email account to check");
                        userInput = new BufferedReader(new InputStreamReader(System.in));
                        username = userInput.readLine();
                        answer = (getSingleUserCurrentProtocolSetting(username));
                        logger.info(username + "'s protocol settings are set to : " + answer);
                        break;
                    case 3:
                        checkAndFixAllUsers();
                        break;
                    case 4:
                        logger.info("Enter a email account to update");
                        userInput = new BufferedReader(new InputStreamReader(System.in));
                        username = userInput.readLine();
                        logger.info("Enter status: enter true or leave blank for false");
                        userInput = new BufferedReader(new InputStreamReader(System.in));
                        status = userInput.readLine();
                        answer = status.toLowerCase().equals("true");
                        updateSingleUserProtocols(username, answer);
                        break;
                    case 5:
                        insertIntoReportSheet();
                        break;
                    case 6:
                        break;
                }

                logger.info("Select an option Below:\n1: Run method\n2: Run " + appName + "\n0: Exit");
                userInput = new BufferedReader(new InputStreamReader(System.in));
                selection = Integer.parseInt(userInput.readLine());
                if (selection == 1) {
                    keepDebugging = true;
                } else if (selection == 2) {
                    break;
                } else if (selection == 0) {
                    System.exit(0);
                }

            } catch (Exception e) {
                logger.info("User input Error!!!!!!!!!\n");
                e.printStackTrace();
                keepDebugging = true;
            }
        }
    }

    //Global call to terminate program with "Press enter button" requests
    public static void terminate() {
        try {
            logger.info("Press enter to exit....");
            System.in.read();
            System.exit(0);
        } catch (Exception exit) {
            System.exit(0);
        }
    }


    /*----------------------------------------------Sheets Methods--------------------------------------------------*/
    //Reads all of the rows of a sheet based off of a set range
    static ArrayList<String> readAllRowsOfAsheet(Sheet sheetToUse, String range1, String range2) throws IOException {

        ArrayList<String> list = new ArrayList<>();

        String range = sheetToUse.getProperties().getTitle() + "!" + range1 + ":" + range2;
        ValueRange valueRange = sheetsService.spreadsheets().values()
                .get(properties.getMasterSpreadsheetID(), range).execute();

        List<List<Object>> values = valueRange.getValues();

        for (int i = 0; i < values.size(); i++) {
            //First get is the Cell number
            //Second et is the
            list.add(values.get(i).get(0).toString());
        }

        return list;
    }

    //Returns a Sheet object by name
    public static Sheet getSheetByName(String sheetName) {
        Sheet sheetToReturn = null;
        while (sheetToReturn == null) {
            try {
                Spreadsheet spreadsheet = sheetsService.spreadsheets().get(properties.getMasterSpreadsheetID()).execute();
                for (Sheet currentSheet : spreadsheet.getSheets()) {
                    if (currentSheet.getProperties().getTitle().equals(sheetName)) {
                        sheetToReturn = currentSheet;
                    }
                }

                //If no sheet is found create a new one
                if (sheetToReturn == null) {
                    createSheet(sheetName);
                }

            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        logger.info("Located Sheet: [" + sheetName + "] inside of [" + properties.getMasterSpreadsheetName() + "]");
        return sheetToReturn;
    }

    /*//Returns a sheet by url
    public static void getSheetByURL(String sheetURL) {
        Sheet sheetToReturn = null;

        try {
            Spreadsheet spreadsheet = getSheetsService().spreadsheets().get("12y87ZdB3LSCsiyU1HPcmf80qitqcSYrd-2BsJHIbqvM").execute();
            spreadsheet.getProperties().getTitle();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }*/

    //Updates the ReportSheet with the values gathered from this program
    static void insertIntoReportSheet() {
        clearSheet(reportSheet.getProperties().getSheetId());
        List<Request> requests = new ArrayList<>();
        // Date for logging
        SimpleDateFormat dateFormat = new SimpleDateFormat("MMM dd yyyy HH:mm:ss");
        Date timeStamp = new Date();

        int rowIterator = insertCellandReturnRowLocation(requests, "Report Date: " + dateFormat.format(timeStamp), 0);
        rowIterator = insertCellandReturnRowLocation(requests, "Number of users tested:  " + totalTested, rowIterator);
        rowIterator = insertCellandReturnRowLocation(requests, "Number of users that had to be fixed: " + failedUsersList.size(), rowIterator);
        rowIterator = insertCellandReturnRowLocation(requests, "Number of users excluded by the program:   " + totalUsersExcluded, rowIterator);
        rowIterator = insertCellandReturnRowLocation(requests, "List of users who had IMAP or POP enabled and had to be disabled by the program:", rowIterator);
        if (failedUsersList.size() > 0) {
            for (String userEmail : failedUsersList) {
                insertCellandReturnRowLocation(requests, userEmail, rowIterator);
            }
        } else {
            insertCellandReturnRowLocation(requests, "None", rowIterator);
        }

        executeSpreadsheetsRequests(requests);

        logger.info("");
    }

    //creates a new spreadsheet
    static void createSheet(String sheetName) {
        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setAddSheet(new AddSheetRequest()
                .setProperties(new SheetProperties().setTitle(sheetName))));
        executeSpreadsheetsRequests(requests);

    }

    //Returns the currentRow location after inserting a cell into a sheet
    static int insertCellandReturnRowLocation(List<Request> requests, String text, int counter) {
        requests.add(new Request()
                .setRepeatCell(new RepeatCellRequest()
                        .setRange(
                                new GridRange()
                                        .setSheetId(reportSheet.getProperties().getSheetId())
                                        .setStartRowIndex(counter)
                                        .setEndRowIndex(++counter)
                                        .setStartColumnIndex(0)
                                        .setEndColumnIndex(1)
                        )
                        .setCell(new CellData()
                                .setUserEnteredValue(new ExtendedValue()
                                        .setStringValue(text)
                                )
                        )
                        .setFields("userEnteredValue")));
        return counter;
    }

    //Deletes a sheet by sheetID
    static void deleteSheet(int sheetID) {
        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setDeleteSheet(new DeleteSheetRequest().setSheetId(sheetID)));
        executeSpreadsheetsRequests(requests);
    }

    //Clears a sheet by sheetID
    static void clearSheet(int sheetID) {
        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setUpdateCells(
                new UpdateCellsRequest().setRange(
                        new GridRange().setSheetId(sheetID)
                ).setFields("userEnteredValue")
        ));
        executeSpreadsheetsRequests(requests);
    }

    //Execute requests using the MasterSpreadsheet
    static void executeSpreadsheetsRequests(List<Request> requests) {
        try {
            sheetsService.spreadsheets().batchUpdate(properties.getMasterSpreadsheetID(), new BatchUpdateSpreadsheetRequest()
                    .setRequests(requests))
                    .execute();
        } catch (IOException e) {
            logger.info("Unable to run executeSpreadsheetsRequests() method:" + e);
        }
    }

    static void testAPIS(String userEmail) {

        logger.info("\n--------------------TESTING API's------------------------");

        //Test the Directory API
        try {
            Directory.Users.List requests = null;
            requests = directoryService.users().list().setDomain(properties.getDomain());
            requests.execute().toPrettyString();
            logger.info("Directory API: PASSED");

        } catch (IOException e) {
            logger.info("Directory API: FAILED");
        }

        //test the Sheets API
        try {
            Spreadsheet spreadsheet = sheetsService.spreadsheets().get(properties.getMasterSpreadsheetID()).execute();
            logger.info("SHEETS API: PASSED");
        } catch (Exception e) {
            logger.info("SHEETS API: FAILED");
        }

        logger.info("\n--------------------TESTING API's END------------------------");
    }


}