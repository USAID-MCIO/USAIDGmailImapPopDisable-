package gov.usaid.googleapps.imappop3fixer;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.*;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.admin.directory.Directory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.gdata.client.GoogleService;
import org.apache.commons.configuration.Configuration;
import org.apache.commons.configuration.ConfigurationException;
import org.apache.commons.configuration.PropertiesConfiguration;
import org.apache.log4j.Logger;

import javax.swing.*;
import java.io.*;
import java.util.ArrayList;
import java.util.EmptyStackException;

import static gov.usaid.googleapps.imappop3fixer.Main.*;

public class Properties {

    private static Logger logger = Logger.getLogger(Properties.class);
    private static Properties properties = new Properties();

    /*Retrived from Json File*/
    private static GoogleClientSecrets googleClientSecrets;

    /*Retrieved from existing file or Authorization Method*/
    private static String accessToken;
    private static String refreshToken;

    /*Google API Objects*/
    private static Sheets sheetsService;
    private static GoogleService googleService;
    private static Directory directoryService;

    /*Retrieved from existing file or manual input*/
    private static String domain;
    private static String testMode;
    private static String masterSpreadsheetName;
    private static String exclusionSheetName;
    private static String reportsheetName;
    private static String adminEmail;
    private static String maxUsersPerRequest;

    private static ArrayList<String> propertiesList;
    private static ArrayList<String> automatedPropertiesList;
    private static Credential credential;
    private static Configuration configuration;
    private static String ROOT;
    private static final String omitList = "omitlist.txt";
    private static String spreadsheetURL;
    private static String masterSpreadsheetID;
    private static String propertiesFilePath;
    private static String clientSecretLocation;
    private static int imp3InstanceCounter;
    private static boolean devMode;
    static BufferedReader userInput;
    private static boolean loadedLocalFile;
    private static boolean firstSetup;
    private static java.util.List<String> SCOPES;

    protected Properties() {
        create();
    }


    //********************************************Methods used*********************************************/
    //Instantiates the class
    public void create() {
        logger.info("\n************************************************" +
                "\n************CONFIGURATION START*****************" +
                "\n************************************************");

        final String os = System.getProperty("os.name").toLowerCase();


        if (os.contains("linux")) {
            ROOT = "~\\opt\\imappop3fixerApp\\";
            logger.info("Linux filesystem will be used.");
        } else if (os.contains("windows")) {
            ROOT = "\\opt\\imappop3fixer\\";
            logger.info("Windows filesystem will be used.");
        }

        //Update file Location paths
        clientSecretLocation = ROOT + "client_secret.json";
        propertiesFilePath = ROOT + "imappop3fixer.properties";

        //Read in Client Secret File
        if (googleClientSecrets == null || googleClientSecrets.isEmpty()) {
            // read in data from JSON for use
            readClientSecretFile();
        }

        //Set the propertiesList (Users will input this data themselves)
        propertiesList = new ArrayList<String>() {
            {
                add("ADMIN_EMAIL");
                add("DOMAIN");
                add("MASTERSPREADSHEET_NAME");
                add("SPREADSHEET_URL");
                add("EXCLUSIONSHEET_NAME");
                add("REPORTSHEET_NAME");
                add("MAX_USERS_PER_REQUEST");
                add("TEST_MODE");
            }
        };
        //Set the automatedropertiesList (Program will automatically gather and input this data)
        automatedPropertiesList = new ArrayList<String>() {{
            add("ACCESS_TOKEN");
            add("REFRESH_TOKEN");
            add("OMIT_LIST");
            /*add("MASTERSPREADSHEET_ID");*/
        }};
        //Set Scopes for this program
        SCOPES = new ArrayList<String>() {
            {
                add("https://www.googleapis.com/auth/admin.directory.group");
                add("https://www.googleapis.com/auth/apps.groups.settings");
                add("https://apps-apis.google.com/a/feeds/emailsettings/2.0/");
                add("https://spreadsheets.google.com/feeds");
                add("https://www.googleapis.com/auth/admin.directory.user");
                add("https://www.googleapis.com/auth/spreadsheets");
                add("https://www.googleapis.com/auth/drive");
                add("https://mail.google.com/");
            }
        };


        //While the configuration file is null
        while (configuration == null || configuration.isEmpty()) {
            //Try using pre-existing configuration from propertiesFilePath.
            try {
                loadedLocalFile = true;
                configuration = new PropertiesConfiguration(propertiesFilePath);

                if (configuration.isEmpty()) {
                    throw new ConfigurationException();
                }

            } catch (ConfigurationException e) {
                logger.info("Could not load the properties file at \"" + propertiesFilePath + "\", beginning setup...");
                //Begin configuration process to fix problem
                /*createNewPropertiesFile();*/
                loadedLocalFile = false;
                firstSetup = true;
                configuration = new PropertiesConfiguration();

                logger.info("Inital Setup: Type in a option:" +
                        "\nfind = Locate a properties file in your system" +
                        "\nnew = Create a new properties file via input" +
                        "\nexit = Exit program");
                userInput = new BufferedReader(new InputStreamReader(System.in));

                try {
                    // Read in user input
                    String answer = userInput.readLine().toLowerCase();

                    //Handle input
                    switch (answer) {
                        case "find":
                            logger.info("User selected Find, opening file browser window...");
                            // Try using a local file
                            configuration = new PropertiesConfiguration(
                                    openFileBrowswer("Locate the Properties file").getAbsolutePath());
                            break;

                        case "new":
                            logger.info("User selected New");
                            // Begin Manual Setup
                            logger.info("Beginning manual setup");
                            verifyConfiguration();
                            break;

                        case "exit":
                            logger.info("User chose to exit program..");
                            throw new EmptyStackException();

                        case "":
                            logger.info("User input not recognized... Exiting program");
                            throw new EmptyStackException();

                        case "dev":
                            verifyConfiguration();
                            break;
                    }

                } catch (Exception e0) {
                    logger.info("Program ended unexpectedly....");
                    displayErrors();
                    terminate();
                }
            }

            //After propertiesFile is created try reading in the data
            try {
                verifyConfiguration();
            } catch (IOException e1) {
                displayErrors();
                //Catches if configuration properties contains errors
                configuration.clear();
            }
        }
        logger.info("Configuration set!");


        //Deserialize configuration and pass it into the application
        logger.info("Deserializing configuration and passing into application...");
        domain = configuration.getString("DOMAIN");
        masterSpreadsheetName = configuration.getString("MASTERSPREADSHEET_NAME");
        spreadsheetURL = configuration.getString("SPREADSHEET_URL");
        masterSpreadsheetID = spreadsheetURL.substring(spreadsheetURL.lastIndexOf("=") + 1);
        exclusionSheetName = configuration.getString("EXCLUSIONSHEET_NAME");
        reportsheetName = configuration.getString("REPORTSHEET_NAME");
        adminEmail = configuration.getString("ADMIN_EMAIL");
        maxUsersPerRequest = configuration.getString("MAX_USERS_PER_REQUEST");

        //Validate Developer mode
        try {
            if (configuration.getString("DEV_MODE").toUpperCase().equals("TRUE")) {
                devMode = true;
                configuration.setProperty("TEST_MODE", "TRUE");
                logger.info("-=-=-=-=-=--=-=-DEV MODE ENABLED-=-=-=-=-=--=-=-");
            } else {
                throw new EmptyStackException();
            }
        } catch (Exception e) {
            devMode = false;
        }

        if (configuration.getString("TEST_MODE").toUpperCase().equals("TRUE")) {
            testMode = "TRUE";
        } else {
            testMode = "FALSE";
        }

        //Set the automated program properties
        automaticallyInputProgramProperties();

        logger.info("Configuration params passed successfully");

        // If Local file is not being used
        if (loadedLocalFile == false) {
            // Write data to a properties file
            writePropertiesToFile();
            logger.info("imappop3fixer properties file updated and the program must now restart....");
            terminate();
        }

        logger.info("\n************************************************" +
                    "\n************CONFIGURATION END*******************" +
                    "\n************************************************\n\n");

    }

    static void verifyConfiguration() throws IOException {

        //iterate through the propertiesList and update
        for (int i = 0; i < propertiesList.size(); i++) {
            //add current property name to a String
            String currentProperty = propertiesList.get(i);

            try {
                //if the configuration cannot find the string
                if (configuration.getString(currentProperty).equals("")
                        && !currentProperty.equals("TEST_MODE")) {
                    loadedLocalFile = false;

                    if (firstSetup != true) {
                        logger.info(currentProperty + " is not properly set in the configuration file.");
                    }
                    throw new EmptyStackException();
                }

            } catch (Exception e) {
                if (!currentProperty.equals("TEST_MODE")) {
                    logger.info("Please Enter the " + currentProperty + ":");
                    userInput = new BufferedReader(new InputStreamReader(System.in));
                    String input;
                    try {
                        input = userInput.readLine();
                        configuration.addProperty(currentProperty, input);
                    } catch (IOException e1) {
                        logger.info("User didn't input anything.. Ending program");
                        terminate();
                    }
                    i--;
                } else {
                    configuration.addProperty(currentProperty, "TRUE");
                }
            }
        }
    }

    static void automaticallyInputProgramProperties() {

        //If the ACCESS_TOKEN or REFRESH_TOKEN is empty
        try {
            if (configuration.getString("ACCESS_TOKEN").equals("") ||
                    configuration.getString("REFRESH_TOKEN").equals("")) {
                throw new EmptyStackException();
            }
        } catch (Exception e) {
            logger.info("Google Authorization Tokens are not set!");
            try {
                logger.info("The next step requires you to use a browser to access the Google Oauth2 Tokens." +
                        "\nPlease select an option:\n1: Automatically open browser and data for Google Token (Easy)" +
                        "\n2: Manually open browser then copy code back into this window");
                userInput = new BufferedReader(new InputStreamReader(System.in));
                String input = userInput.readLine();


                switch (input.trim()) {
                    case "1":
                        autoGetGoogleTokens();
                        break;
                    case "2":
                        manualGetGoogleToken();
                        break;
                    case "":
                        logger.info("User Did not select an option");
                        terminate();
                }


            } catch (Exception e1) {
                logger.info("ERROR: " + e.getMessage());
                terminate();
            }
        }

        accessToken = configuration.getString("ACCESS_TOKEN");
        refreshToken = configuration.getString("REFRESH_TOKEN");
        buildServices();


        //Iterate through the automated properties list for errors get the list of things to automate through
        for (String currentProperty : automatedPropertiesList) {

            try {
                configuration.getString(currentProperty);
                if (configuration.getString(currentProperty).equals("")) {
                    throw new EmptyStackException();
                }
            } catch (Exception e) {
                loadedLocalFile = false;

                switch (currentProperty.trim()) {
                    case "CLIENTSECRET_LOCATION":
                        configuration.setProperty(currentProperty, clientSecretLocation);
                        break;
                    case "OMITLIST":
                        configuration.setProperty(currentProperty, omitList);
                        break;
                }
                logger.info("Automatically updated the  " + currentProperty);
            }
        }
    }

    // Retrieves the Client Secret information from a file named "client_secret.json"
    static void readClientSecretFile() {


        try {
            File file = new File(clientSecretLocation);
            InputStream resourceAsStream = new FileInputStream(file);
            googleClientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(resourceAsStream));
            resourceAsStream.close();
        } catch (Exception e) {
            logger.info("Cannot load ClientSecret.json in " + ROOT);
            System.exit(0);
        }

    }

    //Opens a browser and automatically applies the Google authorization tokens to an object
    static void autoGetGoogleTokens() throws Exception {

        logger.info("The follow part of the program will open a new browser tab or window... " +
                "Please press enter to continue..");
        System.in.read();

        HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow =
                new GoogleAuthorizationCodeFlow.Builder(
                        HTTP_TRANSPORT, JSON_FACTORY, googleClientSecrets, SCOPES)
                        .setAccessType("offline")
                        .build();
        credential = new AuthorizationCodeInstalledApp(
                flow, new LocalServerReceiver()).authorize("user");
        //Passed Tokens into the configuration
        configuration.setProperty("ACCESS_TOKEN", credential.getAccessToken());
        configuration.setProperty("REFRESH_TOKEN", credential.getRefreshToken());
        logger.info("Access/RefreshTokens updated! Press Enter to continue.......");
        System.in.read();

    }

    //Opens a browser for user to manually copy the Google Access code and then applies the tokens to an object
    static void manualGetGoogleToken() throws IOException {
        // Creates a new authorization url using the information input
        logger.info("Creating Google Authorization Tokens using the ClientSecret file to create authorization URL......");

        //Generate a authorization Url using the clientId, redirectUri and Scopes
        String authorizationUrl = new GoogleAuthorizationCodeRequestUrl(googleClientSecrets.getInstalled().getClientId(),
                googleClientSecrets.getInstalled().getRedirectUris().get(0), SCOPES)
                .setAccessType("offline").setApprovalPrompt("force").build();

        // Point or redirect your user to the authorizationUrl.
        logger.info("Go to the following link in your browser:\n" + authorizationUrl
                + "\nPlease visit the URL above to authorize your OAuth" +
                "request token.  Once that is complete, type in your access code to continue...");


        // Read the authorization code from the standard input stream.
        logger.info("Enter the authorization code:");
        userInput = new BufferedReader(new InputStreamReader(System.in));
        String code = userInput.readLine().trim();

        HTTP_TRANSPORT = new NetHttpTransport();
        JSON_FACTORY = new JacksonFactory();

        //if Authorization Code is accepted
        if (code != null && !"".equals(code)) {
            GoogleTokenResponse response = null;
            try {
                response = new GoogleAuthorizationCodeTokenRequest(HTTP_TRANSPORT, JSON_FACTORY,
                        googleClientSecrets.getInstalled().getClientId(),
                        googleClientSecrets.getInstalled().getClientSecret(),
                        code,
                        googleClientSecrets.getInstalled().getRedirectUris().get(0)).execute();
            } catch (IOException e) {
                displayErrors();
            }

            //Passed Tokens into the configuration
            configuration.setProperty("ACCESS_TOKEN", response.getAccessToken());
            configuration.setProperty("REFRESH_TOKEN", response.getRefreshToken());
            logger.info("Access/RefreshTokens updated! Press enter to continue...");
            System.in.read();
        }
    }

    // Creates/Updates a properties file
    static void writePropertiesToFile() {
        // Updates to the properties file in the ROOT directory
        try {
            logger.info("Createing a new properties " +
                    "file in the \"" + propertiesFilePath + "\" directory. " +
                    "Press enter to continue.");
            System.in.read();

            File file = new File(propertiesFilePath);
            FileWriter fileWriter;
            fileWriter = new FileWriter(file.getAbsoluteFile());
            BufferedWriter bufferedWriter = new BufferedWriter(fileWriter);
            bufferedWriter.write("##Application Administrator's Email(Demo@gmail.com)"
                    + "\nADMIN_EMAIL=" + adminEmail.trim()
                    + "\n\n##Application's primary domain (gmail.com)"
                    + "\nDOMAIN=" + domain.trim()
                    + "\n\n##Application's current run mode.(TRUE or FALSE)"
                    + "\nTEST_MODE=" + testMode.trim()
                    + "\n\n##Application's Google SpreadSheetID"
                    + "\nSPREADSHEET_URL=" + spreadsheetURL.trim()
                    + "\n\n##Application's primary Spreadsheet name"
                    + "\nMASTERSPREADSHEET_NAME=" + masterSpreadsheetName.trim()
                    + "\n\n##Application's exlusion Worksheet name"
                    + "\nEXCLUSIONSHEET_NAME=" + exclusionSheetName.trim()
                    + "\n\n##Application's report Worksheet name"
                    + "\nREPORTSHEET_NAME=" + reportsheetName.trim()
                    + "\n\n##Application's maximum amount of users to update per token"
                    + "\nMAX_USERS_PER_REQUEST=" + maxUsersPerRequest
                    + "\n\n##------ALL VALUES BELOW ARE AUTOMATICALLY SET BY THE APPLICATION DO NOT ALTER!------"
                    + "\nOMIT_LIST=" + omitList.trim()
                    + "\nACCESS_TOKEN=" + accessToken.trim()
                    + "\nREFRESH_TOKEN=" + refreshToken.trim()
                    + "\nMASTERSPREADSHEET_ID=" + masterSpreadsheetID.trim());
            if (devMode == true) {
                bufferedWriter.append("\nDEV_MODE:TRUE");
            }
            bufferedWriter.close();
            logger.info("Created new imappop3fixer.properties file. in " + ROOT);
        } catch (Exception e) {
            displayErrors();
        }

    }

    // Google credential CORE AOUTH2 component
    static void buildServices() {
        HTTP_TRANSPORT = new NetHttpTransport();
        JSON_FACTORY = new JacksonFactory();
        credential = new GoogleCredential.Builder() //
                .setClientSecrets(googleClientSecrets.getInstalled().getClientId(),
                        googleClientSecrets.getInstalled().getClientSecret()) //
                .setJsonFactory(JSON_FACTORY) //
                .setTransport(HTTP_TRANSPORT) //
                .build() //
                .setAccessToken(accessToken) //
                .setRefreshToken(refreshToken);
        logger.info("Successfully built Google Credentials");

        /*BuildServices*/
        //Build Google Service
        googleService = new GoogleService("", appName);
        googleService.setOAuth2Credentials(credential);

        //Build Sheets Service
        sheetsService = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
                .setApplicationName(appName).build();

        //Build Directory Service
        directoryService = new Directory.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
                .setApplicationName(appName).build();

    }

    // Gets a file
    static File openFileBrowswer
    (String fileType) {
        File file;
        /* Creates new JFilechooser gui */
        JButton open = new JButton();
        JFileChooser fc = new JFileChooser();
        fc.setCurrentDirectory(new File("C:\\"));
        fc.setDialogTitle(fileType);
        fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
        if (fc.showOpenDialog(open) == JFileChooser.APPROVE_OPTION) {
            file = fc.getSelectedFile();
        } else {
            logger.info("User did not choose a file.");
            throw new EmptyStackException();
        }
        return file;
    }

    //Returns a sheet by url
    public static Spreadsheet getSheetByURL(String sheetURL) {
        Spreadsheet sheetToReturn = null;

        try {
            Spreadsheet spreadsheet = sheetsService.spreadsheets().get(sheetURL).execute();
            spreadsheet.getProperties().getTitle();
            sheetToReturn = spreadsheet;
            logger.info(spreadsheet.getProperties().getTitle());
        } catch (Exception e) {
            e.printStackTrace();
        }

        return sheetToReturn;
    }

    // Displays properties errors
    static void displayErrors() {
        String errorMessage = "The ";

        for (int i = 0; i < propertiesList.size(); i++) {
            try {
                configuration.getString(propertiesList.get(i)).toString();
            } catch (Exception e) {
                errorMessage += propertiesList.get(i) + ", ";
            }
        }
        // Log errors
        errorMessage += "is not properly set in the properties file located at " + propertiesFilePath;
        logger.info(errorMessage);
    }

   


    //--------------------------Properties getters and setters--------------------------/
    public static Properties getInstance() {
        System.getProperty("os.name");
        imp3InstanceCounter++;
        return properties;
    }

    public static String getDomain() {
        return domain;
    }

    public static String isTestmode() {
        return testMode;
    }

    public static String getAdminEmail() {
        return adminEmail;
    }

    public static String getMaxUsersPerRequest() {
        return maxUsersPerRequest;
    }

    public static int getImp3InstanceCounter() {
        return imp3InstanceCounter;
    }


    public static String getOmitList() {
        return omitList;
    }

    public static String getROOT() {
        return ROOT;
    }

    public static boolean isDevMode() {
        return devMode;
    }

    public static String getExclusionSheetName() {
        return exclusionSheetName;
    }

    public static String getReportsheetName() {
        return reportsheetName;
    }

    public static String getMasterSpreadsheetName() {
        return masterSpreadsheetName;
    }

    public static String getMasterSpreadsheetID() {
        return masterSpreadsheetID;
    }

    public static Sheets getSheetsService() {
        return sheetsService;
    }

    public static GoogleService getGoogleService() {
        return googleService;
    }

    public static Directory getDirectoryService() {
        return directoryService;
    }


}
