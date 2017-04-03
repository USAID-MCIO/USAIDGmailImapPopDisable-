This application requires:
-Java 1.7 or above
-A Google project within the developers console with allowed scopes
-A folder titled "ImapPop3Fixer" in the /root/opt/ directory.
-A .json file containing the project clientSecret and clientid inside the application's folder.
-A .properties inside the application's folder.




######Setup Local environment###########################################################################################
1. Create a folder titled "imappop3fixer" and place it in the "root/opt/" directory.
(Example: "C:/opt/imappop3fixer/"
2. Place the ImapPop3Fixer.bat and the imappop3fixer.jar files in this directory.
(Example:
"C:/opt/imappop3fixer/ImapPop3Fixer.bat"
"C:/opt/imappop3fixer/ImapPop3Fixer.jar")

######Setting up The Google Project#####################################################################################
1. Using a browser, log on to
"https://console.developers.google.com/project"
using the email account that will be associated with this application.

2. Click on "Create Project"

3. Name the project "ImapPop3Fixer" then click Create.

4. Click the hambuger (3 flat lines) on the top left side of the screen then select
"API Manager".

5. In the left menu pane click on "Library"

6. Within the Library screen under Google APIs search for and enable these APIs:
Groups Settings API
Admin SDK
Google Identity and Access Management API
Gmail API

7. After enabling the 3 API's Click the hambuger on the top left side of the screen then select
"API Manager".

8. On the left sidebar click on Credentials.

9. Click on "Create Credintials" dropdown then select "OAuth Client ID"

10. Click on the "Configure consent screen" button. This will take you to the "Credentials/Oauth consent screen" tab.

11. On the "Oauth consent screen" select the Email that will be associated with this program as the Email address,
then type the ImapPop3Fixer name as the Product name. You can leave everything else blank and now click "Save".

10. On the "Credentials/Create client ID" screen choose the "Other" radio button and name the program "ImapPop3Fixer"
then click the create button.

11. A window will popup titled "OAuth client" that will contain your clientID and client Secret, click the OK button.

12. On the left sidebar click on Credentials then click on the ImapPop3Fixer under the OAuth 2.0 client IDs.

13. Once clicked, you should be taken to a page with a button labeled "Download JSON" in the upper portion. Click on
this "Download JSON" button.

14. Locate this downloaded Json file and rename it to "client_secrect".

15. Next place the "client_secret" file in the /root/opt/imappop3fixer/ directory.
(Example:"C:/opt/imappop3fixer/client_secret.json"



######Create a usable spreadsheet######################################################################################
1. Access the Google drive associated with the email acccount used for this program and create a new spreadsheet.

2. Inside the new spreadsheet create a worksheet titled exclusionlist to use as a list of user emails that will be
excluded from the program. It is very important to note the name of this worksheet because it will be needed for setting up this application.

3. This new worksheet should only contain 1 column with a header named "userid". Every row after A1 should have the
emails of the excluded users
Example:
Cell A1 = userid
Cell A2 = we@us.ORG
Cell A3 = him@her.net


######Starting the program for the first time###########################################################################
1. When running the program for the first time you will be asked:
Inital Setup: Type in a option:
find = Locate a properties file in your system
new = Create a new properties file via input
exit = Exit program



*FIND* OPTION--------------------------------------------------------
This option allows you to locate a pre-existing properties file using the filesystem browser. The file it finds should look
something like this:
###################Properties_File_Start###########################
##Application Administrator's Email
ADMIN_EMAIL=

##Application's primary domain
DOMAIN=

##Application's current run mode.
TEST_MODE=

##Application's primary Spreadsheet name
MASTERSPREADSHEET_NAME=

##Application's exlusion Worksheet name
EXCLUSIONSHEET_NAME=

##Application's report Worksheet name
REPORTSHEET_NAME=

##Application's maximum amount of users to update per token
MAX_USERS_PER_REQUEST=

##THESE VALUES ARE SET BY THE APPLICATION
CLIENTSECRET_LOCATION=
OMIT_LIST=
REDIRECT_URI=
ACCESS_TOKEN=
REFRESH_TOKEN=
###################Properties_File_End###########################
(If you want, you can just copy everything from Properties_File_Start to the Properties_File_End
and input everything necessary to begin the program)


*NEW* OPTION--------------------------------------------------------
This option allows you to manually write in everything needed to begin the application. It requires you to input 6
items
1.The ADMIN_EMAIL is the email account linked to the project
2.The DOMAIN is the domain name associated with the program
3.The MASTERSPREADSHEET_NAME is the name of the spreadsheet that will be used.
4.The EXCLUSIONSHEET_NAME is the sheet within the masterspreadsheet that contains the excluded users
5.The REPORTSHEET_NAME is the name of the sheet that will be created after the program is ran, it will contain the
information gathered and will input it out as a single sheet in a single column, you select the name of it.
6.The MAX_USERS_PER_REQUEST is the max amount of users the program will retrieve and update per Google Token.


*Authorizing Google*------------------------------------------------
During this part of the application you will have to
1.Copy and paste the url into your browser
2.Select the Google account associated with the project
3.Allow the application to access all parameters shown
4.Copy the authorization code Google provides
5.Paste the authorization code into the command prompt window