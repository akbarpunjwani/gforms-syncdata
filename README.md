# Google Forms: Sync Data Utility
This command line utility is useful for all who uses Google Forms for rapid data collection. Using the Google APIs this utility reads the responses of Google Forms from the Google Sheet which is linked with its respective form and returns the well-structured C# DataSet.

## Use Case / Business Scenario
Many organizations, especially the non-profits who works for social projects with regular data collection requirements but have very limited IT resources, may opt to use Google Forms for its ease of setup and availability of responses in its Google Spreadsheet. However, when any workflow is required to be attached with each entry of form response or when at later stages data analytics or integrated report from responses of multiple forms is required, this utility may help in following manner:
1. Projects which preliminary require data collection could be started quickly using the Google Forms design features.
2. Responses of each form is collected in its respective Google Spreadsheet
3. This utility is then used to fetch data from each of the Google Spreadsheet and convert it into C# DataSet structure. 
4. The DataSet is finally persisted into individual XML file, which can be later on read to further populate data into database or to perform next actions of the workflow against the new entry found.

## Usage
1. Clone this repository and open the solution file.
2. If required, use NuGet Package Manager Console to install the Google Data APIs. Below command can be used:
    ` PM> Install-Package Google.GData.Spreadsheets `
3. Provide the Google Drive Account key file in the Debug directory with the below name. ([Watch Video for Google Sources](https://www.youtube.com/watch?v=CXPRd8Hv2Y8)) 
` GAccount-Key.p12 `
4. Set the below parameters:
    - Google Spreadsheet Name
    - Count of Columns
    - Count of Rows
5. Call the below utility method, which will return DataTable:
` GetGoogleSheetData(FormName, true, TotalColumns, MaxRowId) `

## Future Work
This is currently the READONLY utility and hence provide the opportunity to temper the responses in the Google Spreadsheet and then persist to DB. Enhancement like READ and REMOVE responses from Google Spreadsheet would give assuarance to privacy savvy users that data is stationed on Google Spreadsheet only for period while it is not stored in DB. Contributions are welcome.

## How to make Contributions?
We’re eager to work with you, our user community, to improve these materials and develop new ones. Please check out [GOOGLE DATA SPREADSHEET API guide](https://developers.google.com/sheets/api/v3/data) for more information on getting started.
