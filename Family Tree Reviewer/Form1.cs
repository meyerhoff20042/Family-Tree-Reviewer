using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using System.Windows.Forms;
using System.Linq;
using Meyerhoff_Family_Tree;
using System.IO;

namespace Family_Tree_Reviewer
{
    public partial class Form1 : Form
    {

        #region Fields

        readonly List<Person> people = new List<Person>();      // Holds all the people in the Meyerhoff Family Tree
        Person curPerson = new Person();                        // Tracks current person being displayed at all times
        int curIndex = 0;                                       // Displayed person's position in people/matches list
                                                                // Maps indexes to corresponding properties
        string currentTree = "";                                // Holds file path (used when selecting a new tree)
        bool isSelectingNewTree = false;                        // Used if user selects a new tree using the button
        bool searchMode = false;                                // Used for search state
        bool hasLoadedProperties = false;                       // Used if user needs to save new properties
        readonly List<Person> matches = new List<Person>();     // Holds matches for search terms

        // List of required properties for the program to work
        readonly List<string> requiredProperties = new List<string>()
        {
            "ID",
            "Full name",
            "Given names",
            "Nickname",
            "Title",
            "Suffix",
            "Surname now",
            "Surname at birth",
            "Gender",
            "Deceased",
            "Mother ID",
            "Mother name",
            "Father ID",
            "Father name",
            "Parents type",
            "Second mother ID",
            "Second mother name",
            "Second father ID",
            "Second father name",
            "Second parents type",
            "Third mother ID",
            "Third mother name",
            "Third father ID",
            "Third father name",
            "Third parents type",
            "Birth date type",
            "Birth year",
            "Birth month",
            "Birth day",
            "Birth range end",
            "Death date type",
            "Death year",
            "Death month",
            "Death day",
            "Death range end",
            "Partner ID",
            "Partner name",
            "Partner type",
            "Partnership date type",
            "Partnership year",
            "Partnership month",
            "Partnership day",
            "Partnership range end",
            "Ex-partner IDs",
            "Extra partner IDs",
            "Email",
            "Website",
            "Blog",
            "Photo site",
            "Home tel",
            "Work tel",
            "Mobile",
            "Skype",
            "Address",
            "Other contact",
            "Birth place",
            "Death place",
            "Cause of death",
            "Burial place",
            "Burial date type",
            "Burial year",
            "Burial month",
            "Burial day",
            "Burial range end",
            "Profession",
            "Company",
            "Interests",
            "Activities",
            "Bio Notes"
        };

        // List of properties that can be displayed in an item like a MessageBox
        #region PropertyList

        string propertyList =
            "ID" + "\n" +
            "Full name" + "\n" +
            "Given names" + "\n" +
            "Nickname" + "\n" +
            "Title" + "\n" +
            "Suffix" + "\n" +
            "Surname now" + "\n" +
            "Surname at birth" + "\n" +
            "Gender" + "\n" +
            "Deceased" + "\n" +
            "Mother ID" + "\n" +
            "Mother name" + "\n" +
            "Father ID" + "\n" +
            "Father name" + "\n" +
            "Parents type" + "\n" +
            "Second mother ID" + "\n" +
            "Second mother name" + "\n" +
            "Second father ID" + "\n" +
            "Second father name" + "\n" +
            "Second parents type" + "\n" +
            "Third mother ID" + "\n" +
            "Third mother name" + "\n" +
            "Third father ID" + "\n" +
            "Third father name" + "\n" +
            "Third parents type" + "\n" +
            "Birth date type" + "\n" +
            "Birth year" + "\n" +
            "Birth month" + "\n" +
            "Birth day" + "\n" +
            "Birth range end" + "\n" +
            "Death date type" + "\n" +
            "Death year" + "\n" +
            "Death month" + "\n" +
            "Death day" + "\n" +
            "Death range end" + "\n" +
            "Partner ID" + "\n" +
            "Partner name" + "\n" +
            "Partner type" + "\n" +
            "Partnership date type" + "\n" +
            "Partnership year" + "\n" +
            "Partnership month" + "\n" +
            "Partnership day" + "\n" +
            "Partnership range end" + "\n" +
            "Ex-partner IDs" + "\n" +
            "Extra partner IDs" + "\n" +
            "Email" + "\n" +
            "Website" + "\n" +
            "Blog" + "\n" +
            "Photo site" + "\n" +
            "Home tel" + "\n" +
            "Work tel" + "\n" +
            "Mobile" + "\n" +
            "Skype" + "\n" +
            "Address" + "\n" +
            "Other contact" + "\n" +
            "Birth place" + "\n" +
            "Death place" + "\n" +
            "Cause of death" + "\n" +
            "Burial place" + "\n" +
            "Burial date type" + "\n" +
            "Burial year" + "\n" +
            "Burial month" + "\n" +
            "Burial day" + "\n" +
            "Burial range end" + "\n" +
            "Profession" + "\n" +
            "Company" + "\n" +
            "Interests" + "\n" +
            "Activities" + "\n" +
            "Bio notes";

        #endregion

        // Used for search index
        readonly Dictionary<int, Func<Person, object>> propertySelectors = new Dictionary<int, Func<Person, object>>()
            {
            { 0, person => person.ID },
            { 1, person => person.Name.FullName },
            { 2, person => person.Name.FirstNames },
            { 3, person => person.Name.NickName },
            { 4, person => person.Name.Title },
            { 5, person => person.Name.Suffix },
            { 6, person => person.Name.LastName },
            { 7, person => person.Name.BirthName },
            { 8, person => person.Gender },
            { 9, person => person.Living },
            { 10, person => person.BirthDate.DateType },
            { 11, person => person.BirthDate.FullDate },
            { 12, person => person.BirthDate.Year },
            { 13, person => person.BirthDate.Month },
            { 14, person => person.BirthDate.Day },
            { 15, person => person.BirthDateRangeEnd.FullDate },
            { 16, person => person.DeathDate.DateType },
            { 17, person => person.DeathDate.FullDate },
            { 18, person => person.DeathDate.Year },
            { 19, person => person.DeathDate.Month },
            { 20, person => person.DeathDate.Day },
            { 21, person => person.DeathDateRangeEnd.FullDate },
            { 22, person => person.BirthLocation.FullAddress },
            { 23, person => person.BirthLocation.Country },
            { 24, person => person.BirthLocation.District },
            { 25, person => person.BirthLocation.County },
            { 26, person => person.BirthLocation.City },
            { 27, person => person.BirthLocation.Structure },
            { 28, person => person.DeathLocation.FullAddress },
            { 29, person => person.DeathLocation.Country },
            { 30, person => person.DeathLocation.District },
            { 31, person => person.DeathLocation.County },
            { 32, person => person.DeathLocation.City },
            { 33, person => person.DeathLocation.Structure },
            { 34, person => person.DeathCause },
            { 35, person => person.BurialLocation.FullAddress },
            { 36, person => person.BurialLocation.Country },
            { 37, person => person.BurialLocation.District },
            { 38, person => person.BurialLocation.County },
            { 39, person => person.BurialLocation.City },
            { 40, person => person.BurialLocation.Structure },
            { 41, person => person.BurialDate.DateType },
            { 42, person => person.BurialDate.FullDate },
            { 43, person => person.BurialDate.Year },
            { 44, person => person.BurialDate.Month },
            { 45, person => person.BurialDate.Day },
            { 46, person => person.BurialDateRangeEnd.FullDate },
            { 47, person => person.PartnerID },
            { 48, person => person.PartnerName.FullName },
            { 49, person => person.PartnerType },
            { 50, person => person.PartnershipDate.DateType },
            { 51, person => person.PartnershipDate.FullDate },
            { 52, person => person.PartnershipDate.Year },
            { 53, person => person.PartnershipDate.Month },
            { 54, person => person.PartnershipDate.Day },
            { 55, person => person.PartnershipDateRangeEnd.FullDate },
            { 56, person => person.ExPartnerIDs },
            { 57, person => person.ExPartnerIDs },
            { 58, person => person.ExtraPartnerIDs },
            { 59, person => person.ExtraPartnerIDs },
            { 60, person => person.FirstMotherID },
            { 61, person => person.FirstMotherName.FullName },
            { 62, person => person.FirstFatherID },
            { 63, person => person.FirstFatherName.FullName },
            { 64, person => person.FirstParentsType },
            { 65, person => person.SecondMotherID },
            { 66, person => person.SecondMotherName.FullName },
            { 67, person => person.SecondFatherID },
            { 68, person => person.SecondFatherName.FullName },
            { 69, person => person.SecondParentsType },
            { 70, person => person.ThirdMotherID },
            { 71, person => person.ThirdMotherName.FullName },
            { 72, person => person.ThirdFatherID },
            { 73, person => person.ThirdFatherName.FullName },
            { 74, person => person.ThirdParentsType },
            { 75, person => person.Email },
            { 76, person => person.Website },
            { 77, person => person.Blog },
            { 78, person => person.PhotoSite },
            { 79, person => person.HomeTelephone },
            { 80, person => person.WorkTelephone },
            { 81, person => person.Mobile },
            { 82, person => person.Skype },
            { 83, person => person.Address },
            { 84, person => person.OtherContactInfo },
            { 85, person => person.Profession },
            { 86, person => person.Company },
            { 87, person => person.Interests },
            { 88, person => person.Activities },
            { 89, person => person.BioNotes },
            { 90, person => person.FindAGrave },
            { 91, person => person.FamilySearch },
            { 92, person => person.WikiTree },
            { 93, person => person.Wikipedia }
            };

        #endregion

        #region Initializing Form

        // Constructor
        public Form1()
        {
            InitializeComponent();
        }


        // Called when the form first loads
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                LoadData();

                if (people.Count > 0)
                {
                    curIndex = 0;
                    curPerson = people[curIndex];
                    DisplayInfo(curPerson, people);

                    // Initialize AD/BC ListBoxes on "Extra Info" Tab
                    adBCListBox.SelectedIndex = 0;
                    adBCRangeEndListBox.SelectedIndex = 0;
                }

                // Update boolean if necessary
                if (isSelectingNewTree) isSelectingNewTree = false;
            }
            catch
            {
                MessageBox.Show("Failed to load the selected file. Make sure it holds" +
                    " all the required information. Properties listed below:\n" +
                    propertyList, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                Close();
            }
        }


        // This is used for error messages that instruct the user on which properties to enter
        // and where to put them. This method builds the columns for those instructions.
        string GetColumn(int i)
        {
            // Create strings to hold the column name and the letters of the English alphabet
            // The values of loops and digit are matched with the index of letters
            // NOTE: The space in letters is intentional (matches index with actual position)
            string column = "";
            string letters = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            // Check if number is a valid integer
            if (i > 0)
            {
                // Check if number is outside first set of letters
                if (i > 26)
                {
                    // These variables are used to find the two digits for the column name
                    // Loops is the first digit, digit is the second digit (A, B = "AB")
                    double loops = Math.Truncate((double)i / 26);
                    double digit = i - (26 * loops);

                    column += letters.ElementAt(Convert.ToInt32(loops));
                    column += letters.ElementAt(Convert.ToInt32(digit));
                }
                // Still on first set of letters
                else
                {
                    column += letters.ElementAt(i);
                }
            }
            // Negative or zero column (non-existent in Excel)
            else
            {
                column = "N/A";
            }

            return column;
        }


        // Lets the user select a file for the program to break down
        // Must be in .xlsx format, preferably a file from familyecho.com
        string GetFile()
        {
            // Create a variable to hold the file path
            var filePath = string.Empty;

            // Open a File Explorer window for the user to select a file
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.InitialDirectory = "c:\\";
                ofd.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                ofd.FilterIndex = 1;
                ofd.RestoreDirectory = true;

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    // Get the path of the specified file
                    filePath = ofd.FileName;
                }
            }

            // Keep original tree up if applicable
            if (filePath == string.Empty && isSelectingNewTree) filePath = currentTree;

            return filePath;
        }


        // Initializes the properties for the person whose information is being displayed
        // All Person properties that require the Location, Date, Name, and Property classes
        // need initialized before use
        void InitializeProperties(Person person)
        {
            person.Name = new Name();
            person.BirthDate = new Date();
            person.BirthDateRangeEnd = new Date();
            person.DeathDate = new Date();
            person.DeathDateRangeEnd = new Date();
            person.BurialDate = new Date();
            person.BurialDateRangeEnd = new Date();
            person.BirthLocation = new Location();
            person.DeathLocation = new Location();
            person.BurialLocation = new Location();
            person.PartnerName = new Name();
            person.PartnershipDate = new Date();
            person.PartnershipDateRangeEnd = new Date();
            person.FirstMotherName = new Name();
            person.FirstFatherName = new Name();
            person.SecondMotherName = new Name();
            person.SecondFatherName = new Name();
            person.ThirdMotherName = new Name();
            person.ThirdFatherName = new Name();
            person.ExtraProperties = new List<Property>();
        }


        // Loads information from the Excel spreadsheet into the program
        // Trees with lots of people can take a while to load
        void LoadData()
        {
            // Create variables to hold the file path and the list of properties
            string filePath = GetFile();
            List<string> properties = new List<string>();

            // If no file or an invalid file is selected, the program will close.
            if (!string.IsNullOrEmpty(filePath))
            {
                if (filePath != currentTree)
                {
                    if (filePath.EndsWith(".xlsx"))
                    {
                        // Clear previous People list (if applicable)
                        if (isSelectingNewTree) people.Clear();

                        // Load the Excel workbook
                        using (var workbook = new XLWorkbook(filePath))
                        {
                            // Select the first worksheet
                            var worksheet = workbook.Worksheet(1);

                            // Loop through the rows and columns
                            foreach (var row in worksheet.RowsUsed())
                            {
                                // The first row is used to find all properties in the table
                                if (row.RowNumber() != 1)
                                {
                                    // Create person objects for each person on the family tree
                                    Person person = new Person();
                                    InitializeProperties(person);

                                    int columnCount = row.Worksheet.ColumnCount();

                                    // Create variables to hold current cell and iteration counter
                                    var cell = row.Cell(1);

                                    // Add a property for each cell in the row as long as the cell has text
                                    for (int i = 0; i < properties.Count; i++)
                                    {
                                        cell = row.Cell(i + 1);

                                        // Load attributes
                                        switch (i)
                                        {
                                            case 0:
                                                if (properties.Contains("ID")) person.ID = cell.Value.ToString();
                                                break;
                                            case 1:
                                                if (properties.Contains("Full name")) person.Name.FullName = cell.Value.ToString();
                                                break;
                                            case 2:
                                                if (properties.Contains("Given names")) person.Name.FirstNames = cell.Value.ToString();
                                                break;
                                            case 3:
                                                if (properties.Contains("Nickname")) person.Name.NickName = cell.Value.ToString();
                                                break;
                                            case 4:
                                                if (properties.Contains("Title")) person.Name.Title = cell.Value.ToString();
                                                break;
                                            case 5:
                                                if (properties.Contains("Suffix")) person.Name.Suffix = cell.Value.ToString();
                                                break;
                                            case 6:
                                                if (properties.Contains("Surname now")) person.Name.LastName = cell.Value.ToString();
                                                break;
                                            case 7:
                                                if (properties.Contains("Surname at birth")) person.Name.BirthName = cell.Value.ToString();
                                                break;
                                            case 8:
                                                if (properties.Contains("Gender")) person.Gender = cell.Value.ToString();
                                                break;
                                            case 9:
                                                if (properties.Contains("Deceased"))
                                                {
                                                    if (cell.Value.ToString() == "Y")
                                                    {
                                                        person.Living = false;
                                                    }
                                                    else
                                                    {
                                                        person.Living = true;
                                                    }
                                                }
                                                break;
                                            case 10:
                                                if (properties.Contains("Mother ID")) person.FirstMotherID = cell.Value.ToString();
                                                break;
                                            case 11:
                                                if (properties.Contains("Mother name")) person.FirstMotherName.FullName = cell.Value.ToString();
                                                break;
                                            case 12:
                                                if (properties.Contains("Father ID")) person.FirstFatherID = cell.Value.ToString();
                                                break;
                                            case 13:
                                                if (properties.Contains("Father name")) person.FirstFatherName.FullName = cell.Value.ToString();
                                                break;
                                            case 14:
                                                if (properties.Contains("Parents type")) person.FirstParentsType = cell.Value.ToString();
                                                break;
                                            case 15:
                                                if (properties.Contains("Second mother ID")) person.SecondMotherID = cell.Value.ToString();
                                                break;
                                            case 16:
                                                if (properties.Contains("Second mother name")) person.SecondMotherName.FullName = cell.Value.ToString();
                                                break;
                                            case 17:
                                                if (properties.Contains("Second father ID")) person.SecondFatherID = cell.Value.ToString();
                                                break;
                                            case 18:
                                                if (properties.Contains("Second father name")) person.SecondFatherName.FullName = cell.Value.ToString();
                                                break;
                                            case 19:
                                                if (properties.Contains("Second parents type")) person.SecondParentsType = cell.Value.ToString();
                                                break;
                                            case 20:
                                                if (properties.Contains("Third mother ID")) person.ThirdMotherID = cell.Value.ToString();
                                                break;
                                            case 21:
                                                if (properties.Contains("Third mother name")) person.ThirdMotherName.FullName = cell.Value.ToString();
                                                break;
                                            case 22:
                                                if (properties.Contains("Third father ID")) person.ThirdFatherID = cell.Value.ToString();
                                                break;
                                            case 23:
                                                if (properties.Contains("Third father name")) person.ThirdFatherName.FullName = cell.Value.ToString();
                                                break;
                                            case 24:
                                                if (properties.Contains("Third parents type")) person.ThirdParentsType = cell.Value.ToString();
                                                break;
                                            case 25:
                                                if (properties.Contains("Birth date type")) person.BirthDate.DateType = cell.Value.ToString();
                                                break;
                                            case 26:
                                                if (properties.Contains("Birth year"))
                                                {
                                                    string str1 = cell.Value.ToString();

                                                    if (str1 != "")
                                                    {
                                                        person.BirthDate.Year = int.Parse(str1);
                                                    }
                                                }
                                                break;
                                            case 27:
                                                if (properties.Contains("Birth month"))
                                                {
                                                    string str2 = cell.Value.ToString();
                                                    if (str2 != "")
                                                    {
                                                        person.BirthDate.Month = int.Parse(str2);
                                                    }
                                                }
                                                break;
                                            case 28:
                                                if (properties.Contains("Birth day"))
                                                {
                                                    string str3 = cell.Value.ToString();
                                                    if (str3 != "")
                                                    {
                                                        person.BirthDate.Day = int.Parse(str3);

                                                        person.BirthDate.FullDate = person.BirthDate.GetFullDate(
                                                            person.BirthDate.Year, person.BirthDate.Month, person.BirthDate.Day);
                                                    }
                                                }
                                                break;
                                            case 29:
                                                if (properties.Contains("Birth range end"))
                                                {
                                                    if (cell.Value.ToString().Contains("-"))
                                                    {
                                                        if (!int.TryParse(cell.Value.ToString(), out int bred))
                                                        {
                                                            string[] numbers = cell.Value.ToString().Split('-');

                                                            person.BirthDateRangeEnd.FullDate = person.BirthDateRangeEnd.GetFullDate(
                                                                int.Parse(numbers[0]), int.Parse(numbers[1]), int.Parse(numbers[2]));
                                                        }
                                                        // Only the year is known
                                                        else
                                                        {
                                                            person.BirthDateRangeEnd.FullDate = person.BirthDateRangeEnd.GetYear(int.Parse(cell.Value.ToString()));
                                                        }
                                                    }
                                                    else
                                                    {
                                                        person.BirthDateRangeEnd.FullDate = cell.Value.ToString();
                                                    }
                                                }
                                                break;
                                            case 30:
                                                if (properties.Contains("Death date type")) person.DeathDate.DateType = cell.Value.ToString();
                                                break;
                                            case 31:
                                                if (properties.Contains("Death year"))
                                                {
                                                    string str4 = cell.Value.ToString();
                                                    if (str4 != "")
                                                    {
                                                        person.DeathDate.Year = int.Parse(str4);
                                                    }
                                                }
                                                break;
                                            case 32:
                                                if (properties.Contains("Death month"))
                                                {
                                                    string str5 = cell.Value.ToString();
                                                    if (str5 != "")
                                                    {
                                                        person.DeathDate.Month = int.Parse(str5);
                                                    }
                                                }
                                                break;
                                            case 33:
                                                if (properties.Contains("Death day"))
                                                {
                                                    string str6 = cell.Value.ToString();
                                                    if (str6 != "")
                                                    {
                                                        person.DeathDate.Day = int.Parse(str6);

                                                        person.DeathDate.FullDate = person.DeathDate.GetFullDate(
                                                            person.DeathDate.Year, person.DeathDate.Month, person.DeathDate.Day);
                                                    }
                                                }
                                                break;
                                            case 34:
                                                if (properties.Contains("Death range end"))
                                                {
                                                    if (cell.Value.ToString().Contains("-"))
                                                    {
                                                        if (!int.TryParse(cell.Value.ToString(), out int bred))
                                                        {
                                                            string[] numbers = cell.Value.ToString().Split('-');

                                                            person.DeathDateRangeEnd.FullDate = person.DeathDateRangeEnd.GetFullDate(
                                                                int.Parse(numbers[0]), int.Parse(numbers[1]), int.Parse(numbers[2]));
                                                        }
                                                        // Only the year is known
                                                        else
                                                        {
                                                            person.DeathDateRangeEnd.FullDate = person.DeathDateRangeEnd.GetYear(int.Parse(cell.Value.ToString()));
                                                        }
                                                    }
                                                    else
                                                    {
                                                        person.DeathDateRangeEnd.FullDate = cell.Value.ToString();
                                                    }
                                                }
                                                break;
                                            case 35:
                                                if (properties.Contains("Partner ID")) person.PartnerID = cell.Value.ToString();
                                                break;
                                            case 36:
                                                if (properties.Contains("Partner name")) person.PartnerName.FullName = cell.Value.ToString();
                                                break;
                                            case 37:
                                                if (properties.Contains("Partner type")) person.PartnerType = cell.Value.ToString();
                                                break;
                                            case 38:
                                                if (properties.Contains("Partnership date type")) person.PartnershipDate.DateType = cell.Value.ToString();
                                                break;
                                            case 39:
                                                if (properties.Contains("Partnership year"))
                                                {
                                                    string str7 = cell.Value.ToString();
                                                    if (str7 != "")
                                                    {
                                                        person.PartnershipDate.Year = int.Parse(str7);
                                                    }
                                                }
                                                break;
                                            case 40:
                                                if (properties.Contains("Partnership month"))
                                                {
                                                    string str8 = cell.Value.ToString();
                                                    if (str8 != "")
                                                    {
                                                        person.PartnershipDate.Month = int.Parse(str8);
                                                    }
                                                }
                                                break;
                                            case 41:
                                                if (properties.Contains("Partnership day"))
                                                {
                                                    string str9 = cell.Value.ToString();
                                                    if (str9 != "")
                                                    {
                                                        person.PartnershipDate.Day = int.Parse(str9);
                                                    }
                                                }
                                                break;
                                            case 42:
                                                if (properties.Contains("Partnership range end"))
                                                {
                                                    if (cell.Value.ToString().Contains("-"))
                                                    {
                                                        if (!int.TryParse(cell.Value.ToString(), out int bred))
                                                        {
                                                            string[] numbers = cell.Value.ToString().Split('-');

                                                            person.PartnershipDateRangeEnd.FullDate = person.PartnershipDateRangeEnd.GetFullDate(
                                                                int.Parse(numbers[0]), int.Parse(numbers[1]), int.Parse(numbers[2]));
                                                        }
                                                        // Only the year is known
                                                        else
                                                        {
                                                            person.PartnershipDateRangeEnd.FullDate = person.PartnershipDateRangeEnd.GetYear(int.Parse(cell.Value.ToString()));
                                                        }
                                                    }
                                                    else
                                                    {
                                                        person.PartnershipDateRangeEnd.FullDate = cell.Value.ToString();
                                                    }
                                                }
                                                break;
                                            case 43:
                                                if (properties.Contains("Ex-partner IDs"))
                                                {
                                                    if (!string.IsNullOrEmpty(cell.Value.ToString()))
                                                    {
                                                        person.ExPartnerIDs = cell.Value.ToString().Split(' ').ToList();
                                                    }
                                                }
                                                break;
                                            case 44:
                                                if (properties.Contains("Extra partner IDs"))
                                                {
                                                    if (!string.IsNullOrEmpty(cell.Value.ToString()))
                                                    {
                                                        person.ExtraPartnerIDs = cell.Value.ToString().Split(' ').ToList();
                                                    }
                                                }
                                                break;
                                            case 45:
                                                if (properties.Contains("Email")) person.Email = cell.Value.ToString();
                                                break;
                                            case 46:
                                                if (properties.Contains("Website")) person.Website = cell.Value.ToString();
                                                break;
                                            case 47:
                                                if (properties.Contains("Blog")) person.Blog = cell.Value.ToString();
                                                break;
                                            case 48:
                                                if (properties.Contains("Photo site")) person.PhotoSite = cell.Value.ToString();
                                                break;
                                            case 49:
                                                if (properties.Contains("Home tel")) person.HomeTelephone = cell.Value.ToString();
                                                break;
                                            case 50:
                                                if (properties.Contains("Work tel")) person.WorkTelephone = cell.Value.ToString();
                                                break;
                                            case 51:
                                                if (properties.Contains("Mobile")) person.Mobile = cell.Value.ToString();
                                                break;
                                            case 52:
                                                if (properties.Contains("Skype")) person.Skype = cell.Value.ToString();
                                                break;
                                            case 53:
                                                if (properties.Contains("Address")) person.Address = cell.Value.ToString();
                                                break;
                                            case 54:
                                                if (properties.Contains("Other contact")) person.OtherContactInfo = cell.Value.ToString();
                                                break;
                                            case 55:
                                                if (properties.Contains("Birth place")) person.BirthLocation.FullAddress = cell.Value.ToString();
                                                break;
                                            case 56:
                                                if (properties.Contains("Death place")) person.DeathLocation.FullAddress = cell.Value.ToString();
                                                break;
                                            case 57:
                                                if (properties.Contains("Cause of death")) person.DeathCause = cell.Value.ToString();
                                                break;
                                            case 58:
                                                if (properties.Contains("Burial place")) person.BurialLocation.FullAddress = cell.Value.ToString();
                                                break;
                                            case 59:
                                                if (properties.Contains("Burial date type")) person.BurialDate.DateType = cell.Value.ToString();
                                                break;
                                            case 60:
                                                if (properties.Contains("Burial year"))
                                                {
                                                    string str10 = cell.Value.ToString();
                                                    if (str10 != "")
                                                    {
                                                        person.BurialDate.Year = int.Parse(str10);
                                                    }
                                                }
                                                break;
                                            case 61:
                                                if (properties.Contains("Burial month"))
                                                {
                                                    string str11 = cell.Value.ToString();
                                                    if (str11 != "")
                                                    {
                                                        person.BurialDate.Month = int.Parse(str11);
                                                    }
                                                }
                                                break;
                                            case 62:
                                                if (properties.Contains("Burial day"))
                                                {
                                                    string str12 = cell.Value.ToString();
                                                    if (str12 != "")
                                                    {
                                                        person.BurialDate.Day = int.Parse(str12);

                                                        person.BurialDate.FullDate = person.BurialDate.GetFullDate(
                                                            person.BurialDate.Year, person.BurialDate.Month, person.BurialDate.Day);
                                                    }
                                                }
                                                break;
                                            case 63:
                                                if (properties.Contains("Burial range end"))
                                                {
                                                    if (cell.Value.ToString().Contains("-"))
                                                    {
                                                        if (!int.TryParse(cell.Value.ToString(), out int bred))
                                                        {
                                                            string[] numbers = cell.Value.ToString().Split('-');

                                                            person.BurialDateRangeEnd.FullDate = person.BurialDateRangeEnd.GetFullDate(
                                                                int.Parse(numbers[0]), int.Parse(numbers[1]), int.Parse(numbers[2]));
                                                        }
                                                        // Only the year is known
                                                        else
                                                        {
                                                            person.BurialDateRangeEnd.FullDate = person.BurialDateRangeEnd.GetYear(int.Parse(cell.Value.ToString()));
                                                        }
                                                    }
                                                    else
                                                    {
                                                        person.BurialDateRangeEnd.FullDate = cell.Value.ToString();
                                                    }
                                                }
                                                break;
                                            case 64:
                                                if (properties.Contains("Profession")) person.Profession = cell.Value.ToString();
                                                break;
                                            case 65:
                                                if (properties.Contains("Company")) person.Company = cell.Value.ToString();
                                                break;
                                            case 66:
                                                if (properties.Contains("Interests")) person.Interests = cell.Value.ToString();
                                                break;
                                            case 67:
                                                if (properties.Contains("Activities")) person.Activities = cell.Value.ToString();
                                                break;
                                            case 68:
                                                if (properties.Contains("Bio notes")) person.BioNotes = cell.Value.ToString();
                                                break;
                                            // Anything after Family Echo's attributes are sorted into an ExtraProperties list
                                            default:
                                                if (properties.Count > i)
                                                {
                                                    // Check if the value is empty
                                                    if (!string.IsNullOrEmpty(cell.Value.ToString()))
                                                    {
                                                        Property property = new Property(properties[i]);
                                                        property.Description = cell.Value.ToString();
                                                        person.ExtraProperties.Add(property);
                                                    }
                                                }

                                                break;
                                        }
                                    }

                                    // Add person to list
                                    people.Add(person);
                                }
                                // Add properties (Row 1 code)
                                else
                                {
                                    // Create variables to hold current cell and iteration counter
                                    var cell = row.Cell(1);
                                    int counter = 1;

                                    // Add a property for each cell in the row as long as the cell has text
                                    while (!string.IsNullOrEmpty(cell.Value.ToString()))
                                    {
                                        cell = row.Cell(counter);

                                        properties.Add(cell.Value.ToString());

                                        // Increment loop
                                        counter++;
                                    }

                                    // Check that all required Family Echo properties are present
                                    if (properties.Count < 69)
                                    {
                                        // If properties are missing, compare the property list with requiredProperties and build a string
                                        // containing the missing properties to display to the user.
                                        string missingProperties = "";

                                        for (int i = 0; i < requiredProperties.Count; i++)
                                        {
                                            if (!properties.Contains(requiredProperties[i]))
                                            {
                                                if (i > 1 && i < requiredProperties.Count - 1)
                                                {
                                                    missingProperties += "\n" + requiredProperties[i] + " at Column " + GetColumn(i + 1);
                                                }
                                            }
                                        }

                                        MessageBox.Show("Some properties are missing from this file that must be added" +
                                            " before the program can properly break down your family tree." +
                                            " Your chart needs the following properties: \n============================" + missingProperties, "Error", MessageBoxButtons.OK);

                                        Close();
                                    }
                                }
                            }

                            // Check if all people are actually present in the tree by seeing if they have a 
                            // full name and gender. Sometimes Family Echo doesn't actually remove someone when they're
                            // deleted and they still show up when the tree is downloaded.
                            for (int i = 0; i < people.Count; i++)
                            {
                                if (string.IsNullOrEmpty(people[i].Name.FullName) &&
                                    string.IsNullOrEmpty(people[i].Gender))
                                {
                                    people.RemoveAt(i);
                                    i--;
                                }
                                else
                                {
                                    people[i].BirthLocation.BuildLocations();
                                    people[i].DeathLocation.BuildLocations();
                                    people[i].BurialLocation.BuildLocations();
                                    people[i].GetResearchLinks();
                                }
                            }

                            // Update currentTree field if a file was successfully loaded
                            if (people.Count > 0) currentTree = filePath;

                            // Toggle navigation buttons if applicable
                            CheckButtons();
                        }
                    }
                    // Family Echo exports Excel spreadsheets in the .csv format
                    else
                    {
                        if (filePath.EndsWith(".csv"))
                        {
                            MessageBox.Show("File must be a .XLSX file. To create one from an" +
                                " imported Family Echo tree, copy all the contents from the .CSV" +
                                " file and paste them into an empty Excel spreadsheet. It will be" +
                                " saved as a .XLSX file by default.", "Error", MessageBoxButtons.OK);
                        }
                        else
                        {
                            // ClosedXML.Excel is only compatible with .XLSX files
                            MessageBox.Show("File must be a .XLSX file.", "Error", MessageBoxButtons.OK);
                        }

                        Close();
                    }
                }
            }
            else
            {
                if (!isSelectingNewTree)
                {
                    Close();
                }
            }
        }

        #endregion

        #region EventHandlers

        #region CheckedChanged EventHandlers

        // Displays DateRangeEnd controls when rangeRadioButton is checked
        private void rangeRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            // Range controls visible
            if (rangeRadioButton.Checked)
            {
                rangeBeginPromptLabel.Visible = true;
                rangeEndPromptLabel.Visible = true;
                monthRangeEndTextBox.Visible = true;
                dayRangeEndTextBox.Visible = true;
                yearRangeEndTextBox.Visible = true;
                adBCRangeEndListBox.Visible = true;
            }
            // Hide range controls
            else
            {
                rangeBeginPromptLabel.Visible = false;
                rangeEndPromptLabel.Visible = false;
                monthRangeEndTextBox.Visible = false;
                dayRangeEndTextBox.Visible = false;
                yearRangeEndTextBox.Visible = false;
                adBCRangeEndListBox.Visible = false;
            }
        }

        #endregion

        #region Click EventHandlers

        // When a user clicks on an ID label this fetches the person linked to that ID
        // and displays their information
        #region Person Accessors

        // Get partner information
        private void partnerIDLabel_Click(object sender, EventArgs e)
        {
            GetPerson(partnerIDLabel);
        }


        // Get first mother information
        private void firstMotherIDLabel_Click(object sender, EventArgs e)
        {
            GetPerson(firstMotherIDLabel);
        }


        // Get first father information
        private void firstFatherIDLabel_Click(object sender, EventArgs e)
        {
            GetPerson(firstFatherIDLabel);
        }


        // Get second mother information
        private void secondMotherIDLabel_Click(object sender, EventArgs e)
        {
            GetPerson(secondMotherIDLabel);
        }


        // Get second father information
        private void secondFatherIDLabel_Click(object sender, EventArgs e)
        {
            GetPerson(secondFatherIDLabel);
        }


        // Get third mother information
        private void thirdMotherIDLabel_Click(object sender, EventArgs e)
        {
            GetPerson(thirdMotherIDLabel);
        }


        // Get third father information
        private void thirdFatherIDLabel_Click(object sender, EventArgs e)
        {
            GetPerson(thirdFatherIDLabel);
        }

        #endregion


        // Buttons that navigate through everyone in the list
        // "First", "Previous", "Next", and "Last"
        #region Navigation Buttons

        // Displays info for the first person in the list
        private void firstButton_Click(object sender, EventArgs e)
        {
            curIndex = 0;

            if (searchMode == false)
            {
                curPerson = people[curIndex];

                DisplayInfo(curPerson, people);
            }
            else
            {
                curPerson = matches[curIndex];

                DisplayInfo(curPerson, matches);
            }
        }


        // Displays info for the last person in the list
        private void lastButton_Click(object sender, EventArgs e)
        {
            if (searchMode == false)
            {
                curIndex = people.Count - 1;
                curPerson = people[curIndex];
                DisplayInfo(curPerson, people);
            }
            else
            {
                curIndex = matches.Count - 1;
                curPerson = matches[curIndex];
                DisplayInfo(curPerson, matches);
            }
        }


        // Displays info for the next person in the list
        private void nextButton_Click(object sender, EventArgs e)
        {
            if (searchMode == false)
            {
                if (curIndex != people.Count - 1)
                {
                    curIndex++;

                    curPerson = people[curIndex];

                    DisplayInfo(curPerson, people);
                }
            }
            else
            {
                if (curIndex != matches.Count - 1)
                {
                    curIndex++;

                    curPerson = matches[curIndex];

                    DisplayInfo(curPerson, matches);
                }
            }
        }


        // Displays info for the previous person in the list
        private void previousButton_Click(object sender, EventArgs e)
        {
            if (curIndex != 0)
            {
                curIndex--;

                if (searchMode == false)
                {
                    curPerson = people[curIndex];
                    DisplayInfo(curPerson, people);
                }
                else
                {
                    curPerson = matches[curIndex];
                    DisplayInfo(curPerson, matches);
                }
            }
        }

        #endregion


        // Adds a new property in the "Extra Info" Tab
        private void addPropertyButton_Click(object sender, EventArgs e)
        {
            // Create variables to hold ListBox count and new property name
            int number = propertiesListBox.Items.Count + 1;
            string name = "Property #" + number.ToString();

            // Add property to ListBox and person's ExtraProperties list
            propertiesListBox.Items.Add(name);

            Property property = new Property(name);

            curPerson.ExtraProperties.Add(property);

            // Enable "Delete Property" button if applicable
            deletePropertyButton.Enabled = true;

            // Focus on new property
            propertiesListBox.SelectedIndex = propertiesListBox.Items.Count - 1;
            DisplayPropertyInfo(curPerson, property);
        }


        // Print a list of ancestors that are on a family tree
        private void ancestryButton_Click(object sender, EventArgs e)
        {
            // Create a tuple to hold IDs, people, and relationships
            var ancestors = new List<(int id, Person ancestor, string relationship)> { };

            // Get parents of subject
            void GetParents(Person startingPerson)
            {
                // Create booleans to track which parents have been found
                bool foundMother = false;
                bool foundFather = false;

                // Iterate through each person and find the parents by comparing IDs
                foreach (Person person in people)
                {
                    #region Variables

                    // Create strings to hold startingPerson's relationship to curPerson (personRelation)
                    // and the parents' relation to startingPerson (parentRelation)
                    string personRelation = "";
                    string parentRelation = "";

                    // Create a string to hold the new relationship to be added
                    string newRelation = "";

                    // Create a boolean to track if the startingPerson is already in the ancestors list
                    bool personInList = false;

                    #endregion

                    #region startingPerson's Relationship to curPerson

                    // Find startingPerson's relationship to curPerson if applicable
                    for (int i = 0; i < ancestors.Count; i++)
                    {
                        if (ancestors[i].ancestor == startingPerson)
                        {
                            personRelation = ancestors[i].relationship;
                            personInList = true;
                        }
                    }

                    #endregion

                    #region Conditional Statements for Mother/Father

                    // Found mother
                    if (person.ID == startingPerson.FirstMotherID && foundMother == false)
                    {
                        parentRelation = "M";
                        foundMother = true;
                    }
                    // Found father
                    if (person.ID == startingPerson.FirstFatherID && foundFather == false)
                    {
                        parentRelation = "F";
                        foundFather = true;
                    }

                    #endregion

                    #region startingPerson's Parents' Relationship to curPerson

                    // Found a match in people list
                    if (person.ID == startingPerson.FirstMotherID || person.ID == startingPerson.FirstFatherID)
                    {
                        // Combine relationships if startingPerson is an ancestor of curPerson
                        // i.e. "M" + "F" = "MF" (mother's father, or maternal grandfather)
                        if (personInList) newRelation = personRelation + parentRelation;
                        // Parents of curPerson only need "M" or "F"
                        else newRelation = parentRelation;

                        // Add person to ancestors list
                        ancestors.Add((ancestors.Count, person, newRelation));
                    }

                    #endregion

                    // Break the loop if both parents have been found
                    if (foundMother && foundFather) break;
                }
            }

            // Parents; these are retrieved regardless of ancestorsTrackBar's value
            GetParents(curPerson);

            // Retrieve ancestry; i is the number of generations to trace; j iterates through each person in the ancestry list
            for (int i = 0; i < ancestorsTrackBar.Value; i++)
            {
                for (int j = 0; j < ancestors.Count; j++)
                {
                    if (ancestors[j].relationship.Length == i) GetParents(ancestors[j].ancestor);
                }
            }

            // Create a string to hold the list of ancestors
            string ancestry = "";

            // Build list
            for (int i = 0; i < ancestors.Count; i++)
            {
                // Add dashes to indicate generations
                for (int j = 0; j < ancestors[i].relationship.Length; j++) ancestry += "-";

                // Add full name
                if (ancestors[i].ancestor.Name.LastName == ancestors[i].ancestor.Name.BirthName)
                {
                    ancestry += ancestors[i].ancestor.Name.FullName + "\r\n";
                }
                // Include birth name in parentheses when applicable
                else
                {
                    // Check if birth name is known
                    if (!string.IsNullOrEmpty(ancestors[i].ancestor.Name.BirthName))
                    {
                        ancestry += ancestors[i].ancestor.Name.FirstNames + " (" + ancestors[i].ancestor.Name.BirthName + ") "
                            + ancestors[i].ancestor.Name.LastName + "\r\n";
                    }
                    // Unknown birth name
                    else
                    {
                        ancestry += ancestors[i].ancestor.Name.FirstNames + " (UNKNOWN) "
                            + ancestors[i].ancestor.Name.LastName + "\r\n";
                    }
                }

                // Add birth year if known
                ancestry += "(";
                if (ancestors[i].ancestor.BirthDate.Year != 9999) ancestry += ancestors[i].ancestor.BirthDate.Year.ToString();
                else ancestry += "????";

                // Only show birth year is living
                if (ancestors[i].ancestor.Living) ancestry += ")" + "\r\n";
                // Add death year if the person is deceased & death year is known
                else
                {
                    // 9999 is the default year used when the death date is unknown
                    if (ancestors[i].ancestor.DeathDate.Year != 9999)
                    {
                        ancestry += "-" + ancestors[i].ancestor.DeathDate.Year.ToString() + ")\r\n";
                    }
                    // Unknown death year
                    else
                    {
                        ancestry += "-????)\r\n";
                    }
                }

                // Add relation
                ancestry += ancestors[i].relationship + "\r\n";
            }

            // Check if any ancestors were found
            if (!string.IsNullOrEmpty(ancestry))
            {
                // Get person's first and last name (no middle names)
                string[] firstNames = curPerson.Name.FirstNames.Split(' ');
                string shortName = firstNames[0] + " " + curPerson.Name.LastName;

                // Display family on FamilyForm
                FamilyForm familyForm = new FamilyForm(shortName, ancestry, "Ancestry");
                familyForm.ShowDialog();
            }
            // No ancestors found
            else
            {
                MessageBox.Show("No ancestors found for this individual.", "No results", MessageBoxButtons.OK);
            }
        }


        // Copies property information to the system's clipboard (WIP)
        private void copyToClipboardButton_Click(object sender, EventArgs e)
        {
            // Create variables to hold clipboard text and property
            // NOTE: Property name is initially set to this so it's highly unlikely a user will accidentally name a property the default name
            string propertyInfo;
            Property property = new Property("15dab045658b7bd0043298d4cc73b331e1559fece81aadd11ce64f5d9d63c205dbf33e4f593d2fa37516ca3ed41a5b0a");

            // Iterate through each property and compare the names
            foreach (Property property1 in curPerson.ExtraProperties)
            {
                // Found a match
                if (property1.Name == propertiesListBox.SelectedItem.ToString())
                {
                    property = property1;
                    break;
                }
            }

            // Check if a property was found
            if (property.Name != "15dab045658b7bd0043298d4cc73b331e1559fece81aadd11ce64f5d9d63c205dbf33e4f593d2fa37516ca3ed41a5b0a")
            {
                // Get property information
                propertyInfo = GetProperty(property);

                // Copy concatenated string to clipboard
                Clipboard.SetText(propertyInfo);
            }
            // Property was not found or not saved before trying to copy to clipboard
            else
            {
                MessageBox.Show("Property could not be found in the individual's information. Make sure a property has all necessary information " +
                    "before copying to the clipboard.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // Deletes the selected property in propertiesListBox on the "Extra Info" Tab
        private void deletePropertyButton_Click(object sender, EventArgs e)
        {
            // Check if a property was selected
            if (propertiesListBox.SelectedIndex != -1)
            {
                // Remove property from person's list
                for (int i = 0; i < curPerson.ExtraProperties.Count; i++)
                {
                    if (curPerson.ExtraProperties[i].Name == propertiesListBox.SelectedItem.ToString())
                    {
                        curPerson.ExtraProperties.RemoveAt(i);
                    }
                }

                // Remove property from ListBox
                propertiesListBox.Items.RemoveAt(propertiesListBox.SelectedIndex);

                // Disable "Delete Property" if there are none
                if (propertiesListBox.Items.Count == 0)
                {
                    deletePropertyButton.Enabled = false;
                }
            }

            // Focus on newest property if applicable
            if (propertiesListBox.Items.Count > 0)
            {
                propertiesListBox.SelectedIndex = propertiesListBox.Items.Count - 1;
                DisplayPropertyInfo(curPerson, curPerson.ExtraProperties[curPerson.ExtraProperties.Count - 1]);
            }
            // Clear all controls if there are no properties for this person
            else
            {
                ClearPropertyForm();
            }
        }


        // Print a list of descendants that are on a family tree
        private void descendantsButton_Click(object sender, EventArgs e)
        {
            // Create a tuple to hold IDs, people, and relationships
            var descendants = new List<(int id, Person descendant, string relationship)> { };

            // Get children of subject
            void GetChildren(Person startingPerson)
            {
                // Iterate through each person and find all children by comparing IDs
                foreach (Person person in people)
                {
                    #region Variables

                    // Create strings to hold startingPerson's relationship to curPerson (personRelation)
                    // and the children's relation to startingPerson (childrenRelation)
                    string personRelation = "";
                    string childrenRelation = "";

                    // Create a string to hold the new relationship to be added
                    string newRelation = "";

                    // Create a boolean to track if the startingPerson is already in the descendants list
                    bool personInList = false;

                    #endregion

                    #region startingPerson's Relationship to curPerson

                    // Find startingPerson's relationship to curPerson if applicable
                    for (int i = 0; i < descendants.Count; i++)
                    {
                        if (descendants[i].descendant == startingPerson)
                        {
                            personRelation = descendants[i].relationship;
                            personInList = true;
                        }
                    }

                    #endregion

                    #region Finding Children of startingPerson

                    // Found a match in people list
                    if (person.FirstMotherID == startingPerson.ID || person.FirstFatherID == startingPerson.ID)
                    {
                        // Get person's relation to startingPerson
                        if (person.Gender == "Male") childrenRelation = "S";
                        else childrenRelation = "D";

                        // Combine relationships if startingPerson is a descendant of curPerson
                        // i.e. "D" + "S" = "DS" (daughter's son, or grandson)
                        if (personInList) newRelation = personRelation + childrenRelation;
                        // Children of curPerson only need "D" or "S"
                        else newRelation = childrenRelation;

                        // Add person to ancestors list
                        descendants.Add((descendants.Count, person, newRelation));
                    }

                    #endregion

                }
            }

            // Create string to hold list of descendants
            string descendance = "";

            // Get children of curPerson
            GetChildren(curPerson);

            // Retrieve ancestry; i is the number of generations to trace; j iterates through each person in the ancestry list
            for (int i = 0; i < descendantsTrackBar.Value; i++)
            {
                for (int j = 0; j < descendants.Count; j++)
                {
                    if (descendants[j].relationship.Length == i) GetChildren(descendants[j].descendant);
                }
            }

            // Build list
            for (int i = 0; i < descendants.Count; i++)
            {
                // Add dashes to indicate generations
                for (int j = 0; j < descendants[i].relationship.Length; j++) descendance += "-";

                // Add full name
                if (descendants[i].descendant.Name.LastName == descendants[i].descendant.Name.BirthName)
                {
                    descendance += descendants[i].descendant.Name.FullName + "\r\n";
                }
                // Include birth name in parentheses when applicable
                else
                {
                    // Check if birth name is known
                    if (!string.IsNullOrEmpty(descendants[i].descendant.Name.BirthName))
                    {
                        descendance += descendants[i].descendant.Name.FirstNames + " (" + descendants[i].descendant.Name.BirthName + ") "
                            + descendants[i].descendant.Name.LastName + "\r\n";
                    }
                    // Unknown birth name
                    else
                    {
                        descendance += descendants[i].descendant.Name.FirstNames + " (UNKNOWN) "
                            + descendants[i].descendant.Name.LastName + "\r\n";
                    }
                }

                // Add birth year if known
                descendance += "(";
                if (descendants[i].descendant.BirthDate.Year != 9999) descendance += descendants[i].descendant.BirthDate.Year.ToString();
                else descendance += "????";

                // Only show birth year is living
                if (descendants[i].descendant.Living) descendance += ")" + "\r\n";
                // Add death year if the person is deceased & death year is known
                else
                {
                    // 9999 is the default year used when the death date is unknown
                    if (descendants[i].descendant.DeathDate.Year != 9999)
                    {
                        descendance += "-" + descendants[i].descendant.DeathDate.Year.ToString() + ")\r\n";
                    }
                    // Unknown death year
                    else
                    {
                        descendance += "-????)\r\n";
                    }
                }

                // Add relation
                descendance += descendants[i].relationship + "\r\n";
            }

            // Check if any ancestors were found
            if (!string.IsNullOrEmpty(descendance))
            {
                // Get person's first and last name (no middle names)
                string[] firstNames = curPerson.Name.FirstNames.Split(' ');
                string shortName = firstNames[0] + " " + curPerson.Name.LastName;

                // Display family on FamilyForm
                FamilyForm familyForm = new FamilyForm(shortName, descendance, "Descendance");
                familyForm.ShowDialog();
            }
            // No ancestors found
            else
            {
                MessageBox.Show("No descendants found for this individual.", "No results", MessageBoxButtons.OK);
            }
        }


        // Exits the program
        private void exitButton_Click(object sender, EventArgs e)
        {
            Close();
        }


        // Find and list all children of the person currently displayed
        private void findChildrenButton_Click(object sender, EventArgs e)
        {
            // Get ID of parent to compare people to
            string parentID = idLabel.Text;

            // Create a list to hold the person's children and their relationship to their parent
            List<Person> children = new List<Person>();
            List<string> childrenTypes = new List<string>();

            // Find all children of the person
            foreach (Person person in people)
            {
                if (person.FirstFatherID == parentID ||
                    person.FirstMotherID == parentID ||
                    person.SecondFatherID == parentID ||
                    person.SecondMotherID == parentID ||
                    person.ThirdFatherID == parentID ||
                    person.ThirdMotherID == parentID)
                {
                    // Find relationship type based on set of parents
                    if (person.FirstFatherID == parentID || person.FirstMotherID == parentID)
                    {
                        if (!string.IsNullOrEmpty(person.FirstParentsType))
                        {
                            childrenTypes.Add(person.FirstParentsType.ToString());
                        }
                    }

                    if (person.SecondFatherID == parentID || person.SecondMotherID == parentID)
                    {
                        if (!string.IsNullOrEmpty(person.SecondParentsType))
                        {
                            childrenTypes.Add(person.SecondParentsType.ToString());
                        }
                    }

                    if (person.ThirdFatherID == parentID || person.ThirdMotherID == parentID)
                    {
                        if (!string.IsNullOrEmpty(person.ThirdParentsType))
                        {
                            childrenTypes.Add(person.ThirdParentsType.ToString());
                        }
                    }

                    // Add child to list
                    children.Add(person);

                    // Add "Not specified" if the relationship type is not included
                    if (childrenTypes.Count < children.Count)
                    {
                        childrenTypes.Add("Not specified");
                    }
                }
            }

            // Clear previous results if applicable
            childrenLabel.Text = "";

            // If children were found, list them on the label
            if (children.Count > 0)
            {
                for (int i = 0; i < children.Count; i++)
                {
                    childrenLabel.Text += $"{i + 1}. {children[i].Name.FullName}, Relationship: {childrenTypes[i]} \n";
                }
            }
            else
            {
                childrenLabel.Text = "No children found for this individual.";
            }
        }


        // Find and list all siblings of the person currently displayed
        private void findSiblingsButton_Click(object sender, EventArgs e)
        {
            // Clear siblingsLabel if applicable
            siblingsLabel.Text = "";

            // Check if an individual has at least one set of parents
            if (!string.IsNullOrEmpty(firstFatherIDLabel.Text) || !string.IsNullOrEmpty(firstMotherIDLabel.Text))
            {
                // Create variables to hold the parent IDs
                string firstFatherID = ""; string firstMotherID = "";
                string secondFatherID = ""; string secondMotherID = "";
                string thirdFatherID = ""; string thirdMotherID = "";

                // Create list to hold all siblings and relationship type
                List<Person> siblings = new List<Person>();
                List<string> siblingTypes = new List<string>();

                // Get parent IDs if applicable
                if (!string.IsNullOrEmpty(firstFatherIDLabel.Text)) { firstFatherID = firstFatherIDLabel.Text; }
                if (!string.IsNullOrEmpty(firstMotherIDLabel.Text)) { firstMotherID = firstMotherIDLabel.Text; }
                if (!string.IsNullOrEmpty(secondFatherIDLabel.Text)) { secondFatherID = secondFatherIDLabel.Text; }
                if (!string.IsNullOrEmpty(secondMotherIDLabel.Text)) { secondMotherID = secondMotherIDLabel.Text; }
                if (!string.IsNullOrEmpty(thirdFatherIDLabel.Text)) { thirdFatherID = thirdFatherIDLabel.Text; }
                if (!string.IsNullOrEmpty(thirdMotherIDLabel.Text)) { thirdMotherID = thirdMotherIDLabel.Text; }

                // Compare the parents of everybody in the tree to the highlighted person
                foreach (Person person in people)
                {
                    // Make sure person isn't being compared to themselves
                    if (person.ID != idLabel.Text)
                    {
                        // Full siblings (share both parents)
                        if (person.FirstFatherID == firstFatherID && person.FirstMotherID == firstMotherID)
                        {
                            siblings.Add(person);
                            siblingTypes.Add("Full");
                        }
                        // Half-siblings (share one parent)
                        else if (person.FirstFatherID == firstFatherID || person.FirstMotherID == firstMotherID)
                        {
                            siblings.Add(person);
                            siblingTypes.Add("Half");
                        }
                    }
                }

                // Add siblings to label if applicable
                if (siblings.Count > 0)
                {
                    for (int i = 0; i < siblings.Count; i++)
                    {
                        siblingsLabel.Text += $"{i + 1}. {siblings[i].Name.FullName}, Relationship: {siblingTypes[i]} \n";
                    }
                }
                // No siblings on tree
                else
                {
                    siblingsLabel.Text = "No siblings found for this individual.";
                }
            }
            // An individual cannot have siblings on Family Echo unless they have parents
            else
            {
                siblingsLabel.Text = "No siblings found for this individual. People with no listed parents cannot have siblings on Family Echo.";
            }
        }


        // Lets the user select a file of properties to load in alongside a family tree
        private void loadPropertiesButton_Click(object sender, EventArgs e)
        {
            // Create a variable to hold the file path and a variable to hold file contents
            string filePath;
            string fileContents = "";

            // Open a File Explorer window for the user to select a file
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.InitialDirectory = "c:\\";
                ofd.RestoreDirectory = true;

                // Get selected file
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    // Get the path of the specified file
                    filePath = ofd.FileName;

                    // Get the file name in the file path
                    string[] folders = filePath.ToString().Split('\\');
                    string fileName = folders[folders.Length - 1];

                    // Display file name
                    propertyFileLabel.Text = fileName;

                    // Get contents of file
                    StreamReader sr = new StreamReader(filePath);
                    fileContents = sr.ReadToEnd();
                    sr.Close();
                }
            }

            // Check if fileContents is empty
            if (!string.IsNullOrEmpty(fileContents))
            {
                // Try to build properties based on the contents of the file
                try
                {
                    // Split file contents into lines
                    string[] lines = fileContents.Split('\n');

                    // Iterate through each line and build properties
                    for (int i = 0; i < lines.Length; i++)
                    {
                        // ID of a new person
                        if (lines[i].Contains("ID:"))
                        {
                            // Split the line; the words array will be used for the entire property
                            string[] words = lines[i].Split(' ');

                            // Iterate through each person and compare IDs
                            // If there is a match, add the property to that person
                            foreach (Person person in people)
                            {
                                // Match found
                                if (person.ID.Contains(words[1].Trim()))
                                {
                                    // Used to get attributes depending on what is provided
                                    string GetAttribute(string[] wordsArray, string location)
                                    {
                                        // Check if the location is more than one word
                                        if (wordsArray.Length > 2)
                                        {
                                            // Build location using all words
                                            for (int j = 1; j < wordsArray.Length; j++)
                                            {
                                                location += words[j] + " ";
                                            }
                                        }
                                        // One word location
                                        else if (wordsArray.Length == 2) location = wordsArray[1];
                                        else location = "";

                                        return location;
                                    }

                                    #region Name & Description

                                    // Move to name line and split contents
                                    i++; i++; words = lines[i].Split(' ');

                                    // Create strings to hold property name & description
                                    string propertyName = "";
                                    string description = "";

                                    // Check if property name is more than one word
                                    if (words.Length > 3)
                                    {
                                        // Build property name using all words
                                        for (int j = 2; j < words.Length; j++)
                                        {
                                            propertyName += words[j] + " ";
                                        }
                                    }
                                    // Property name is only one word
                                    else propertyName = words[2];

                                    // Create a new property using the property name
                                    Property property = new Property(propertyName);

                                    // Description
                                    i++; words = lines[i].Split(' ');
                                    description = GetAttribute(words, description);
                                    property.Description = description;

                                    #endregion

                                    #region Dates

                                    // Move to date type
                                    i++; words = lines[i].Split(' ');

                                    // Check if a date type is provided (should be one)
                                    if (words.Length == 3) property.Date.DateType = words[2];
                                    // Default
                                    else property.Date.DateType = "Exact";

                                    // Check if date type was set to "Range"
                                    if (property.Date.DateType.Contains("Range"))
                                    {
                                        // Fix date type if extra \r is present
                                        property.Date.DateType = "Range";

                                        #region Range Begin

                                        // Move to Range Begin Day
                                        i++; words = lines[i].Split(' ');

                                        // Check if a day was provided
                                        if (words.Length == 3)
                                        {
                                            // Check if a valid integer is used for the day
                                            if (int.TryParse(words[2], out int day)) property.Date.Day = day;
                                        }

                                        // Move to Range Begin Month
                                        i++; words = lines[i].Split(' ');

                                        // Check if a month was provided
                                        if (words.Length == 3)
                                        {
                                            // Check if a valid integer is used for the month
                                            if (int.TryParse(words[2], out int month)) property.Date.Month = month;
                                        }

                                        // Move to Range Begin Year
                                        i++; words = lines[i].Split(' ');

                                        // Check if a year was provided
                                        if (words.Length == 3)
                                        {
                                            // Check if a valid integer is used for the year
                                            if (int.TryParse(words[2], out int year)) property.Date.Year = year;
                                        }

                                        // Move to next line (can be several different things)
                                        i++; words = lines[i].Split(' ');

                                        // Check for a space (means it's not BC)
                                        if (!lines[i].Contains(" "))
                                        {
                                            // Check if the line says "BC"
                                            if (lines[i].Contains("BC"))
                                            {
                                                // Check if there is a valid year
                                                if (property.Date.Year != 0) property.Date.Year = property.Date.Year - (2 * property.Date.Year);
                                            }


                                            // Move to next line
                                            i++; words = lines[i].Split(' ');
                                        }

                                        #endregion

                                        #region Range End

                                        // Check if the line contains "End Day:"
                                        if (lines[i].Contains("End Day:"))
                                        {
                                            // Check if a day was provided
                                            if (words.Length == 3)
                                            {
                                                // Check if a valid integer was provided for the day
                                                if (int.TryParse(words[2], out int day)) property.DateRangeEnd.Day = day;
                                            }
                                        }

                                        // Move to Range End Month
                                        i++; words = lines[i].Split(' ');

                                        // Check if the line contains "End Month:"
                                        if (lines[i].Contains("End Month:"))
                                        {
                                            // Check if a month was provided
                                            if (words.Length == 3)
                                            {
                                                // Check if a valid integer was provided for the month
                                                if (int.TryParse(words[2], out int month)) property.DateRangeEnd.Month = month;
                                            }
                                        }

                                        // Move to Range End Year
                                        i++; words = lines[i].Split(' ');

                                        // Check if the line contains "End Year:"
                                        if (lines[i].Contains("End Year:"))
                                        {
                                            // Check if a year was provided
                                            if (words.Length == 3)
                                            {
                                                // Check if a valid integer was provided for the year
                                                if (int.TryParse(words[2], out int year)) property.DateRangeEnd.Year = year;
                                            }
                                        }

                                        // Move to next line (can be several different things)
                                        i++; words = lines[i].Split(' ');

                                        // Check for a space (means it's not BC)
                                        if (!lines[i].Contains(" "))
                                        {
                                            // Check if the line says "BC"
                                            if (lines[i].Contains("BC"))
                                            {
                                                // Check if there is a valid year
                                                if (property.DateRangeEnd.Year != 0)
                                                {
                                                    property.DateRangeEnd.Year = property.DateRangeEnd.Year - (2 * property.DateRangeEnd.Year);
                                                }
                                            }

                                            // Move to the next line if "BC" was present
                                            i++; words = lines[i].Split(' ');
                                        }

                                        #endregion

                                    }
                                    // Only one date (every other date type)
                                    else
                                    {
                                        // Move to Day
                                        i++; words = lines[i].Split(' ');

                                        // Check if a day was provided
                                        if (words.Length == 2)
                                        {
                                            // Check if a valid integer is used for the day
                                            if (int.TryParse(words[1], out int day)) property.Date.Day = day;
                                        }

                                        // Move to Month
                                        i++; words = lines[i].Split(' ');

                                        // Check if a month was provided
                                        if (words.Length == 2)
                                        {
                                            // Check if a valid integer is used for the month
                                            if (int.TryParse(words[1], out int month)) property.Date.Month = month;
                                        }

                                        // Move to Year
                                        i++; words = lines[i].Split(' ');

                                        // Check if a year was provided
                                        if (words.Length == 2)
                                        {
                                            // Check if a valid integer is used for the year
                                            if (int.TryParse(words[1], out int year)) property.Date.Year = year;
                                        }

                                        // Move to next line (can be several different things)
                                        i++; words = lines[i].Split(' ');

                                        // Check for BC first
                                        if (!lines[i].Contains(" "))
                                        {
                                            // Check if the line says "BC" and nothing else
                                            if (lines[i].Contains("BC"))
                                            {
                                                // Check if there is a valid year
                                                if (property.Date.Year != 0) property.Date.Year = property.Date.Year - (2 * property.Date.Year);
                                            }

                                            // Move to the next line if "BC" was present
                                            i++; words = lines[i].Split(' ');
                                        }
                                    }

                                    #endregion

                                    #region Location

                                    // Create strings to hold the location attributes
                                    string country = ""; string district = "";
                                    string county = ""; string city = "";
                                    string structure = ""; string locationType = "";

                                    // Country
                                    words = lines[i].Split(' ');
                                    country = GetAttribute(words, country);
                                    property.Location.Country = country;

                                    // District
                                    i++; words = lines[i].Split(' ');
                                    district = GetAttribute(words, district);
                                    property.Location.District = district;

                                    // County
                                    i++; words = lines[i].Split(' ');
                                    county = GetAttribute(words, county);
                                    property.Location.County = county;

                                    // City
                                    i++; words = lines[i].Split(' ');
                                    city = GetAttribute(words, city);
                                    property.Location.City = city;

                                    // Structure
                                    i++; words = lines[i].Split(' ');
                                    structure = GetAttribute(words, structure);
                                    property.Location.Structure = structure;

                                    // Location Type
                                    i++;

                                    // Only execute if there are more lines
                                    if (i < lines.Length - 1)
                                    {
                                        words = lines[i].Split(' ');

                                        // Check for Location Type
                                        if (lines[i].Contains("Location Type:"))
                                        {
                                            locationType = GetAttribute(words, locationType);

                                            // Remove "Type:" (error with GetAttribute())
                                            locationType = locationType.Replace("Type:", "");
                                            property.Location.LocationType = locationType;
                                        }

                                        // Check for Present Location line
                                        if (lines[i].Contains("Not Present Location"))
                                        {
                                            property.Location.IsPresentLocation = false;
                                        }
                                        else property.Location.IsPresentLocation = true;

                                        #endregion

                                        // Add property if the divider has been reached; if not, keep going
                                        // until the divider is reached
                                        while (!lines[i].Contains("=============="))
                                        {
                                            if (i < lines.Length - 1) i++;
                                            else break;
                                        }
                                    }

                                    // Add property
                                    person.ExtraProperties.Add(property);

                                    // Display property information if person = curPerson
                                    if (person.ID == curPerson.ID)
                                    {
                                        // Clear previous properties if necessary
                                        propertiesListBox.Items.Clear();

                                        // Add each propery into propertiesListBox
                                        foreach (Property property1 in person.ExtraProperties)
                                        {
                                            propertiesListBox.Items.Add(property1.Name);
                                        }

                                        DisplayPropertyInfo(person, property);
                                    }

                                    break;
                                }
                            }
                        }
                        // Line containing unknown contents
                        else
                        {
                            MessageBox.Show($"Error at Line {i}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            break;
                        }
                    }

                    // Update hasLoadedProperties if applicable
                    if (!hasLoadedProperties) hasLoadedProperties = true;
                }
                // Unable to process file
                catch
                {
                    MessageBox.Show("Unable to load the selected file. Add a property and copy it to the clipboard" +
                        " to see how a property file should be formatted.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


        // Lets the user select a new family tree to examine
        private void newFileButton_Click(object sender, EventArgs e)
        {
            // Update boolean
            isSelectingNewTree = true;

            // Reset fields and exit search mode if applicable
            // NOTE: People is not reset in case a new file isn't selected
            searchMode = false;
            curIndex = 0;
            matches.Clear();
            resultsListBox.Items.Clear();
            resultsLabel.Text = "";
            searchListBox.SelectedIndex = -1;

            // Reload form
            Form1_Load(sender, e);
        }


        // Lets the user save a person's properties to a .txt file of their choice
        private void savePropertiesButton_Click(object sender, EventArgs e)
        {
            // Create variables to hold the file path & contents
            var filePath = string.Empty;
            string fileContents = "";

            // Let the user choose a file
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.InitialDirectory = "c:\\";
                ofd.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                ofd.FilterIndex = 1;
                ofd.RestoreDirectory = true;

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    // Get the path of the specified file
                    filePath = ofd.FileName;

                    // Read file contents
                    StreamReader sr = new StreamReader(filePath);
                    fileContents = sr.ReadToEnd();
                    sr.Close();
                }
            }

            // These hold the two responses to the MessageBoxes
            DialogResult result = new DialogResult();
            DialogResult result1 = new DialogResult();

            // Switches to true if Cancel is hit on first MessageBox
            bool userCancels = false;

            // Use a series of MessageBoxes to safeguard file contents
            while (result1 != DialogResult.Yes)
            {
                // Ask user if they want to overwrite file contents or add properties to the end of the file
                result = MessageBox.Show("This file already has content. Select \"Yes\" to overwrite the contents" +
                    " of this file, \"No\" to add the properties to the end of the file, and \"Cancel\" to" +
                    " leave the file untouched.", "Warning", MessageBoxButtons.YesNoCancel);

                // Add properties depending on user response
                if (result == DialogResult.Yes || result == DialogResult.No)
                {
                    // Clear contents if user selected "Yes"
                    if (result == DialogResult.Yes)
                    {
                        result1 = MessageBox.Show("Are you sure? This cannot be undone.", "Warning",
                            MessageBoxButtons.YesNo);
                    }
                    else result1 = DialogResult.Yes;
                }
                // User cancels
                else
                {
                    userCancels = true;
                    break;
                }
            }

            // Check if the user cancelled the action
            if (!userCancels)
            {
                // Clear file contents if user agrees twice
                if (result == DialogResult.Yes && result1 == DialogResult.Yes)
                {
                    fileContents = "";
                }

                // Create list to hold info for all properties
                List<string> properties = new List<string>();

                // Iterate through each property and save them
                foreach (Property property in curPerson.ExtraProperties)
                {
                    // Create string to hold property info
                    string data = GetProperty(property);

                    // Add divider line
                    data += "==================================================\n";

                    // Add property to list
                    properties.Add(data);
                }

                // Iterate through each property and add information to file
                foreach (string property in properties)
                {
                    fileContents += property;
                }

                // Add contents to file
                FileStream fs = new FileStream(filePath, FileMode.Open);
                StreamWriter sw = new StreamWriter(fs);

                sw.Write(fileContents);
                sw.Close(); fs.Close();
            }
        }


        // Searches for the term specified by the user
        private void searchButton_Click(object sender, EventArgs e)
        {
            // Check if the user entered a search term
            if (!string.IsNullOrEmpty(searchTextBox.Text))
            {
                // Iterate through each item in resultsListBox and compare with the search term
                for (int i = resultsListBox.Items.Count - 1; i >= 0; i--)
                {
                    // Create a string to hold the search term
                    // and an array to hold all words in the search term
                    string searchTerm = searchTextBox.Text.Trim();
                    string[] terms = new string[0];

                    // Check if the search term contains more than one word
                    if (searchTerm.Contains(" "))
                    {
                        terms = searchTerm.Split(' ');
                    }

                    // Check if the search term was successfully split
                    if (terms.Length > 0)
                    {
                        // Compare each result to each term in the list
                        for (int j = 0; j < terms.Length; j++)
                        {
                            // Remove searches that don't contain all words in the search term
                            if (!resultsListBox.Items[i].ToString().Contains(terms[j]))
                            {
                                resultsListBox.Items.RemoveAt(i);
                                break;
                            }
                        }
                    }
                    // Use entire search term if just one word
                    else
                    {
                        // Remove searches that don't contain the search term
                        if (!resultsListBox.Items[i].ToString().Contains(searchTerm))
                        {
                            resultsListBox.Items.RemoveAt(i);
                        }
                    }
                }

                // resultsLabel Plural
                if (resultsListBox.Items.Count > 1)
                {
                    resultsLabel.Text = resultsListBox.Items.Count.ToString("n0") + " results";
                }
                // resultsLabel Singular
                else if (resultsListBox.Items.Count == 1)
                {
                    resultsLabel.Text = resultsListBox.Items.Count.ToString("n0") + " result";
                }
                // No results
                else
                {
                    resultsLabel.Text = "No results";
                }
            }
        }


        // Disables search mode if it was previously enabled
        // and displays info of the first person in system
        private void searchModeButton_Click(object sender, EventArgs e)
        {
            searchMode = false;
            searchModeButton.Enabled = false;
            mainTabControl.SelectedIndex = 0;
            firstButton_Click(sender, e);

            // Toggle navigation buttons
            CheckButtons();
        }


        // Updates the selected property in propertiesListBox on the "Extra Info" Tab
        private void updatePropertyButton_Click(object sender, EventArgs e)
        {
            // Variable to hold all errors
            List<string> errors = new List<string>();

            // Check if a property was selected
            if (propertiesListBox.SelectedIndex != -1)
            {
                // Check if the property has a name
                if (propertyNameTextBox.Text != "")
                {
                    int selectedIndex = propertiesListBox.SelectedIndex;

                    // Create variables to hold property and selected property
                    Property newProperty = new Property(propertyNameTextBox.Text);
                    newProperty.Description = propertyDescriptionTextBox.Text;

                    // Add year
                    if (yearTextBox.Text != "")
                    {
                        if (int.TryParse(yearTextBox.Text, out int year))
                        {
                            // BC
                            if (adBCListBox.SelectedIndex == 1)
                            {
                                year = year - (2 * year);
                            }

                            newProperty.Date.Year = year;
                        }
                        else
                        {
                            errors.Add("Year must be a valid integer.");
                        }
                    }

                    // Add month and day
                    if (monthTextBox.Text != "")
                    {
                        if (int.TryParse(monthTextBox.Text, out int month))
                        {
                            // Check if month is in range
                            if (month > 0 && month < 13)
                            {
                                newProperty.Date.Month = month;

                                // Add day (can only be added if a month is added)
                                if (dayTextBox.Text != "")
                                {
                                    if (int.TryParse(dayTextBox.Text, out int day))
                                    {
                                        bool changedDay = false;

                                        // If day is out of range, set it to min or max based on original value
                                        if (day < 1)
                                        {
                                            day = 1;
                                            changedDay = true;
                                        }

                                        // 31-day months
                                        if (month == 1 || month == 3 || month == 5
                                            || month == 7 || month == 8 || month == 10
                                            || month == 12)
                                        {
                                            if (day > 31)
                                            {
                                                day = 31;
                                                changedDay = true;
                                            }
                                        }
                                        // 30-day months
                                        else if (month == 4 || month == 6 || month == 9
                                            || month == 11)
                                        {
                                            if (day > 30)
                                            {
                                                day = 30;
                                                changedDay = true;
                                            }
                                        }
                                        // February (handles normal & leap years)
                                        else if (month == 2)
                                        {
                                            if (yearTextBox.Text != "" && int.TryParse(yearTextBox.Text, out int year))
                                            {
                                                // Leap year
                                                if (year % 4 == 0)
                                                {
                                                    if (day > 29)
                                                    {
                                                        day = 29;
                                                        changedDay = true;
                                                    }
                                                }
                                                // Normal year
                                                else
                                                {
                                                    if (day > 28)
                                                    {
                                                        day = 28;
                                                        changedDay = true;
                                                    }
                                                }
                                            }
                                        }

                                        if (changedDay == true)
                                        {
                                            errors.Add("Day was changed to a valid number.");
                                        }

                                        newProperty.Date.Day = day;
                                    }
                                    else
                                    {
                                        errors.Add("Day must be a valid integer.");
                                    }
                                }
                            }
                            else
                            {
                                errors.Add("Month must be between 1 and 12.");
                            }
                        }
                        else
                        {
                            errors.Add("Month must be a valid integer.");
                            monthTextBox.Focus(); monthTextBox.SelectAll();
                        }
                    }

                    // Check if newProperty has a date added
                    if (newProperty.Date != null)
                    {
                        // Create a full date if any category is filled
                        if (newProperty.Date.Year != 9999 || newProperty.Date.Month != 0 || newProperty.Date.Day != 0)
                        {
                            newProperty.Date.FullDate = newProperty.Date.GetFullDate(newProperty.Date.Year, newProperty.Date.Month, newProperty.Date.Day);
                        }
                    }
                    // Check if newProperty has a date range end added
                    if (newProperty.DateRangeEnd != null)
                    {
                        // Create a full date if any category is filled
                        if (newProperty.DateRangeEnd.Year == 9999 || newProperty.DateRangeEnd.Month != 0 || newProperty.DateRangeEnd.Day != 0)
                        {
                            newProperty.DateRangeEnd.FullDate = newProperty.DateRangeEnd.GetFullDate(newProperty.DateRangeEnd.Year, newProperty.DateRangeEnd.Month, newProperty.DateRangeEnd.Day);
                        }
                    }

                    // Add Date Type component
                    if (exactRadioButton.Checked == true) newProperty.Date.DateType = "Known";
                    else if (estimateRadioButton.Checked == true) newProperty.Date.DateType = "Estimate";
                    else if (beforeRadioButton.Checked == true) newProperty.Date.DateType = "Before";
                    else if (afterRadioButton.Checked == true) newProperty.Date.DateType = "After";
                    else if (rangeRadioButton.Checked == true) newProperty.Date.DateType = "Range";

                    // Add Range End data if applicable
                    if (newProperty.Date.DateType == "Range")
                    {
                        // Add year and/or month
                        if (yearRangeEndTextBox.Text != "" || monthRangeEndTextBox.Text != "" || dayRangeEndTextBox.Text != "")
                        {
                            // Add year
                            if (yearRangeEndTextBox.Text != "")
                            {
                                // Make sure an integer was entered
                                if (int.TryParse(yearRangeEndTextBox.Text, out int yearRange))
                                {
                                    // BC
                                    if (adBCListBox.SelectedIndex == 1)
                                    {
                                        yearRange = yearRange - (2 * yearRange);
                                    }

                                    newProperty.DateRangeEnd.Year = yearRange;
                                }
                                // Year must be an integer
                                else
                                {
                                    errors.Add("Year Range End must be a valid integer.");
                                }
                            }
                            // Add month
                            if (monthRangeEndTextBox.Text != "")
                            {
                                // Make sure an integer was entered
                                if (int.TryParse(monthRangeEndTextBox.Text, out int monthRange))
                                {
                                    // Check if month is in range
                                    if (monthRange > 0 && monthRange < 13)
                                    {
                                        newProperty.DateRangeEnd.Month = monthRange;

                                        // Add day (can only be added if a month is added)
                                        if (dayRangeEndTextBox.Text != "")
                                        {
                                            // Make sure an integer was entered
                                            if (int.TryParse(dayRangeEndTextBox.Text, out int dayRange))
                                            {
                                                bool changedDay = false;

                                                // If day is out of range, set it to min or max based on original value
                                                if (dayRange < 1)
                                                {
                                                    dayRange = 1;
                                                    changedDay = true;
                                                }

                                                // 31-day months
                                                if (monthRange == 1 || monthRange == 3 || monthRange == 5
                                                    || monthRange == 7 || monthRange == 8 || monthRange == 10
                                                    || monthRange == 12)
                                                {
                                                    // Set dayRange to 31 if over maximum
                                                    if (dayRange > 31)
                                                    {
                                                        dayRange = 31;
                                                        changedDay = true;
                                                    }
                                                }
                                                // 30-day months
                                                else if (monthRange == 4 || monthRange == 6 || monthRange == 9
                                                    || monthRange == 11)
                                                {
                                                    // Set dayRange to 30 if over maximum
                                                    if (dayRange > 30)
                                                    {
                                                        dayRange = 30;
                                                        changedDay = true;
                                                    }
                                                }
                                                // February (handles normal & leap years)
                                                else if (monthRange == 2)
                                                {
                                                    // Check if a year range end was entered and validate the user's input if so
                                                    if (yearRangeEndTextBox.Text != "" && int.TryParse(yearRangeEndTextBox.Text, out int yearRange))
                                                    {
                                                        // Leap year
                                                        if (yearRange % 4 == 0)
                                                        {
                                                            // Set dayRange to 29 if over maximum
                                                            if (dayRange > 29)
                                                            {
                                                                dayRange = 29;
                                                                changedDay = true;
                                                            }
                                                        }
                                                        // Normal year
                                                        else
                                                        {
                                                            // Set dayRange to 28 if over maximum
                                                            if (dayRange > 28)
                                                            {
                                                                dayRange = 28;
                                                                changedDay = true;
                                                            }
                                                        }
                                                    }
                                                }

                                                // Notify the user that the date was adjusted
                                                if (changedDay == true)
                                                {
                                                    errors.Add("Day Range End was changed to a valid number.");
                                                }

                                                newProperty.DateRangeEnd.Day = dayRange;
                                            }
                                            // Day Range End must be an integer
                                            else
                                            {
                                                errors.Add("Day Range End must be a valid integer.");
                                            }
                                        }
                                    }
                                    // Month must be in range
                                    else
                                    {
                                        errors.Add("Month Range End must be between 1 and 12.");
                                    }
                                }
                                // Month Range End must be an integer
                                else
                                {
                                    errors.Add("Month Range End must be a valid integer.");
                                    monthRangeEndTextBox.Focus(); monthRangeEndTextBox.SelectAll();
                                }
                            }
                        }
                        // If the user selects the Range option, a range end date must be provided
                        else
                        {
                            errors.Add("A range end must be provided with the Range option selected.");
                        }
                    }

                    // Get selected property name
                    string selectedProperty = propertiesListBox.SelectedItem.ToString();

                    // Location
                    if (propertyBuildingTextBox.Text != "") newProperty.Location.Structure = propertyBuildingTextBox.Text;
                    if (propertyCityTextBox.Text != "") newProperty.Location.City = propertyCityTextBox.Text;
                    if (propertyCountyTextBox.Text != "") newProperty.Location.County = propertyCountyTextBox.Text;
                    if (propertyDistrictTextBox.Text != "") newProperty.Location.District = propertyDistrictTextBox.Text;
                    if (propertyCountryTextBox.Text != "") newProperty.Location.Country = propertyCountryTextBox.Text;
                    if (locationTypeTextBox.Text != "") newProperty.Location.LocationType = locationTypeTextBox.Text;
                    if (notPresentLocationCheckBox.Checked == true) newProperty.Location.IsPresentLocation = false;
                    else newProperty.Location.IsPresentLocation = true;

                    // Add fullAddress to property
                    newProperty.Location.FullAddress = newProperty.Location.BuildFullAddress();

                    // Find selected property and set it to currentProperty
                    for (int i = 0; i < curPerson.ExtraProperties.Count; i++)
                    {
                        if (curPerson.ExtraProperties[i].Name == selectedProperty)
                        {
                            curPerson.ExtraProperties[i] = newProperty;
                            break;
                        }
                    }

                    // Clear previous property names in propertiesListBox
                    propertiesListBox.Items.Clear();

                    // Update propertiesListBox
                    foreach (Property property in curPerson.ExtraProperties)
                    {
                        propertiesListBox.Items.Add(property.Name);
                    }

                    // Print error list
                    if (errors.Count > 0)
                    {
                        string errorList = "";

                        // Iterate through errors and add each to the errorList
                        for (int i = 0; i < errors.Count; i++)
                        {
                            string newError = "\n" + (i + 1).ToString() + " " + errors[i];

                            errorList += newError;
                        }

                        // Display MessageBox
                        MessageBox.Show(errors.Count.ToString() + " error(s) were found while adding this property.\n" + errorList, "Error", MessageBoxButtons.OK);
                    }

                    // Focus on newly updated property
                    propertiesListBox.SelectedIndex = selectedIndex;
                }
                // Property must have a name
                else
                {
                    MessageBox.Show("A name is required.", "Error", MessageBoxButtons.OK);
                    propertyNameTextBox.Focus(); propertyNameTextBox.SelectAll();
                }
            }
        }

        #endregion

        #region MouseEnter & MouseLeave EventHandlers (StatusStrip)

        #region "Personal" Tab

        private void idLabel_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("ID number in the Family Echo tree");
        }

        private void idLabel_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        private void genderLabel_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("Gender");
        }

        private void genderLabel_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        private void deathLabel_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("Is the person deceased?");
        }

        private void deathLabel_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        private void birthDateTypeLabel_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("Birth date confidence");
        }

        private void birthDateTypeLabel_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        private void deathDateTypeLabel_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("Death date confidence");
        }

        private void deathDateTypeLabel_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        private void burialDateTypeLabel_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("Burial date confidence");
        }

        private void burialDateTypeLabel_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        #endregion

        #region "Partners" Tab

        private void partnerIDLabel_MouseEnter(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "ID number in the Family Echo tree";
        }

        private void partnerIDLabel_MouseLeave(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "";
        }

        private void partnerTypeLabel_MouseEnter(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "Partner's relationship to focal person";
        }

        private void partnerTypeLabel_MouseLeave(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "";
        }

        private void partnerNameLabel_MouseEnter(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "Full name";
        }

        private void partnerNameLabel_MouseLeave(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "";
        }

        private void partnershipDateLabel_MouseEnter(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "Partnership date (i.e. start of relationship, wedding)";
        }

        private void partnershipDateLabel_MouseLeave(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "";
        }

        private void partnershipDateTypeLabel_MouseEnter(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "Partnership date confidence";
        }

        private void partnershipDateTypeLabel_MouseLeave(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "";
        }

        private void exPartnerIDsLabel_MouseEnter(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "All listed ex-partners";
        }

        private void exPartnerIDsLabel_MouseLeave(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "";
        }

        private void extraPartnerIDsLabel_MouseEnter(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "All listed extra partners";
        }

        private void extraPartnerIDsLabel_MouseLeave(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "";
        }

        #endregion

        #region "Parents" Tab

        private void firstMotherIDLabel_MouseEnter(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "ID number in the Family Echo tree";
        }

        private void firstMotherIDLabel_MouseLeave(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "";
        }

        private void firstFatherIDLabel_MouseEnter(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "ID number in the Family Echo tree";
        }

        private void firstFatherIDLabel_MouseLeave(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "";
        }

        private void secondMotherIDLabel_MouseEnter(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "ID number in the Family Echo tree";
        }

        private void secondMotherIDLabel_MouseLeave(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "";
        }

        private void secondFatherIDLabel_MouseEnter(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "ID number in the Family Echo tree";
        }

        private void secondFatherIDLabel_MouseLeave(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "";
        }

        private void thirdMotherIDLabel_MouseEnter(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "ID number in the Family Echo tree";
        }

        private void thirdMotherIDLabel_MouseLeave(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "";
        }

        private void thirdFatherIDLabel_MouseEnter(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "ID number in the Family Echo tree";
        }

        private void thirdFatherIDLabel_MouseLeave(object sender, EventArgs e)
        {
            helpStatusLabel.Text = "";
        }

        #endregion

        #region "Extra Info" Tab

        private void exactRadioButton_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("I am confident that this date is correct.");
        }

        private void exactRadioButton_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        private void estimateRadioButton_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("This event happened around this date.");
        }

        private void estimateRadioButton_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        private void beforeRadioButton_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("This event happened before this date.");
        }

        private void beforeRadioButton_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        private void afterRadioButton_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("This event happened after this date.");
        }

        private void afterRadioButton_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        private void rangeRadioButton_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("This event happened between two dates.");
        }

        private void rangeRadioButton_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        private void propertiesListBox_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("All extra properties for this individual are listed here.");
        }

        private void propertiesListBox_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        private void updatePropertyButton_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("Update this property with new details.");
        }

        private void updatePropertyButton_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        private void copyToClipboardButton_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("Copy this property's information to the clipboard.");
        }

        private void copyToClipboardButton_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        #endregion

        #region "Options" Tab

        private void ancestryButton_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("Create a list of ancestors of this individual.");
        }

        private void ancestryButton_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        private void ancestorsTrackBar_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("Set the number of generations of ancestors to list.");
        }

        private void ancestorsTrackBar_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        private void createBioButton_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("Create a short bio based on information provided in the tree.");
        }

        private void createBioButton_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        private void descendantsButton_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("Create a list of descendants of this individual.");
        }

        private void descendantsButton_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        private void descendantsTrackBar_MouseEnter(object sender, EventArgs e)
        {
            StatusLabelSet("Set the number of generations of descendants to list.");
        }

        private void descendantsTrackBar_MouseLeave(object sender, EventArgs e)
        {
            StatusLabelClear();
        }

        #endregion

        #endregion

        #region Scroll EventHandlers

        // Update ancestryGenerationsLabel when scrolling
        private void ancestorsTrackBar_Scroll(object sender, EventArgs e)
        {
            switch (ancestorsTrackBar.Value)
            {
                case 1:
                    ancestryGenerationsLabel.Text = "Parents";
                    break;
                case 2:
                    ancestryGenerationsLabel.Text = "Grandparents";
                    break;
                case 3:
                    ancestryGenerationsLabel.Text = "Great-Grandparents";
                    break;
                default:
                    ancestryGenerationsLabel.Text = $"{ancestorsTrackBar.Value - 2}x-Great-Grandparents";
                    break;
            }
        }


        // Update descendantGenerationsLabel when scrolling
        private void descendantsTrackBar_Scroll(object sender, EventArgs e)
        {
            switch (descendantsTrackBar.Value)
            {
                case 1:
                    descendantGenerationsLabel.Text = "Children";
                    break;
                case 2:
                    descendantGenerationsLabel.Text = "Grandchildren";
                    break;
                case 3:
                    descendantGenerationsLabel.Text = "Great-Grandchildren";
                    break;
                default:
                    descendantGenerationsLabel.Text = $"{descendantsTrackBar.Value - 2}x-Great-Grandchildren";
                    break;
            }
        }

        #endregion

        #region SelectedIndexChanged EventHandlers

        // Selects a property to edit on the "Extra Info" Tab
        private void propertiesListBox_SelectedIndexChanged(object sender, EventArgs e)
        {

            // Check if a property was selected
            if (propertiesListBox.SelectedIndex != -1)
            {
                // Enable updatePropertyButton if applicable
                updatePropertyButton.Enabled = true;

                // Display information for selected property
                string propertyName = propertiesListBox.SelectedItem.ToString();

                foreach (Property property1 in curPerson.ExtraProperties)
                {
                    if (property1.Name == propertyName)
                    {
                        DisplayPropertyInfo(curPerson, property1);
                        break;
                    }
                }
            }
            else
            {
                updatePropertyButton.Enabled = false;
            }
        }


        // Selects a property to search for and handles sorting the results
        private void searchListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Make sure an item was selected
            if (searchListBox.SelectedIndex != -1)
            {
                // Get the selected index
                int selectedIndex = searchListBox.SelectedIndex;

                // Get the search term
                string attribute = searchListBox.SelectedItem.ToString();

                // Check if the selected index is in the dictionary
                if (propertySelectors.TryGetValue(selectedIndex, out var selector))
                {
                    // Clear previous results (if applicable)
                    resultsListBox.Items.Clear();

                    // Make sure the item doesn't require special handling
                    // List properties need to have their counts specified programmatically
                    if (!attribute.Contains("Partner Count") && !attribute.Contains("IDs")
                        && !attribute.Contains("Living/Deceased"))
                    {
                        // Add matching people into ListBox
                        foreach (Person person in people)
                        {
                            if (selector(person) != null)
                            {
                                if (!resultsListBox.Items.Contains(selector(person).ToString()))
                                {
                                    resultsListBox.Items.Add(selector(person).ToString());
                                }
                            }
                        }

                        // Remove empty values from the ListBox
                        // This is done twice; this is the first time
                        for (int i = resultsListBox.Items.Count - 1; i >= 0; i--)
                        {
                            if (resultsListBox.Items[i].ToString() == null
                                || string.IsNullOrEmpty(resultsListBox.Items[i].ToString()))
                            {
                                resultsListBox.Items.RemoveAt(i);
                            }
                        }

                        #region Sorting Results

                        // Integers use their own sorting system
                        if (attribute.Contains("Year") || attribute.Contains("Month") || attribute.Contains("Day")
                            || attribute.Contains("Find A Grave"))
                        {
                            // The Sorted property only sorts alphabetically, which is counterproductive when
                            // dealing with numbers
                            resultsListBox.Sorted = false;

                            // If the search property is a year
                            if (attribute.Contains("Year"))
                            {
                                // Remove "9999" as that is the default value when the year is unknown
                                // and useless when searching for a specific year
                                for (int i = resultsListBox.Items.Count - 1; i >= 0; i--)
                                {
                                    if (resultsListBox.Items[i].ToString() == "9999")
                                    {
                                        resultsListBox.Items.RemoveAt(i);
                                    }
                                }

                                // Sort years in descending order
                                // NOTE: Sort properties in Lists take numbers into consideration as well, as
                                // opposed to the Sorted property in ListBoxes
                                List<int> years = resultsListBox.Items.Cast<string>().Select(item => int.Parse(item)).Distinct().ToList();
                                years.Sort();
                                years.Reverse();
                                resultsListBox.Items.Clear();

                                // Add sorted years into ListBox
                                foreach (int year in years)
                                {
                                    resultsListBox.Items.Add(year);
                                }

                                // Format negative years as those are Before Christ
                                for (int i = resultsListBox.Items.Count - 1; i >= 0; i--)
                                {
                                    if (int.Parse(resultsListBox.Items[i].ToString()) < 0)
                                    {
                                        string year = Math.Abs(int.Parse(resultsListBox.Items[i].ToString())).ToString();
                                        year += " BC";
                                        resultsListBox.Items[i] = year;
                                    }
                                }
                            }

                            // If the search property is a month or day
                            else if (attribute.Contains("Month") || attribute.Contains("Day"))
                            {
                                // Remove 0 from ListBox
                                for (int i = resultsListBox.Items.Count - 1; i >= 0; i--)
                                {
                                    if (int.Parse(resultsListBox.Items[i].ToString()) == 0)
                                    {
                                        resultsListBox.Items.RemoveAt(i);
                                    }
                                }

                                // Sort numbers in list
                                List<int> numbers = resultsListBox.Items.Cast<string>().Select(item => int.Parse(item)).Distinct().ToList();
                                numbers.Sort();
                                resultsListBox.Items.Clear();

                                // Add sorted numbers into ListBox
                                foreach (int number in numbers)
                                {
                                    resultsListBox.Items.Add(number);
                                }

                                // Convert numbers to months if applicable
                                if (attribute.Contains("Month"))
                                {
                                    for (int i = 0; i < resultsListBox.Items.Count; i++)
                                    {
                                        resultsListBox.Items[i] = GetMonth(i + 1);
                                    }
                                }
                            }

                            // If the search property is a Find A Grave profile ID
                            else
                            {
                                // Sort IDs in descending order
                                // NOTE: Sort properties in Lists take numbers into consideration as well, as
                                // opposed to the Sorted property in ListBoxes
                                List<int> ids = resultsListBox.Items.Cast<string>().Select(item => int.Parse(item)).Distinct().ToList();
                                ids.Sort();
                                resultsListBox.Items.Clear();

                                // Add sorted years into ListBox
                                foreach (int id in ids)
                                {
                                    resultsListBox.Items.Add(id);
                                }
                            }
                        }

                        // Full dates need a special sorting system
                        else if (attribute.Contains("Date") && !attribute.Contains("Type"))
                        {
                            // Remove sort from ListBox if applicable
                            resultsListBox.Sorted = false;

                            // List that will hold dates that will be sorted three times
                            // First by days, then by months, then by years
                            List<string> dates = resultsListBox.Items.Cast<string>().ToList();

                            List<int> days = new List<int>();

                            #region Days

                            foreach (string str in dates)
                            {
                                // Split full dates and find the day value
                                string[] words = str.Split(' ');

                                if (words.Length == 3)
                                {
                                    try
                                    {
                                        string day = words[1].Remove(words[1].Length - 1, 1);

                                        // Convert day to an integer if it is within a valid range
                                        if (Convert.ToInt32(day) > 0 && Convert.ToInt32(day) < 32)
                                        {
                                            days.Add(Convert.ToInt32(day));
                                        }
                                        else
                                        {
                                            days.Add(0);
                                        }
                                    }
                                    catch
                                    {
                                        MessageBox.Show("Unable to convert " + words[1] + " to a day.", "Error", MessageBoxButtons.OK);
                                    }
                                }
                            }

                            // Sort by days while keeping the lists parallel
                            var zipped1 = dates.Zip(days, (date, day) => new { Date = date, Day = day })
                                                    .OrderBy(pair => pair.Day)
                                                    .ToList();

                            dates = zipped1.Select(pair => pair.Date).ToList();
                            days = zipped1.Select(pair => pair.Day).ToList();

                            resultsListBox.Items.Clear();

                            // Add sorted dates list, now organized by days
                            foreach (string date in dates)
                            {
                                resultsListBox.Items.Add(dates);
                            }

                            #endregion

                            #region Months

                            List<int> months = new List<int>();

                            // Add the integer version of each month into the list
                            foreach (string str in dates)
                            {
                                if (str.Contains("January")) months.Add(1);
                                else if (str.Contains("February")) months.Add(2);
                                else if (str.Contains("March")) months.Add(3);
                                else if (str.Contains("April")) months.Add(4);
                                else if (str.Contains("May")) months.Add(5);
                                else if (str.Contains("June")) months.Add(6);
                                else if (str.Contains("July")) months.Add(7);
                                else if (str.Contains("August")) months.Add(8);
                                else if (str.Contains("September")) months.Add(9);
                                else if (str.Contains("October")) months.Add(10);
                                else if (str.Contains("November")) months.Add(11);
                                else if (str.Contains("December")) months.Add(12);
                                else months.Add(0);
                            }

                            // Sort by months while keeping the lists parallel
                            var zipped2 = dates.Zip(months, (date, month) => new { Date = date, Month = month })
                                                    .OrderBy(pair => pair.Month)
                                                    .ToList();

                            dates = zipped2.Select(pair => pair.Date).ToList();
                            months = zipped2.Select(pair => pair.Month).ToList();

                            resultsListBox.Items.Clear();

                            // Add sorted dates list, now organized by days, then months
                            foreach (string date in dates)
                            {
                                resultsListBox.Items.Add(date);
                            }

                            #endregion

                            #region Years

                            List<int> years = new List<int>();

                            // Sort by years first
                            for (int i = 0; i < dates.Count; i++)
                            {
                                string[] values = dates[i].Split(' ');

                                // Take the year from the date and add it to a parallel list

                                // Test for full AD date
                                if (int.TryParse(values[values.Length - 1], out int year1))
                                {
                                    years.Add(int.Parse(values[values.Length - 1]));
                                }
                                // Test for full BC date
                                else if (int.TryParse(values[values.Length - 2], out int year2))
                                {
                                    int newValue = int.Parse(values[values.Length - 2]);
                                    newValue -= 2 * newValue;

                                    years.Add(newValue);
                                }
                                // Dates with unknown years are moved to the bottom of the list
                                else
                                {
                                    years.Add(9999);
                                }
                            }

                            // Sort by years while keeping the lists parallel
                            var zipped3 = dates.Zip(years, (date, year) => new { Date = date, Year = year })
                                                .OrderBy(pair => pair.Year)
                                                .ToList();

                            dates = zipped3.Select(pair => pair.Date).ToList();
                            years = zipped3.Select(pair => pair.Year).ToList();

                            dates.Reverse();

                            resultsListBox.Items.Clear();

                            // Add sorted dates list, now organized by days, then months, then years
                            foreach (string date in dates)
                            {
                                resultsListBox.Items.Add(date);
                            }

                            #endregion

                        }

                        // Use for every other scenario
                        else
                        {
                            resultsListBox.Sorted = true;
                        }

                        #endregion

                    }

                    // List properties will show up as (Collection) if left unmodified
                    else if (attribute.Contains("IDs"))
                    {
                        if (attribute == "Ex-Partner IDs")
                        {
                            foreach (Person person in people)
                            {
                                if (person.ExPartnerIDs != null && person.ExPartnerIDs.Count > 0)
                                {
                                    string exPartners = "";

                                    if (person.ExPartnerIDsConverted == false) ConvertIDs(person, person.ExPartnerIDs);

                                    for (int i = 0; i < person.ExPartnerIDs.Count; i++)
                                    {
                                        if (i != person.ExPartnerIDs.Count - 1)
                                        {
                                            exPartners += person.ExPartnerIDs[i] + ", ";
                                        }
                                        else
                                        {
                                            exPartners += person.ExPartnerIDs[i];
                                        }
                                    }

                                    resultsListBox.Items.Add(exPartners);
                                }
                            }
                        }
                        // Extra Partner IDs
                        else
                        {
                            foreach (Person person in people)
                            {
                                if (person.ExtraPartnerIDs != null && person.ExtraPartnerIDs.Count > 0)
                                {
                                    string extraPartners = "";

                                    if (person.ExtraPartnerIDsConverted == false) ConvertIDs(person, person.ExtraPartnerIDs);

                                    for (int i = 0; i < person.ExtraPartnerIDs.Count; i++)
                                    {
                                        if (i != person.ExtraPartnerIDs.Count - 1)
                                        {
                                            extraPartners += person.ExtraPartnerIDs[i] + ", ";
                                        }
                                        else
                                        {
                                            extraPartners += person.ExtraPartnerIDs[i];
                                        }
                                    }

                                    resultsListBox.Items.Add(extraPartners);
                                }
                            }
                        }

                        // Sort the results
                        resultsListBox.Sorted = true;

                        // Remove empty slots
                        for (int i = resultsListBox.Items.Count - 1; i >= 0; i--)
                        {
                            if (resultsListBox.Items[i].ToString() == ", ")
                            {
                                resultsListBox.Items.RemoveAt(i);
                            }
                        }
                    }

                    // Handle Living/Deceased boolean
                    else if (attribute.Contains("Living/Deceased"))
                    {
                        resultsListBox.Items.Add("Deceased");
                        resultsListBox.Items.Add("Living");
                    }

                    // Handle Ex-Partner and Extra Partner Counts
                    else
                    {
                        // List to hold the different numbers of exes or extra partners
                        List<int> counts = new List<int>();

                        // Count ex-partners for each person
                        if (searchListBox.SelectedItem.ToString() == "Ex-Partner Count")
                        {
                            foreach (Person person in people)
                            {
                                if (person.ExPartnerIDs != null)
                                {
                                    int exes = person.ExPartnerIDs.Count;

                                    if (!counts.Contains(exes))
                                    {
                                        counts.Add(exes);
                                    }
                                }
                            }
                        }
                        // Count extra partners for each person
                        else
                        {
                            foreach (Person person in people)
                            {
                                if (person.ExtraPartnerIDs != null)
                                {
                                    int extras = person.ExtraPartnerIDs.Count;

                                    if (!counts.Contains(extras))
                                    {
                                        counts.Add(extras);
                                    }
                                }
                            }
                        }

                        // Sort the numbers and remove the sorting from the ListBox if applicable
                        counts.Sort();
                        resultsListBox.Sorted = false;

                        foreach (int count in counts)
                        {
                            resultsListBox.Items.Add(count);
                        }
                    }

                    // Remove empty values from ListBox
                    // This is done twice; this is the second time
                    for (int i = resultsListBox.Items.Count - 1; i >= 0; i--)
                    {
                        if (resultsListBox.Items[i].ToString() == null
                            || string.IsNullOrEmpty(resultsListBox.Items[i].ToString()))
                        {

                            resultsListBox.Items.RemoveAt(i);
                        }
                    }

                    // resultsLabel
                    if (resultsListBox.Items.Count != 1)
                    {
                        resultsLabel.Text = resultsListBox.Items.Count.ToString("n0") + " results";
                    }
                    else
                    {
                        resultsLabel.Text = resultsListBox.Items.Count.ToString("n0") + " result";
                    }

                    // Create a boolean to track if the user searched for a location
                    bool isLocation = false;

                    // Check if the search term is a location
                    if (attribute.Contains("Country")) isLocation = true;
                    if (attribute.Contains("District")) isLocation = true;
                    if (attribute.Contains("County")) isLocation = true;
                    if (attribute.Contains("City")) isLocation = true;
                    if (attribute.Contains("Building")) isLocation = true;
                    if (attribute.Contains("Cemetery")) isLocation = true;
                    if (attribute.Contains("Place")) isLocation = true;

                    // Display Warning Label if the search term is a location
                    if (isLocation == true) warningLabel.Visible = true;
                    else warningLabel.Visible = false;
                }
                // Selected index not defined in dictionary
                else
                {
                    MessageBox.Show("The selected property " + attribute + " was not found" +
                        " in the tree.", "Error", MessageBoxButtons.OK);
                }
            }
        }


        // Shows info for all people that match the search term
        private void resultsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Make sure an actual value was selected in the ListBox
            if (resultsListBox.SelectedIndex != -1)
            {
                // Clear previous results
                matches.Clear();

                // Make sure an actual property was selected
                if (propertySelectors.TryGetValue(searchListBox.SelectedIndex, out var selector))
                {
                    // Get the search term
                    string attribute = searchListBox.SelectedItem.ToString();

                    // Cycle through each person in the system
                    // NOTE: The entire list of people has to be cycled through as the results are not
                    // inherently tied to any Person object
                    foreach (Person person in people)
                    {
                        // Create variable to hold the value of the person's attribute
                        // (the one being compared to the resultsListBox item)
                        var selectedValue = selector(person);

                        // Only add people if the selected attribute has a value
                        if (selectedValue != null && !string.IsNullOrEmpty(selectedValue.ToString()))
                        {
                            // This handles the months as they are written out in the list
                            // but represented with numbers in the person objects
                            if (attribute.Contains("Month"))
                            {
                                int month = 0;

                                // Converts all months into numbers to match the Month property
                                // of the Date class
                                if (resultsListBox.SelectedItem.ToString() == "January") month = 1;
                                if (resultsListBox.SelectedItem.ToString() == "February") month = 2;
                                if (resultsListBox.SelectedItem.ToString() == "March") month = 3;
                                if (resultsListBox.SelectedItem.ToString() == "April") month = 4;
                                if (resultsListBox.SelectedItem.ToString() == "May") month = 5;
                                if (resultsListBox.SelectedItem.ToString() == "June") month = 6;
                                if (resultsListBox.SelectedItem.ToString() == "July") month = 7;
                                if (resultsListBox.SelectedItem.ToString() == "August") month = 8;
                                if (resultsListBox.SelectedItem.ToString() == "September") month = 9;
                                if (resultsListBox.SelectedItem.ToString() == "October") month = 10;
                                if (resultsListBox.SelectedItem.ToString() == "November") month = 11;
                                if (resultsListBox.SelectedItem.ToString() == "December") month = 12;

                                // Add person if their Month value matches the converted month
                                if (Convert.ToInt32(selectedValue) == month)
                                {
                                    matches.Add(person);
                                }
                            }

                            // Years need converted from BC to a negative number when applicable
                            else if (attribute.Contains("Year"))
                            {
                                string year = resultsListBox.SelectedItem.ToString();

                                // Convert BC to a negative number
                                if (resultsListBox.SelectedItem.ToString().Contains(" BC"))
                                {
                                    year = year.Replace(" BC", "");
                                    year = "-" + year;
                                }

                                // Add person if their Year value matches the converted year
                                if (selectedValue.ToString() == year)
                                {
                                    matches.Add(person);
                                }
                            }

                            // Dates with years also need converted from BC to a negative number when applicable
                            else if (attribute.Contains("Date"))
                            {
                                string value = resultsListBox.SelectedItem.ToString();

                                // Remove BC
                                if (value.Contains(" BC"))
                                {
                                    value = value.Replace(" BC", "");
                                }

                                // Add person if their Year value matches the converted year
                                if (selectedValue.ToString() == value)
                                {
                                    matches.Add(person);
                                }
                            }

                            // Search by # of Ex-Partners or Extra Partners
                            else if (attribute.Contains("Partner Count"))
                            {
                                if (attribute == "Ex-Partner Count")
                                {
                                    if (person.ExPartnerIDs != null)
                                    {
                                        if (person.ExPartnerIDs.Count.ToString() == resultsListBox.SelectedItem.ToString())
                                        {
                                            matches.Add(person);
                                        }
                                    }
                                }
                                else
                                {
                                    if (person.ExtraPartnerIDs != null)
                                    {
                                        if (person.ExtraPartnerIDs.Count.ToString() == resultsListBox.SelectedItem.ToString())
                                        {
                                            matches.Add(person);
                                        }
                                    }
                                }
                            }

                            // Search with the Living/Deceased boolean
                            else if (attribute.Contains("Living/Deceased"))
                            {
                                // A block in the searchListBox EventHandler will ensure that
                                // the only two options when "Living/Deceased" is clicked are
                                // Living and Deceased
                                if (resultsListBox.SelectedItem.ToString() == "Living")
                                {
                                    if (person.Living == true)
                                    {
                                        matches.Add(person);
                                    }
                                }
                                else
                                {
                                    if (person.Living == false)
                                    {
                                        matches.Add(person);
                                    }
                                }
                            }

                            // All other values match the value in the Person property
                            else
                            {
                                if (selectedValue.ToString() == resultsListBox.SelectedItem.ToString())
                                {
                                    matches.Add(person);
                                }
                            }
                        }
                    }
                }

                // Take user back to the Personal tab and display the information for Person #1
                mainTabControl.SelectedIndex = 0;

                curIndex = 0;
                curPerson = matches[curIndex];

                searchMode = true;
                searchModeButton.Enabled = true;

                // Display information for first match
                if (matches.Count > 0)
                {
                    DisplayInfo(curPerson, matches);
                }
            }
        }

        #endregion

        #region TextChanged EventHandlers

        // Verifies that the day entered is valid
        private void dayRangeEndTextBox_TextChanged(object sender, EventArgs e)
        {
            // Check if a valid integer was entered
            if (int.TryParse(dayRangeEndTextBox.Text, out int day))
            {
                // Set day to 1 if below minimum
                if (day < 1) day = 1;

                // Find maximum if day is over 31
                else if (day > 31)
                {
                    // Check if a valid integer was entered for month
                    if (int.TryParse(monthRangeEndTextBox.Text, out int month))
                    {
                        // 31-day months
                        if (month == 1 || month == 3 || month == 5 || month == 7
                            || month == 8 || month == 10 | month == 12)
                        {
                            // Set day to 31 if above maximum
                            if (day > 31)
                            {
                                day = 31;
                                dayRangeEndTextBox.Text = day.ToString();
                            }
                        }
                        // 30-day months
                        else if (month == 4 || month == 6 || month == 9 || month == 11)
                        {
                            // Set day to 30 if above maximum
                            if (day > 30)
                            {
                                day = 30;
                                dayRangeEndTextBox.Text = day.ToString();
                            }
                        }
                        // February (month will always be set between 1 and 12)
                        // 2 is the only other option
                        else
                        {
                            // Check if a valid integer was entered for year
                            if (int.TryParse(yearRangeEndTextBox.Text, out int year))
                            {
                                // Leap year
                                if (year % 4 == 0)
                                {
                                    // Set day to 29 if above maximum
                                    if (day > 29)
                                    {
                                        day = 29;
                                        dayRangeEndTextBox.Text = day.ToString();
                                    }
                                }
                                // Normal year
                                else
                                {
                                    // Set day to 28 if above maximum
                                    if (day > 28)
                                    {
                                        day = 28;
                                        dayRangeEndTextBox.Text = day.ToString();
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }


        // Verifies that the day entered is valid
        private void dayTextBox_TextChanged(object sender, EventArgs e)
        {
            // Check if a valid integer was entered
            if (int.TryParse(dayTextBox.Text, out int day))
            {
                // Set day to 1 if below minimum
                if (day < 1) day = 1;

                // Find maximum if day is over 28 (shortest month)
                else if (day > 28)
                {
                    // Check if a valid integer was entered for month
                    if (int.TryParse(monthTextBox.Text, out int month))
                    {
                        // 31-day months
                        if (month == 1 || month == 3 || month == 5 || month == 7
                            || month == 8 || month == 10 | month == 12)
                        {
                            // Set day to 31 if above maximum
                            if (day > 31)
                            {
                                day = 31;
                                dayTextBox.Text = day.ToString();
                            }
                        }
                        // 30-day months
                        else if (month == 4 || month == 6 || month == 9 || month == 11)
                        {
                            // Set day to 30 if above maximum
                            if (day > 30)
                            {
                                day = 30;
                                dayTextBox.Text = day.ToString();
                            }
                        }
                        // February (month will always be set between 1 and 12)
                        // 2 is the only other option
                        else if (month == 2)
                        {
                            // Check if a valid integer was entered for year
                            if (int.TryParse(yearTextBox.Text, out int year))
                            {
                                // Leap year
                                if (year % 4 == 0)
                                {
                                    // Set day to 29 if above maximum
                                    if (day > 29)
                                    {
                                        day = 29;
                                        dayTextBox.Text = day.ToString();
                                    }
                                }
                                // Normal year
                                else
                                {
                                    // Set day to 28 if above maximum
                                    if (day > 28)
                                    {
                                        day = 28;
                                        dayTextBox.Text = day.ToString();
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }


        // Verifies that the month entered is between 1 and 12
        private void monthRangeEndTextBox_TextChanged(object sender, EventArgs e)
        {
            // Check if a valid integer was entered for month
            if (int.TryParse(monthRangeEndTextBox.Text, out int month))
            {
                // Set month to a valid number if necessary
                if (month > 13) month = 12;
                if (month < 1) month = 1;

                monthRangeEndTextBox.Text = month.ToString();

                // Update days if necessary
                dayRangeEndTextBox_TextChanged(sender, e);
            }
        }


        // Verifies that the month entered is between 1 and 12
        private void monthTextBox_TextChanged(object sender, EventArgs e)
        {
            // Check if a valid integer was entered for month
            if (int.TryParse(monthTextBox.Text, out int month))
            {
                // Set month to a valid number if necessary
                if (month > 13) month = 12;
                if (month < 1) month = 1;

                monthTextBox.Text = month.ToString();

                // Update days if necessary
                dayTextBox_TextChanged(sender, e);
            }
        }


        // Adjust for leap year if necessary
        private void yearRangeEndTextBox_TextChanged(object sender, EventArgs e)
        {
            monthRangeEndTextBox_TextChanged(sender, e);
        }


        // Adjust for leap year if necessary
        private void yearTextBox_TextChanged(object sender, EventArgs e)
        {
            monthTextBox_TextChanged(sender, e);
        }

        #endregion

        #endregion

        #region Methods

        // Takes a list of partners and creates a formatted string
        // Used in DisplayInfo()
        void BuildPartnerList(Label label, List<string> ids)
        {
            label.Text = "";

            for (int i = 0; i < ids.Count; i++)
            {
                if (i != ids.Count - 1)
                {
                    label.Text += ids[i] + " || ";
                }
                // Last item in list
                else
                {
                    label.Text += ids[i];
                }
            }
        }


        // Used to prevent clicking a navigation button and encountering an
        // IndexOutOfRange exception
        void CheckButtons()
        {
            void AtStart()
            {
                firstButton.Enabled = false;
                previousButton.Enabled = false;
                nextButton.Enabled = true;
                lastButton.Enabled = true;
            }
            void AtEnd()
            {
                firstButton.Enabled = true;
                previousButton.Enabled = true;
                nextButton.Enabled = false;
                lastButton.Enabled = false;
            }
            void InBetween()
            {
                firstButton.Enabled = true;
                previousButton.Enabled = true;
                nextButton.Enabled = true;
                lastButton.Enabled = true;
            }
            void OnePerson()
            {
                firstButton.Enabled = false;
                previousButton.Enabled = false;
                nextButton.Enabled = false;
                lastButton.Enabled = false;
            }

            if (searchMode == false)
            {
                if (curIndex == 0 && people.Count > 1) AtStart();
                else if (curIndex == people.Count - 1 && people.Count != 1) AtEnd();
                else if (people.Count == 1) OnePerson();
                else InBetween();
            }
            else
            {
                if (curIndex == 0 && matches.Count > 1) AtStart();
                else if (curIndex == matches.Count - 1 && matches.Count != 1) AtEnd();
                else if (matches.Count == 1) OnePerson();
                else InBetween();
            }
        }


        // Clears all controls on "Extra Info" Tab
        void ClearPropertyForm()
        {
            propertyNameTextBox.Clear();
            propertyDescriptionTextBox.Clear();
            if (propertiesListBox.Items.Count == 0) deletePropertyButton.Enabled = false;
            else deletePropertyButton.Enabled = true;

            exactRadioButton.Checked = true;

            monthTextBox.Clear();
            dayTextBox.Clear();
            yearTextBox.Clear();
            adBCListBox.SelectedIndex = 0;

            monthRangeEndTextBox.Clear();
            if (monthRangeEndTextBox.Visible) { monthRangeEndTextBox.Visible = false; }
            dayRangeEndTextBox.Clear();
            if (dayRangeEndTextBox.Visible) { dayRangeEndTextBox.Visible = false; }
            yearRangeEndTextBox.Clear();
            if (yearRangeEndTextBox.Visible) { yearRangeEndTextBox.Visible = false; }
            adBCRangeEndListBox.SelectedIndex = 0;
            if (adBCRangeEndListBox.Visible) { adBCRangeEndListBox.Visible = false; }

            propertyBuildingTextBox.Clear();
            propertyCityTextBox.Clear();
            propertyCountyTextBox.Clear();
            propertyDistrictTextBox.Clear();
            propertyCountryTextBox.Clear();
            locationTypeTextBox.Clear();
            notPresentLocationCheckBox.Checked = false;
        }


        // Converts the IDs in lists to the people they represent
        // Used with ExPartnerIDs and ExtraPartnerIDs
        void ConvertIDs(Person person, List<string> list)
        {
            if (list == person.ExPartnerIDs || list == person.ExtraPartnerIDs)
            {
                if (list != null)
                {
                    if (list.Count > 0)
                    {
                        for (int i = 0; i < list.Count; i++)
                        {
                            foreach (Person person1 in people)
                            {
                                if (person1.ID == list[i])
                                {
                                    list[i] = person1.Name.FullName;
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("List in ConvertIDs() must be the ExPartnerIDs or ExtraPartnerIDs property of the" +
                    " person in the other parameter.", "Error", MessageBoxButtons.OK);
            }
        }


        // Displays the person's info in the four tabs
        // Should always be used with curPerson and curIndex
        void DisplayInfo(Person person, List<Person> list)
        {
            if (person != null)
            {

                #region "Personal" Tab

                idLabel.Text = person.ID;
                nameLabel.Text = person.Name.FullName;

                // Build title text (title + first name + last name + suffix)
                if (!string.IsNullOrEmpty(person.Name.FirstNames) && person.Name.FirstNames.Contains(" "))
                {
                    string[] names = person.Name.FirstNames.Split(' ');

                    titleLabel.Text = $"{person.Name.Title} {names[0]} {person.Name.LastName} {person.Name.Suffix}".Trim();
                }
                else if (!string.IsNullOrEmpty(person.Name.FirstNames))
                {
                    titleLabel.Text = $"{person.Name.Title} {person.Name.FirstNames} {person.Name.LastName} {person.Name.Suffix}".Trim();
                }
                else
                {
                    titleLabel.Text = $"{person.Name.FullName}";
                }


                // Default text if name is unknown
                if (string.IsNullOrEmpty(titleLabel.Text)) titleLabel.Text = "Family Tree Reviewer";

                genderLabel.Text = person.Gender;
                birthNameLabel.Text = person.Name.BirthName;
                person.BirthDate.FullDate = person.BirthDate.GetFullDate(
                                                person.BirthDate.Year,
                                                person.BirthDate.Month,
                                                person.BirthDate.Day);
                if (person.BirthDate.FullDate != null)
                {
                    // Check if the date's a range
                    if (person.BirthDateRangeEnd.FullDate == "")
                    {
                        birthDateLabel.Text = person.BirthDate.FullDate;
                    }
                    else
                    {
                        birthDateLabel.Text = "Between " + person.BirthDate.FullDate +
                            " and " + person.BirthDateRangeEnd.FullDate;
                    }
                    if (person.BirthDate.DateType != null)
                    {
                        birthDateTypeLabel.Text = person.BirthDate.DateType;
                    }
                }

                person.DeathDate.FullDate = person.DeathDate.GetFullDate(
                                                person.DeathDate.Year,
                                                person.DeathDate.Month,
                                                person.DeathDate.Day);
                if (person.DeathDate != null)
                {
                    // Check if the date's a range
                    if (person.DeathDateRangeEnd.FullDate == "")
                    {
                        deathDateLabel.Text = person.DeathDate.FullDate;
                    }
                    else
                    {
                        deathDateLabel.Text = "Between " + person.DeathDate.FullDate +
                            " and " + person.DeathDateRangeEnd.FullDate;
                    }
                    if (person.DeathDate.DateType != null)
                    {
                        deathDateTypeLabel.Text = person.DeathDate.DateType;
                    }
                }

                if (person.Living == true)
                {
                    deathLabel.Text = "Living";
                }
                else
                {
                    deathLabel.Text = "Deceased";
                }

                if (person.BurialDate != null)
                {
                    // Check if the date's a range
                    if (person.BurialDateRangeEnd.FullDate == "")
                    {
                        burialDateLabel.Text = person.BurialDate.FullDate;
                    }
                    else
                    {
                        burialDateLabel.Text = "Between " + person.BurialDate.FullDate +
                            " and " + person.BurialDateRangeEnd.FullDate;
                    }
                    if (person.BurialDate.DateType != null)
                    {
                        burialDateTypeLabel.Text = person.BurialDate.DateType;
                    }
                }

                #endregion

                #region "Partners & Children" Tab

                partnerIDLabel.Text = person.PartnerID;
                partnerTypeLabel.Text = person.PartnerType;
                partnerNameLabel.Text = person.PartnerName.FullName;

                person.PartnershipDate.FullDate = person.PartnershipDate.GetFullDate(
                                person.PartnershipDate.Year,
                                person.PartnershipDate.Month,
                                person.PartnershipDate.Day);

                if (person.PartnershipDate != null)
                {
                    // Check if the date's a range
                    if (person.PartnershipDateRangeEnd.FullDate == "")
                    {
                        partnershipDateLabel.Text = person.PartnershipDate.FullDate;
                    }
                    else
                    {
                        partnershipDateLabel.Text = "Between " + person.PartnershipDate.FullDate +
                            " and " + person.PartnershipDateRangeEnd.FullDate;
                    }
                    if (person.PartnershipDate.DateType != null)
                    {
                        partnershipDateLabel.Text = person.PartnershipDate.FullDate;
                    }
                }

                partnershipDateTypeLabel.Text = person.PartnershipDate.DateType;

                exPartnerIDsLabel.Text = "";
                if (person.ExPartnerIDsConverted == false)
                {
                    if (person.ExPartnerIDs != null)
                    {
                        for (int i = 0; i < person.ExPartnerIDs.Count; i++)
                        {
                            // Get name of each person whose ID is listed
                            foreach (Person person1 in people)
                            {
                                // Replace ID with the person's full name
                                if (person1.ID == person.ExPartnerIDs[i])
                                {
                                    person.ExPartnerIDs[i] = person1.Name.FullName;
                                    break;
                                }
                                else
                                {
                                    // Unable to find name in the list of people
                                    if (person1.ID == people[people.Count - 1].ID)
                                    {
                                        person.ExPartnerIDs[i] = "Not Found";
                                    }
                                }
                            }

                            person.ExPartnerIDsConverted = true;

                            // Build list of Ex-Partners
                            BuildPartnerList(exPartnerIDsLabel, person.ExPartnerIDs);
                        }
                    }
                }

                extraPartnerIDsLabel.Text = "";
                if (person.ExtraPartnerIDsConverted == false)
                {
                    if (person.ExtraPartnerIDs != null)
                    {
                        for (int i = 0; i < person.ExtraPartnerIDs.Count; i++)
                        {
                            // Get name of each person whose ID is listed
                            foreach (Person person1 in people)
                            {
                                // Replace ID with the person's full name
                                if (person1.ID == person.ExtraPartnerIDs[i])
                                {
                                    person.ExtraPartnerIDs[i] = person1.Name.FullName;
                                    break;
                                }
                                else
                                {
                                    // Unable to find name in the list of people
                                    if (person1.ID == people[people.Count - 1].ID)
                                    {
                                        person.ExtraPartnerIDs[i] = "Not Found";
                                    }
                                }
                            }

                            person.ExtraPartnerIDsConverted = true;

                            if (i != person.ExtraPartnerIDs.Count - 1)
                            {
                                extraPartnerIDsLabel.Text += person.ExtraPartnerIDs[i] + " || ";
                            }
                            // Last item in list
                            else
                            {
                                extraPartnerIDsLabel.Text += person.ExtraPartnerIDs[i];
                            }
                        }
                    }
                }
                // Build list of Extra Partners
                else
                {
                    BuildPartnerList(extraPartnerIDsLabel, person.ExtraPartnerIDs);
                }

                // Clear childrenLabel (user has to click findChildrenButton for children to appear)
                childrenLabel.Text = "";

                #endregion

                #region "Parents & Siblings" Tab

                firstMotherIDLabel.Text = person.FirstMotherID;
                firstMotherNameLabel.Text = person.FirstMotherName.FullName;
                firstFatherIDLabel.Text = person.FirstFatherID;
                firstFatherNameLabel.Text = person.FirstFatherName.FullName;
                firstParentsTypeLabel.Text = person.FirstParentsType;

                secondMotherIDLabel.Text = person.SecondMotherID;
                secondMotherNameLabel.Text = person.SecondMotherName.FullName;
                secondFatherIDLabel.Text = person.SecondFatherID;
                secondFatherNameLabel.Text = person.SecondFatherName.FullName;
                secondParentsTypeLabel.Text = person.SecondParentsType;

                thirdMotherIDLabel.Text = person.ThirdMotherID;
                thirdMotherNameLabel.Text = person.ThirdMotherName.FullName;
                thirdFatherIDLabel.Text = person.ThirdFatherID;
                thirdFatherNameLabel.Text = person.ThirdFatherName.FullName;
                thirdParentsTypeLabel.Text = person.ThirdParentsType;

                #endregion

                #region "Biography" Tab

                birthplaceLabel.Text = person.BirthLocation.FullAddress;
                deathplaceLabel.Text = person.DeathLocation.FullAddress;
                deathCauseLabel.Text = person.DeathCause;
                burialPlaceLabel.Text = person.BurialLocation.FullAddress;
                professionLabel.Text = person.Profession;
                companyLabel.Text = person.Company;
                interestsLabel.Text = person.Interests;
                activitiesLabel.Text = person.Activities;
                bioNotesLabel.Text = person.BioNotes;

                #endregion

                #region "Extra Info" Tab

                // Clear previous ListBox results
                propertiesListBox.Items.Clear();

                // Extra Info (display info for first property if applicable)
                if (person.ExtraProperties.Count > 0)
                {
                    foreach (Property property in person.ExtraProperties)
                    {
                        propertiesListBox.Items.Add(property.Name);
                    }

                    DisplayPropertyInfo(person, person.ExtraProperties[0]);
                }
                else
                {
                    // Clear property form if there are no extra properties
                    ClearPropertyForm();
                }

                #endregion

                // Labels in top-right corners that show total count
                listLabel1.Text = (curIndex + 1).ToString("n0") + " of " + list.Count.ToString("n0");
                listLabel2.Text = (curIndex + 1).ToString("n0") + " of " + list.Count.ToString("n0");
                listLabel3.Text = (curIndex + 1).ToString("n0") + " of " + list.Count.ToString("n0");
                listLabel4.Text = (curIndex + 1).ToString("n0") + " of " + list.Count.ToString("n0");
                listLabel5.Text = (curIndex + 1).ToString("n0") + " of " + list.Count.ToString("n0");

                // Clear siblingsLabel (user has to click findSiblingsButton for siblings to appear)
                siblingsLabel.Text = "";
            }
            else
            {
                MessageBox.Show("Couldn't show person #" + curIndex, "Error", MessageBoxButtons.OK);

                // Default to first person in list
                curIndex = 0;
                curPerson = list[curIndex];
                DisplayInfo(curPerson, list);
            }

            // Update navigation buttons
            CheckButtons();
        }


        // Displays info for an extra property of a person
        // Used on "Extra Info" Tab
        void DisplayPropertyInfo(Person person, Property property)
        {
            // Check if the person has any extra properties
            if (person.ExtraProperties.Count > 0)
            {
                // Enable "Delete Property" button if applicable
                deletePropertyButton.Enabled = true;

                // Name and Description
                propertyNameTextBox.Text = property.Name;

                if (!string.IsNullOrEmpty(property.Description))
                {
                    propertyDescriptionTextBox.Text = property.Description;
                }
                else
                {
                    propertyDescriptionTextBox.Clear();
                }

                // Date or Date Range Beginning
                if (property.Date != null)
                {
                    // Check if property values are the default values
                    if (property.Date.Month != 0)
                    {
                        monthTextBox.Text = property.Date.Month.ToString();
                    }
                    else { monthTextBox.Clear(); }
                    if (property.Date.Day != 0)
                    {
                        dayTextBox.Text = property.Date.Day.ToString();
                    }
                    else { dayTextBox.Clear(); }
                    if (property.Date.Year != 9999)
                    {
                        // BC year
                        if (property.Date.Year < 0)
                        {
                            adBCListBox.SelectedIndex = 1;
                        }
                        else
                        {
                            // AD
                            adBCListBox.SelectedIndex = 0;
                        }

                        yearTextBox.Text = Math.Abs(property.Date.Year).ToString();
                    }
                    else
                    {
                        yearTextBox.Clear();
                        adBCListBox.SelectedIndex = 0;
                    }

                    // Display Date Confidence (radio buttons)
                    if (!string.IsNullOrEmpty(property.Date.DateType))
                    {
                        if (property.Date.DateType == "Known")
                        {
                            exactRadioButton.Checked = true;
                        }
                        else if (property.Date.DateType == "Estimate")
                        {
                            estimateRadioButton.Checked = true;
                        }
                        else if (property.Date.DateType == "Before")
                        {
                            beforeRadioButton.Checked = true;
                        }
                        else if (property.Date.DateType == "After")
                        {
                            afterRadioButton.Checked = true;
                        }
                        else if (property.Date.DateType == "Range")
                        {
                            rangeRadioButton.Checked = true;
                        }
                        else
                        {
                            // Default
                            exactRadioButton.Checked = true;
                        }
                    }
                    else
                    {
                        // Default
                        exactRadioButton.Checked = true;
                    }
                }

                // No Date property
                else
                {
                    monthTextBox.Clear();
                    dayTextBox.Clear();
                    yearTextBox.Clear();
                    adBCListBox.SelectedIndex = 0;
                }

                // Date Range End
                if (property.Date.DateType == "Range")
                {
                    rangeRadioButton.Checked = true;

                    // Check if property values are the default values
                    if (property.DateRangeEnd.Month != 0)
                    {
                        monthRangeEndTextBox.Text = property.DateRangeEnd.Month.ToString();
                    }
                    else { monthRangeEndTextBox.Clear(); }
                    if (property.DateRangeEnd.Day != 0)
                    {
                        dayRangeEndTextBox.Text = property.DateRangeEnd.Day.ToString();
                    }
                    else { dayRangeEndTextBox.Clear(); }
                    if (property.DateRangeEnd.Year != 9999)
                    {
                        // BC year
                        if (property.DateRangeEnd.Year < 0)
                        {
                            adBCRangeEndListBox.SelectedIndex = 1;
                        }
                        else
                        {
                            // AD
                            adBCRangeEndListBox.SelectedIndex = 0;
                        }

                        yearRangeEndTextBox.Text = Math.Abs(property.DateRangeEnd.Year).ToString();
                    }
                    // No year
                    else
                    {
                        yearRangeEndTextBox.Clear();
                        adBCRangeEndListBox.SelectedIndex = 0;
                    }
                }

                // No Date Range End property
                else
                {
                    RemoveDateRangeControls();
                }

                // Location
                if (property.Location != null)
                {
                    if (!string.IsNullOrEmpty(property.Location.Structure))
                    {
                        propertyBuildingTextBox.Text = property.Location.Structure;
                    }
                    if (!string.IsNullOrEmpty(property.Location.City))
                    {
                        propertyCityTextBox.Text = property.Location.City;
                    }
                    if (!string.IsNullOrEmpty(property.Location.County))
                    {
                        propertyCountyTextBox.Text = property.Location.County;
                    }
                    if (!string.IsNullOrEmpty(property.Location.District))
                    {
                        propertyDistrictTextBox.Text = property.Location.District;
                    }
                    if (!string.IsNullOrEmpty(property.Location.Country))
                    {
                        propertyCountryTextBox.Text = property.Location.Country;
                    }
                    if (!string.IsNullOrEmpty(property.Location.LocationType))
                    {
                        locationTypeTextBox.Text = property.Location.LocationType;
                    }
                    if (property.Location.IsPresentLocation) { notPresentLocationCheckBox.Checked = true; }
                    else { notPresentLocationCheckBox.Checked = false; }
                }
                // No Location
                else
                {
                    propertyBuildingTextBox.Clear();
                    propertyCityTextBox.Clear();
                    propertyCountyTextBox.Clear();
                    propertyDistrictTextBox.Clear();
                    propertyCountryTextBox.Clear();
                    locationTypeTextBox.Clear();
                    notPresentLocationCheckBox.Checked = false;
                }
            }
            else
            {
                // Clear all controls on "Extra Info" Tab (no properties)
                ClearPropertyForm();
            }
        }


        // Converts a month number to the month name (i.e. 1 => January)
        // NOTE: This needs to be its own method. The GetMonth() property
        // in the Date class cannot be used in its scenario because no
        // Person or Date object is being used
        string GetMonth(int month)
        {
            string monthName;

            switch (month)
            {
                case 1:
                    monthName = "January";
                    break;
                case 2:
                    monthName = "February";
                    break;
                case 3:
                    monthName = "March";
                    break;
                case 4:
                    monthName = "April";
                    break;
                case 5:
                    monthName = "May";
                    break;
                case 6:
                    monthName = "June";
                    break;
                case 7:
                    monthName = "July";
                    break;
                case 8:
                    monthName = "August";
                    break;
                case 9:
                    monthName = "September";
                    break;
                case 10:
                    monthName = "October";
                    break;
                case 11:
                    monthName = "November";
                    break;
                case 12:
                    monthName = "December";
                    break;
                default:
                    monthName = "";
                    break;
            }

            return monthName;
        }


        // Gets the person based on the ID of the label clicked
        void GetPerson(Label label)
        {
            // Leave search mode if applicable
            if (searchMode)
            {
                searchMode = false;
                searchModeButton.Enabled = false;
            }

            if (label.Text != "")
            {
                foreach (Person person in people)
                {
                    if (person.ID == label.Text)
                    {
                        curPerson = person;

                        DisplayInfo(curPerson, people);
                        break;
                    }
                }
            }
        }


        // Create a string holding all property info
        string GetProperty(Property property)
        {
            // Create string to hold property info
            string propertyInfo = "";

            // Print person's ID and Full Name
            propertyInfo += "ID: " + curPerson.ID + "\n";
            propertyInfo += "Name: " + curPerson.Name.FullName + "\n";

            // Property characteristics
            propertyInfo += "Property Name: " + property.Name + "\n";
            propertyInfo += "Description: " + property.Description + "\n";

            // Check if property has a date
            if (!string.IsNullOrEmpty(property.Date.FullDate))
            {
                // Add date type
                propertyInfo += "Date Type: " + property.Date.DateType + "\n";

                // Check if Date Range End is active
                if (property.Date.DateType != "Range")
                {
                    propertyInfo += "Day: " + property.Date.Day + "\n";
                    propertyInfo += "Month: " + property.Date.Month + "\n";
                    propertyInfo += "Year: " + property.Date.Year + "\n";

                    // Check if Date is AD or BC; extra line if BC
                    if (adBCListBox.SelectedItem.ToString() == "BC") propertyInfo += "BC\n";
                }
                // Both Date and DateRangeEnd are needed
                else
                {
                    propertyInfo += "Begin Day: " + property.Date.Day + "\n";
                    propertyInfo += "Begin Month: " + property.Date.Month + "\n";
                    propertyInfo += "Begin Year: " + property.Date.Year + "\n";

                    // Check if Date is AD or BC; extra line if BC
                    if (adBCListBox.SelectedItem.ToString() == "BC") propertyInfo += "BC\n";

                    propertyInfo += "End Day: " + property.DateRangeEnd.Day + "\n";
                    propertyInfo += "End Month: " + property.DateRangeEnd.Month + "\n";
                    propertyInfo += "End Year: " + property.DateRangeEnd.Year + "\n";

                    // Check if Date is AD or BC; extra line if BC
                    if (adBCRangeEndListBox.SelectedItem.ToString() == "BC")
                    {
                        propertyInfo += "BC\n";
                    }
                }
            }

            // Add location if applicable
            if (!string.IsNullOrEmpty(property.Location.FullAddress))
            {
                propertyInfo += "Country: " + property.Location.Country + "\n";
                propertyInfo += "District: " + property.Location.District + "\n";
                propertyInfo += "County: " + property.Location.County + "\n";
                propertyInfo += "City: " + property.Location.City + "\n";
                propertyInfo += "Building: " + property.Location.Structure + "\n";
            }
            // Add location type if applicable
            if (!string.IsNullOrEmpty(property.Location.LocationType))
            {
                propertyInfo += "Location Type: " + property.Location.LocationType + "\n";
            }
            // Add present-day disclaimer if applicable
            if (notPresentLocationCheckBox.Checked) propertyInfo += "Not Present Location" + "\n";

            return propertyInfo;
        }


        // Checks if the Date Range controls need to be visible on "Extra Info" Tab
        void RemoveDateRangeControls()
        {
            monthRangeEndTextBox.Visible = false;
            dayRangeEndTextBox.Visible = false;
            yearRangeEndTextBox.Visible = false;
            adBCRangeEndListBox.Visible = false;
            rangeBeginPromptLabel.Visible = false;
            rangeEndPromptLabel.Visible = false;
        }


        // Clears text from status label (used with MouseLeave EventHandlers)
        void StatusLabelClear()
        {
            helpStatusLabel.Text = "";
        }


        // Adds text to status label (used with MouseEnter EventHandlers)
        void StatusLabelSet(string msg)
        {
            helpStatusLabel.Text = msg;
        }

        #endregion

    }
}