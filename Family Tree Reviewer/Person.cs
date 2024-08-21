using DocumentFormat.OpenXml.Drawing.Charts;
using System.Collections.Generic;
using System.Linq;

namespace Family_Tree_Reviewer
{
    internal class Person
    {
        #region Fields

        #region Primary Info
        string _id;                         // Unique identifying value
        Name _name;                         // Holds first names, last name, nickname, title, and suffix
        string _gender;
        bool _living;

        Date _birthDate;                    // Also serves as birthDateRangeStart if needed
        Date _birthDateRangeEnd;            // Only used if _birthDateType is "Range"

        Date _deathDate;                    // Also serves as deathDateRangeStart if needed
        Date _deathDateRangeEnd;            // Only used if _deathDateType is "Range"

        Date _burialDate;                   // Also serves as burialDateRangeStart if needed
        Date _burialDateRangeEnd;           // Only used if _burialDateType is "Range"

        Location _birthLocation;
        Location _deathLocation;
        Location _burialLocation;

        #endregion

        #region Relationships

        #region Partners

        // Primary/Current Partner
        string _partnerID;
        Name _partnerName;
        string _partnerType;

        // Info about relationship (length, status, etc.)
        Date _partnershipDate;
        Date _partnershipDateRangeEnd;

        // Exes/extras
        List<string> _exPartnerIDs;
        List<string> _extraPartnerIDs;
        bool _exPartnerIDsConverted = false;
        bool _extraPartnerIDsConverted = false;

        #endregion

        #region Parents

        // Primary parents (usually biological)
        string _firstMotherID;
        Name _firstMotherName;
        string _firstFatherID;
        Name _firstFatherName;
        string _firstParentsType;

        // Second parents (step, foster, etc.)
        string _secondMotherID;
        Name _secondMotherName;
        string _secondFatherID;
        Name _secondFatherName;
        string _secondParentsType;

        // Third parents (Family Echo doesn't allow any more sets)
        string _thirdMotherID;
        Name _thirdMotherName;
        string _thirdFatherID;
        Name _thirdFatherName;
        string _thirdParentsType;

        #endregion

        #endregion

        #region Contact

        string _email;
        string _website;
        string _blog;
        string _photoSite;
        string _homeTelephone;
        string _workTelephone;
        string _mobile;
        string _skype;
        string _address;
        string _otherContactInfo;

        #endregion

        #region Biography

        string _deathCause;
        string _profession;
        string _company;
        string _activities;
        string _interests;
        string _bioNotes;

        #endregion

        // Links to profiles on other sites
        #region Further Research

        string _findAGrave;
        string _familySearch;
        string _wikiTree;
        string _wikipedia;

        #endregion

        // These properties allow for extra details about a person that someone
        // using this program can add that aren't part of the normal Family Echo table.
        #region Extra Property Slots

        List<Property> _extraProperties;

        #endregion

        #endregion

        #region Properties

        #region Primary Info
        public string ID
        {
            get { return _id; }
            set { _id = value; }
        }

        public Name Name
        {
            get { return _name; }
            set { _name = value; }
        }

        public string Gender
        {
            get { return _gender; }
            set { _gender = value; }
        }

        public bool Living
        {
            get { return _living; }
            set { _living = value; }
        }

        public Date BirthDate
        {
            get { return _birthDate; }
            set { _birthDate = value; }
        }

        public Date BirthDateRangeEnd
        {
            get { return _birthDateRangeEnd; }
            set { _birthDateRangeEnd = value; }
        }

        public Date DeathDate
        {
            get { return _deathDate; }
            set { _deathDate = value; }
        }

        public Date DeathDateRangeEnd
        {
            get { return _deathDateRangeEnd; }
            set { _deathDateRangeEnd = value; }
        }

        public Date BurialDate
        {
            get { return _burialDate; }
            set { _burialDate = value; }
        }

        public Date BurialDateRangeEnd
        {
            get { return _burialDateRangeEnd; }
            set { _burialDateRangeEnd = value; }
        }

        public Location BirthLocation
        {
            get { return _birthLocation; }
            set { _birthLocation = value; }
        }

        public Location DeathLocation
        {
            get { return _deathLocation; }
            set { _deathLocation = value; }
        }

        public Location BurialLocation
        {
            get { return _burialLocation; }
            set { _burialLocation = value; }
        }

        #endregion

        #region Relationships

        #region Partners

        public string PartnerID
        {
            get { return _partnerID; }
            set { _partnerID = value; }
        }

        public Name PartnerName
        {
            get { return _partnerName; }
            set { _partnerName = value; }
        }

        public string PartnerType
        {
            get { return _partnerType; }
            set { _partnerType = value; }
        }

        public Date PartnershipDate
        {
            get { return _partnershipDate; }
            set { _partnershipDate = value; }
        }

        public Date PartnershipDateRangeEnd
        {
            get { return _partnershipDateRangeEnd; }
            set { _partnershipDateRangeEnd = value; }
        }

        public List<string> ExPartnerIDs
        {
            get { return _exPartnerIDs; }
            set { _exPartnerIDs = value; }
        }

        public List<string> ExtraPartnerIDs
        {
            get { return _extraPartnerIDs; }
            set { _extraPartnerIDs = value; }
        }

        public bool ExPartnerIDsConverted
        {
            get
            {
                return _exPartnerIDsConverted;
            }
            set
            {
                _exPartnerIDsConverted = value;
            }
        }

        public bool ExtraPartnerIDsConverted
        {
            get
            {
                return _extraPartnerIDsConverted;
            }
            set
            {
                _extraPartnerIDsConverted = value;
            }
        }

        #endregion

        #region Parents

        public string FirstMotherID
        {
            get { return _firstMotherID; }
            set { _firstMotherID = value; }
        }

        public Name FirstMotherName
        {
            get { return _firstMotherName; }
            set { _firstMotherName = value; }
        }

        public string FirstFatherID
        {
            get { return _firstFatherID; }
            set { _firstFatherID = value; }
        }

        public Name FirstFatherName
        {
            get { return _firstFatherName; }
            set { _firstFatherName = value; }
        }

        public string FirstParentsType
        {
            get { return _firstParentsType; }
            set { _firstParentsType = value; }
        }

        public string SecondMotherID
        {
            get { return _secondMotherID; }
            set { _secondMotherID = value; }
        }

        public Name SecondMotherName
        {
            get { return _secondMotherName; }
            set { _secondMotherName = value; }
        }

        public string SecondFatherID
        {
            get { return _secondFatherID; }
            set { _secondFatherID = value; }
        }

        public Name SecondFatherName
        {
            get { return _secondFatherName; }
            set { _secondFatherName = value; }
        }

        public string SecondParentsType
        {
            get { return _secondParentsType; }
            set { _secondParentsType = value; }
        }

        public string ThirdMotherID
        {
            get { return _thirdMotherID; }
            set { _thirdMotherID = value; }
        }

        public Name ThirdMotherName
        {
            get { return _thirdMotherName; }
            set { _thirdMotherName = value; }
        }

        public string ThirdFatherID
        {
            get { return _thirdFatherID; }
            set { _thirdFatherID = value; }
        }

        public Name ThirdFatherName
        {
            get { return _thirdFatherName; }
            set { _thirdFatherName = value; }
        }

        public string ThirdParentsType
        {
            get { return _thirdParentsType; }
            set { _thirdParentsType = value; }
        }

        #endregion

        #endregion

        #region Contact

        public string Email
        {
            get { return _email; }
            set { _email = value; }
        }

        public string Website
        {
            get { return _website; }
            set { _website = value; }
        }

        public string Blog
        {
            get { return _blog; }
            set { _blog = value; }
        }

        public string PhotoSite
        {
            get { return _photoSite; }
            set { _photoSite = value; }
        }

        public string HomeTelephone
        {
            get { return _homeTelephone; }
            set { _homeTelephone = value; }
        }

        public string WorkTelephone
        {
            get { return _workTelephone; }
            set { _workTelephone = value; }
        }

        public string Mobile
        {
            get { return _mobile; }
            set { _mobile = value; }
        }

        public string Skype
        {
            get { return _skype; }
            set { _skype = value; }
        }

        public string Address
        {
            get { return _address; }
            set { _address = value; }
        }

        public string OtherContactInfo
        {
            get { return _otherContactInfo; }
            set { _otherContactInfo = value; }
        }

        #endregion

        #region Biography

        public string DeathCause
        {
            get { return _deathCause; }
            set { _deathCause = value; }
        }

        public string Profession
        {
            get { return _profession; }
            set { _profession = value; }
        }

        public string Company
        {
            get { return _company; }
            set { _company = value; }
        }

        public string Activities
        {
            get { return _activities; }
            set { _activities = value; }
        }

        public string Interests
        {
            get { return _interests; }
            set { _interests = value; }
        }

        public string BioNotes
        {
            get { return _bioNotes; }
            set { _bioNotes = value; }
        }

        #endregion

        #region Further Research

        public string FindAGrave
        {
            get { return _findAGrave; }
            set { _findAGrave = value; }
        }

        public string FamilySearch
        {
            get { return _familySearch; }
            set { _familySearch = value; }
        }

        public string WikiTree
        {
            get { return _wikiTree; }
            set { _wikiTree = value; }
        }

        public string Wikipedia
        {
            get { return _wikipedia; }
            set { _wikipedia = value; }
        }

        #endregion

        #region Extra Property Slots

        public List<Property> ExtraProperties
        {
            get
            {
                return _extraProperties;
            }
            set
            {
                _extraProperties = value;
            }
        }

        #endregion

        #endregion

        #region Methods

        public void GetResearchLinks()
        {
            // Check if BioNotes contains any text, specifically links
            if (!string.IsNullOrEmpty(BioNotes))
            {
                // Add all words in Bio Notes to an array to be looked through
                string[] words = BioNotes.Split(' ');

                // Get links to other websites
                foreach (string word in words)
                {

                    #region Find A Grave

                    // Find a potential Find A Grave link
                    if (word.Contains("findagrave"))
                    {
                        FindAGrave = word;

                        // Remove extra characters from Find A Grave ID
                        if (!string.IsNullOrEmpty(FindAGrave))
                        {
                            string[] link = FindAGrave.Split('/');

                            foreach (string str in link)
                            {
                                // Find A Grave IDs are numbers
                                if (int.TryParse(str, out int id))
                                {
                                    FindAGrave = str;
                                    break;
                                }
                            }

                            // No extra characters need removed as unlike other sites,
                            // Find A Grave's IDs are in the middle of their link
                        }
                    }

                    #endregion

                    #region FamilySearch

                    // Find a potential FamilySearch link
                    else if (word.Contains("familysearch"))
                    {
                        FamilySearch = word;

                        // Remove extra characters from FamilySearch ID
                        if (!string.IsNullOrEmpty(FamilySearch))
                        {
                            string[] link = FamilySearch.Split('/');

                            foreach (string str in link)
                            {
                                // FamilySearch IDs look like this: ABCD-123
                                if (str.Contains("-"))
                                {
                                    FamilySearch = str;
                                    break;
                                }
                            }

                            // All FamilySearch IDs are 8 characters long (7 letters/numbers and a hyphen)
                            if (FamilySearch.Length > 8)
                            {
                                FamilySearch = FamilySearch.Remove(8);
                            }
                        }
                    }

                    #endregion

                    #region WikiTree

                    // Find a potential WikiTree link
                    else if (word.Contains("wikitree"))
                    {
                        WikiTree = word;

                        // Remove extra characters from WikiTree ID
                        if (!string.IsNullOrEmpty(WikiTree))
                        {
                            string[] link = WikiTree.Split('/');

                            foreach (string str in link)
                            {
                                // WikiTree IDs use this format: lastname-##
                                if (str.Contains("-"))
                                {
                                    WikiTree = str;
                                    break;
                                }
                            }

                            // Split WikiTree by dashes, then edit the last item in the new array
                            string[] sections = WikiTree.Split('-');
                            int lastItem = sections.Length - 1;

                            while (!int.TryParse(sections[lastItem], out int value))
                            {
                                sections[lastItem] = sections[lastItem].Remove(sections[lastItem].Length - 1);
                            }

                            // Clear WikiTree and rebuild it with the edited items
                            WikiTree = "";

                            for (int i = 0; i < sections.Length; i++)
                            {
                                WikiTree += sections[i];
                                if (i < sections.Length - 1)
                                {
                                    WikiTree += "-";
                                }
                            }
                        }
                    }

                    #endregion

                    #region Wikipedia

                    if (word.Contains("wikipedia"))
                    {
                        Wikipedia = word;

                        // Remove extra characters from Wikipedia article name
                        if (!string.IsNullOrEmpty(Wikipedia))
                        {
                            string[] link = Wikipedia.Split('/');

                            // Change Wikipedia link to only include article name
                            Wikipedia = link[link.Length - 1];

                            // This array gets all characters from the last section of the link,
                            // which gets only the article name and allows the program to check for extra characters
                            List<char> chars = link[link.Length - 1].ToCharArray().ToList();

                            // Remove "My" if next link is "My Ancestry"
                            if (chars[chars.Count - 2] == 'M' && chars[chars.Count - 1] == 'y')
                            {
                                chars.RemoveAt(chars.Count - 1);
                                chars.RemoveAt(chars.Count - 1);

                                Wikipedia = "";

                                // Rebuild the Wikipedia link
                                foreach (char c in chars)
                                {
                                    Wikipedia += c;
                                }
                            }
                            // Remove "Geni:" if next link is to Geni.com
                            else if (Wikipedia.Contains("Geni:"))
                            {
                                Wikipedia = Wikipedia.Remove(Wikipedia.Length - 5);
                            }
                            // Remove "FamilySearch": if next link is to Familysearch.com
                            else if (Wikipedia.Contains("FamilySearch:"))
                            {
                                Wikipedia = Wikipedia.Remove(Wikipedia.Length - 13);
                            }
                            // Remove "Serbian:" if next word is a Serbian translation of their name
                            else if (Wikipedia.Contains("Serbian:"))
                            {
                                Wikipedia = Wikipedia.Remove(Wikipedia.Length - 8);
                            }
                            // Remove "Serbian" if next word is a Serbian translation of their name (typo)
                            else if (Wikipedia.Contains("Serbian"))
                            {
                                Wikipedia = Wikipedia.Remove(Wikipedia.Length - 7);
                            }
                            // Remove "Russian:" if next word is a Russian translation of their name
                            else if (Wikipedia.Contains("Russian:"))
                            {
                                Wikipedia = Wikipedia.Remove(Wikipedia.Length - 8);
                            }
                            // Remove "Arabic:" if next word is an Arabic translation of their name
                            else if (Wikipedia.Contains("Arabic:"))
                            {
                                Wikipedia = Wikipedia.Remove(Wikipedia.Length - 7);
                            }
                            // Remove "Greek:" if next word is a Greek translation of their name
                            else if (Wikipedia.Contains("Greek:"))
                            {
                                Wikipedia = Wikipedia.Remove(Wikipedia.Length - 6);
                            }
                            // Remove "Executed" if the person was executed (SHOULD BE IN CAUSE OF DEATH NOT BIO NOTES)
                            else if (Wikipedia.Contains("Executed"))
                            {
                                Wikipedia = Wikipedia.Remove(Wikipedia.Length - 8);
                            }
                            // Remove "House" if next link is to the royal house's Wikipedia article
                            // (SHOULD BE IN ACTIVITIES NOT BIO NOTES)
                            else if (Wikipedia.Contains("House"))
                            {
                                Wikipedia = Wikipedia.Remove(Wikipedia.Length - 5);
                            }
                        }
                    }

                    #endregion

                }
            }
        }

        #endregion

        // Constructor
        // Doesn't have any required parameters as this program cannot be used to
        // add people to the Meyerhoff Family Tree. Instead, the properties for the
        // Person object are given values from the Excel sheet
        public Person()
        {

        }
    }
}