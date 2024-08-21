namespace Family_Tree_Reviewer
{
    internal class Name
    {
        // Fields
        string _firstNames;
        string _lastName;
        string _birthName;          // Last name at birth (i.e. maiden name)
        string _nickName;
        string _title;
        string _suffix;
        string _fullName;           // Concatenation of all other names

        // Constructors
        public Name()
        {

        }

        // Properties
        public string FirstNames
        {
            get
            {
                return _firstNames;
            }
            set
            {
                _firstNames = value;
            }
        }

        public string LastName
        {
            get
            {
                return _lastName;
            }
            set
            {
                _lastName = value;
            }
        }

        public string BirthName
        {
            get
            {
                return _birthName;
            }
            set
            {
                _birthName = value;
            }
        }

        public string NickName
        {
            get
            {
                return _nickName;
            }
            set 
            { 
                _nickName = value;
            }
        }

        public string Title
        {
            get
            {
                return _title;
            }
            set
            {
                _title = value;
            }
        }

        public string Suffix
        {
            get
            {
                return _suffix;
            }
            set
            {
                _suffix = value;
            }
        }

        public string FullName
        {
            get
            {
                return _fullName;
            }
            set
            {
                _fullName = value;
            }
        }
    }
}