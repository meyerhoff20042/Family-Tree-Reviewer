using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Family_Tree_Reviewer
{
    internal class Property
    {
        // Fields
        string _name;
        string _description;
        Date _date;
        Date _dateRangeEnd;
        Location _location;
        
        // Properties
        public string Name
        {
            get
            {
                return _name;
            }
            set
            {
                _name = value;
            }
        }

        public string Description
        {
            get
            {
                return _description;
            }
            set
            {
                _description = value;
            }
        }

        public Date Date
        {
            get
            {
                return _date;
            }
            set
            {
                _date = value;
            }
        }

        public Date DateRangeEnd
        {
            get
            {
                return _dateRangeEnd;
            }
            set
            {
                _dateRangeEnd = value;
            }
        }

        public Location Location
        {
            get
            {
                return _location;
            }
            set
            {
                _location = value;
            }
        }

        public Property(string name)
        {
            _name = name;
            Location = new Location();
            Date = new Date();
            DateRangeEnd = new Date();
        }
    }
}