using System.Linq;

namespace Family_Tree_Reviewer
{
    internal class Location
    {
        public Location(string fullAddress = "")
        {
            _fullAddress = fullAddress;
        }

        // Fields
        string _structure;
        string _city;
        string _county;
        string _district;
        string _country;
        string _fullAddress;

        // LocationType is used when a location is estimated, unknown, or near somewhere.
        // It can be used to mark if something is NESW of the location or around that area.
        string _locationType;

        // Used when an outdated location is given (i.e. Confederacy or Yugoslavia)
        bool _isPresentLocation;
        
        // Properties
        public string Structure
        {
            get
            {
                return _structure;
            }
            set
            {
                _structure = value;
            }
        }
        
        public string City
        {
            get
            {
                return _city;
            }
            set
            {
                _city = value;
            }
        }

        public string County
        {
            get
            {
                return _county;
            }
            set
            {
                _county = value;
            }
        }

        public string District
        {
            get
            {
                return _district;
            }
            set
            {
                _district = value;
            }
        }

        public string Country
        {
            get
            {
                return _country;
            }
            set
            {
                _country = value;
            }
        }

        public string FullAddress
        {
            get
            {
                return _fullAddress;
            }
            set
            {
                _fullAddress = value;
            }
        }

        public string LocationType
        {
            get
            {
                return _locationType;
            }
            set
            {
                _locationType = value;
            }
        }

        public bool IsPresentLocation
        {
            get
            {
                return _isPresentLocation;
            }
            set
            {
                _isPresentLocation = value;
            }
        }

        // Breaks down a full address into locations
        public void BuildLocations()
        {
            if (!string.IsNullOrEmpty(FullAddress))
            {
                int count = FullAddress.Count(f => f == ',');

                switch (count)
                {
                    // Only country
                    case 0:
                        Country = FullAddress;
                        break;
                    // District and country
                    case 1:
                        string[] zones1 = FullAddress.Split(',');
                        District = zones1[0];
                        Country = zones1[1];
                        break;
                    // County, district and country
                    case 2:
                        string[] zones2 = FullAddress.Split(',');
                        County = zones2[0];
                        District = zones2[1];
                        Country = zones2[2];
                        break;
                    // City, county, district, and country
                    case 3:
                        string[] zones3 = FullAddress.Split(',');
                        City = zones3[0];
                        County = zones3[1];
                        District = zones3[2];
                        Country = zones3[3];
                        break;
                    // Building, city, county, district, and country
                    case 4:
                        string[] zones4 = FullAddress.Split(',');
                        Structure = zones4[0];
                        City = zones4[1];
                        County = zones4[2];
                        District = zones4[3];
                        Country = zones4[4];
                        break;

                        // If there are more than five areas named (i.e. >4 commas),
                        // the method cannot split the address as the areas are undefined.
                }
            }

            // IsPresentLocation
            if (FullAddress.ToLower().Contains("present-day") || FullAddress.ToLower().Contains("now"))
            {
                IsPresentLocation = false;
            }
            else
            {
                IsPresentLocation = true;
            }

            // Trim whitespace from ends of strings if applicable
            if (!string.IsNullOrEmpty(Country)) Country = Country.Trim();
            if (!string.IsNullOrEmpty(District)) District = District.Trim();
            if (!string.IsNullOrEmpty(County)) County = County.Trim();
            if (!string.IsNullOrEmpty(City)) City = City.Trim();
            if (!string.IsNullOrEmpty(Structure)) Structure = Structure.Trim();
            if (!string.IsNullOrEmpty(LocationType)) LocationType = LocationType.Trim();
        }

        // Combines locations to form an address
        public string BuildFullAddress()
        {
            // Build full address
            string fullAddress = "";

            if (Structure != "") fullAddress += Structure;
            if (City != "") fullAddress += ", " + City;
            if (County != "") fullAddress += ", " + County;
            if (District != "") fullAddress += ", " + District;
            if (Country != "") fullAddress += ", " + Country;

            // Remove leading comma if applicable
            if (fullAddress.StartsWith(", "))
            {
                fullAddress = fullAddress.Remove(0, 2);
            }

            return fullAddress;
        }
    }
}