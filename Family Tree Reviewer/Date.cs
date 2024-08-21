using System;
using System.Globalization;

namespace Family_Tree_Reviewer
{
    internal class Date
    {
        // Fields
        int _year;
        int _month;
        int _day;
        string _fullDate;
        string _dateType;

        // Constructor
        public Date(int yyyy = 9999, int mm = 0, int dd = 0, string dateType = "known")
        {
            _dateType = dateType;

            _year = yyyy;
            _month = mm;
            _day = dd;
            _fullDate = GetFullDate(yyyy, mm, dd);
        }

        // Properties
        public int Year
        {
            get
            {
                return _year;
            }
            set
            {
                _year = value;
            }
        }

        public int Month
        {
            get
            {
                return _month;
            }
            set
            {
                if (Month < 0 || Month > 12)
                {
                    throw new IndexOutOfRangeException("Month must be between 0 (unknown) and 12.");
                }

                _month = value;
            }
        }

        public int Day
        {
            get
            {
                return _day;
            }
            set
            {
                if (Day < 0 || Day > 31)
                {
                    throw new IndexOutOfRangeException("Day must be between 0 (unknown) and 31.");
                }

                _day = value;
            }
        }

        public string FullDate
        {
            get
            {
                return _fullDate;
            }
            set
            {
                _fullDate = value;
            }
        }

        public string DateType
        {
            get
            {
                return _dateType;
            }
            set
            {
                _dateType = value;
            }
        }

        public string GetFullDate(int year, int month, int day)
        {
            // Formatted as January 1, 1900
            string fullDate;
            string formattedYear = GetYear(year);
            string formattedMonth = GetMonth(month);

            // Year, day, and month are known
            if (formattedYear != "" && formattedMonth != "" && day.ToString() != "")
            {
                fullDate = formattedMonth + " " + day.ToString() + ", " + formattedYear;
            }
            // Day is unknown
            else if (formattedYear != "" && formattedMonth != "" && day.ToString() == "" || day == 0)
            {
                fullDate = formattedMonth + " " + formattedYear;
            }
            // Only year is known
            else if (formattedYear != "" && formattedMonth == "" && day.ToString() == "")
            {
                fullDate = formattedYear;
            }
            // Year is unknown
            else if (formattedYear == "" && formattedMonth != "" && day.ToString() != "")
            {
                fullDate = formattedMonth + " " + day.ToString();
            }
            // Only month is known
            else if (formattedYear == "" && formattedMonth != "" && day.ToString() == "")
            {
                fullDate = formattedMonth;
            }
            // Other scenarios such as only year and day known (which looks weird)
            else
            {
                fullDate = formattedYear;
            }

            // Replace "0," if applicable
            if (fullDate.Contains("0,"))
            {
                fullDate = fullDate.Replace(" 0, ", " ");
            }

            return fullDate;
        }

        // Used to convert negative years to BC
        public string GetYear(int year)
        {
            if (year != 9999)
            {
                if (year > 0)
                {
                    return year.ToString();
                }
                else
                {
                    return Math.Abs(year).ToString() + " BC";
                }
            }
            else
            {
                return "";
            }
        }

        // Takes a month number and converts it into the month name
        public string GetMonth(int month)
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

    }
}