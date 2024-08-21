using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Meyerhoff_Family_Tree
{
    public partial class FamilyForm : Form
    {
        // Fields
        string _personName = "";        // Name of the central person
        string _family = "";            // Holds ancestry/descendants
        string _familyType = "";        // Either "Ancestry" or "Descendants"
        
        string PersonName
        {
            get { return _personName; }
            set { _personName = value; }
        }

        string Family
        {
            get { return _family; }
            set { _family = value; }
        }

        string FamilyType
        {
            get { return _familyType; }
            set { _familyType = value; }
        }

        // Main person and list of relatives required
        public FamilyForm(string personName, string family, string familyType)
        {
            PersonName = personName;
            Family = family;
            FamilyType = familyType;
            
            InitializeComponent();
        }


        // Set up form
        private void FamilyForm_Load(object sender, EventArgs e)
        {
            // Person's ancestry
            if (_familyType == "Ancestry")
            {
                titleLabel.Text = $"Ancestry of {PersonName}";
                Text = $"Ancestry of {PersonName}";
            }
            // Person's descendants
            else if (_familyType == "Descendance")
            {
                titleLabel.Text = $"Descendants of {PersonName}";
                Text = $"Descendants of {PersonName}";
            }
            // Ancestry and descendants are the only two options
            // Any other value will cause the form to close
            else
            {
                MessageBox.Show("Unknown family type.", "Error", MessageBoxButtons.OK);
                Close();
            }

            // Display family
            descriptionTextBox.Text = Family;
        }


        // Copy information in the descriptionTextBox to the system's clipboard
        private async void clipboardButton_Click(object sender, EventArgs e)
        {
            // Copy text to clipboard
            Clipboard.SetText(descriptionTextBox.Text);

            // Display "Copied to Clipboard!" message temporarily
            clipboardButton.Text = "Copied to Clipboard!";
            clipboardButton.Font = new Font(clipboardButton.Font, FontStyle.Italic);

            await Task.Delay(2000);

            clipboardButton.Text = "Copy to Clipboard";
            clipboardButton.Font = new Font(clipboardButton.Font, FontStyle.Regular);
        }


        // Exit the form
        private void exitButton_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}