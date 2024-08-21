# Family Tree Reviewer

## Purpose
I created this app to provide an advanced search system for a Family Echo tree. The user can find individuals in their tree by any property and add extra properties to any person.

# How to Use
### Selecting a File
The Family Tree Reviewer pulls information from an Excel spreadsheet that lists everybody on a family tree and all of their properties. You can download one of these off of Family Tree by selecting the "CSV (Comma-separated)" option from their Download menu. 
This will take all information from the family tree and export it to a .csv file that can be opened with Excel. 

**NOTE: At the moment, the application is unable to read CSV files. To use a recently downloaded file, create a new Excel workbook and copy all contents from the CSV file to the new file.** In a future update I will address this issue.

## Tabs
### Personal
In the "Personal" tab, you can view a person's vital details, including their full name, birth name (if applicable), and birth/death/burial dates. If the person is living, the last two dates will be blank.

### Partners & Children
The "Partners & Children" tab shows a person's relationships. Their current/primary partner will be shown at the top, with ex-partners and any extra partners listed below. If the person has children, clicking the "Find Children" button will generate a list of
children and their relationship to the person (biological, adopted, step, etc.).

### Parents & Siblings
The "Parents & Siblings" tab shows a person's immediate family. At the moment, Family Echo allows the user to add three sets of parents for an individual, so there are three boxes for each set of parents. Similar to the previous tab, clicking the "Find Siblings" button will generate a list of siblings and their relationship to the person (full, half, step, etc.).

### Biography
The "Biography" tab shows all of the information from Family Echo's biography tab:
1. Birthplace
2. Deathplace
3. Cause of Death
4. Burial Place
5. Profession
6. Company
7. Interests
8. Activities
9. Bio Notes

### Extra Info
The "Extra Info" tab allows the user to add extra properties for a person that are not available on Family Echo. This feature is similar to those provided by larger genealogy sites like Ancestry & FamilySearch, but these extra properties can be whatever you want. 

Each property has to have at least a name, but you can add a description, beginning & end dates, and a location as well. You can create as many properties as you want for an individual. 

To save these properties, select the "Save Properties" button and select the .txt file you want the program to use. To retrieve these properties, you can select "Load Properties" and choose the .txt file.

### Options
The "Options" tab allows the user to print a list of the person's ancestors or descendants. The track bars allow you to specify the number of generations, and a new window will pop up with the generated list. From there, you can copy the list to your clipboard and save it wherever you like.

### Search
The "Search" tab contains everything related to the app's search functionality, which I believe is its most useful feature. Here you can search your entire tree by any property.

When you select a result from the right, you will be able to look through everyone whose selected property contains the search term. To return to viewing everyone in the tree, select the "Leave Search Mode" button.

**NOTE: The last four options are not normally found in Family Echo trees. If you are like me and link an individual's profiles on genealogy sites in their "Bio Notes" section, these last four results will allow you to search by that person's ID on each website.**

The only properties that can not be used in the search algorithm are custom properties created in the "Extra Info" tab. Allowing custom properties to be used as search terms creates the possibility of an inconveniently large list of attributes. At the moment, all of the search terms on the left are from a set list of properties that cannot be changed. I might try to fix this in a future update.
