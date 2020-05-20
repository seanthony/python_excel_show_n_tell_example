# python excel example
### 20 may 2020

## dependencies
```
openpyxl
```
documentation for the library can be found here: https://openpyxl.readthedocs.io/en/stable/

# key points for show and tell
* openpyxl can edit/manipulate/template existing excel files
* openpyxl can also write completely new excel files
* can declare formatting in alignment with brand standards
* data is shared between the different python files via locally stored json files (arrays of objects)
    * this uses an if...exists else...write to save on processing time
* using json files for file settings
* using github for version control (hiding sensitive info in the .gitignore)


## note
all data used is fictional

## use
app.py is the primary file which generates the files\
settings for the application are declared in the settings json file. this file is hidden from the repository to prevent the sharing of personal information, but the shape of the file is:
```
{
    "destination folder": "[insert your path to box shared folder on local machine]/show_n_tell/"
}
```
configuring it to a box shared folder will allow collaboration over the reports on box


## included for reference
* scripts used to generate the fake data set
* script used to "stitch" multiple tabular data sets by PID


### future ideas
* potentially using xlsxwriter to manipulate the excel files
* adding a text interface to make decisions
* open to ideas for more!!