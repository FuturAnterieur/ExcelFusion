# Excel Data Fusion and Processing

## Overview

Here are some Excel macros I wrote during the summer of 2017 to facilitate the data analysis of 
large research projects whose data was stored in Excel spreadsheets. 

There are two main components in the software presented here : the Data Fusion system, which
can fuse similar data found on several different spreadsheets and offers many functions
to normalize the data beforehand (like replacing terms and unifying date formats), and the Data Processing system,
which takes one spreadsheet and performs different processing procedures on it (like replacing terms,
fixing dates, apply functions to columns, and some others). 

In addition, there is an incomplete DBMS-like feature through which data entries can be made to have children
and parents. It then enabled the possibility of filtering data through those relational features.
It was in development at the time I stopped working on the project and, besides, I'd be quite confident it 
is possible to replicate it in a DBMS capable of importing data from Excel, like Access.

## Dependencies

Having a copy of Microsoft Excel capable of running macros and offering the VBA editor interface.
Of course, when opening a workbook with macros, you will need to enable macros when asked if you want
to enable them or not.

## How to use

I wrote a comprehensive manual in French for this thing back when I made it. I haven't bothered translating 
it to English yet (it would take quite some time). The code contains many comments in English, at least.

For now, the French manual is available on Google Drive [here](https://drive.google.com/drive/folders/1vXo_eLz3sMelV-H9FsCzccx3a_n6YPH_?usp=sharing).

What you should know is that prior to running the macros, some macro library references will have to be enabled in Excel.
They should already be enabled in the demo workbook.
Anyhow, the process for enabling them goes as follows :

1- Activate the Developer tab if it is not already there. Go into File -> Settings -> Customize Ribbon, and 
check the Developer tab. 

2- From the Developer tab, select the Visual Basic Editor to open it.

3- In the Visual Basic Editor, open Tools -> References.

4- Make sure that all those libraries are checked :
```
 Visual Basic For Applications
 Microsoft Excel Object Library
 OLE Automation
 Microsoft Office Object Library
 Microsoft Scripting Runtime
 Microsoft Forms Object Library
 Ref Edit Control
```
## Odds and ends 

Some of the VBA modules included in the demo workbook and the repository don't do anything (like TooComplicatedTransChartCode and OldCodeVolume2). Others were made to test situations which were not kept in the demo workbook (like CommandesPourDuMonde, genericDFSysSub, SystemTesting and RandomMacros). Basically, of all the regular (non-class and non-Userform) modules, the only useful ones are MainProgramLaunchers and UtilityFunctions.

But all the other modules, including all the class modules, are useful to the various programs.

The files in the repository were extracted with Rubberduck, which added headers at the top of the code files.
I might remove those headers in the future if people ask me to do so.  


