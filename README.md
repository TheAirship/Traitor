# Traitor! - An experimental Excel-based password cracker for Excel files
*Emphasis on the 'experimental'*

Traitor! was born out of years of telling customers that storing passwords in Excel spreadsheets, password-protected or otherwise, isn't a great idea. After some casual remarks about how it was probably possible to use Excel's VBA backend to crack a password on another Excel file, I decided to see if that was actually feasible.

Traitor! is made available here for testing and tinkering. Some general disclaimers:

* In almost every case, there are better methods to bypass protections on Excel files. This was an experiment, not an effort to create an industry-standard tool.
* Traitor! is a script, and any script has the potential to be malicious. If you're not comfortable using a pre-packaged macro-enabled Excel workbook created by a guy who uses a hand-drawn cartoon seagull as his GitHub avatar, review the source code and use installation methods #1 or #2 listed below.
* Traitor! was created for authorized personal and professional testing. Using it to attack targets without prior mutual consent is illegal. It is the end user's responsibility to obey all applicable local, state, and federal laws. The author(s) assume no liability and are not responsible for any misuse or damage caused by this tool.

## Installation

Traitor! is VBA script, so it doesn't need to be installed, but you do need to have a recent version of Microsoft Excel available. There are three main ways to prepare Traitor! for use:

1. Import VBA script only (if you're very comfortable with VBA editor and scripting)
    * Download the Traitor.bas file, or just clone the entire repo
    * Open the VBA editor in Excel (learn how to do that here: https://docs.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office#macros-and-the-visual-basic-editor)
    * Import the Traitor!.bas file into the VBA editor (File > Import File...)

2. Import VBA script with ControlForm userform (if you're somewhat comfortable with VBA editor and scripting)
    * Download the Traitor.bas, ControlForm.frm, and ControlForm.frx files, or just clone the entire repo
    * Open the VBA editor in Excel (see link above for help)
    * Import the three files mentioned above into the VBA editor (File > Import File...). Importing the ControlForm.frm file will also import the .frx file as long as you haven't renamed it and it's in the same directory.

3. Download pre-loaded Excel workbook (if you have limited or no experience with VBA editor and scripting)
    * Access the repos Releases page at https://github.com/TheAirship/Traitor/releases and download the most current xlsm template

## Use

If you plan to use installation method #1 above, locate the user options section of the script as shown in the screenshot above. Make sure all variables listed in that section have values of the expected type and range. Then execute the "main" sub by selecting anywhere in the that sub's code and pressing the F5 key, or by selecting 'Run' from the VBA editor menu.

![helpimage](https://github.com/TheAirship/Traitor/blob/main/Images/TraitorVariables.png)

If you plan to use installation method #2, you can still adjust the variables in user options section by following the instructions for method #1 above. Alternatively, if you've loaded the ControlForm userform and want to use that instead, right-click on the ControlForm userform object in the VBA editor and select "View Code" locate the Private Sub called "openControlForm" in the code for that userform, select anywhere in that sub and press the F5 key. Make your selections in the ControlForm and click the "Go" button.

![helpimage](https://github.com/TheAirship/Traitor/blob/main/Images/TraitorForm.PNG)

Use of installation method #3 is the easiest. Simply download and open the pre-built Traitor! workbook and click the "Click here to start." button to open the ControlForm user form. Make your selections and click the "Go" button.

## Tips

* Without any add-ins (e.g., Power Pivot), Excel currently supports a maximum of 17,179,869,184 cells per worksheet (https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3). It's not that difficult to generate a wordlist of that size or larger, so if you choose the option to import passwords into the spreadsheet from list with that many words, you should break it into multiple lists and pass each individually.
* VBA doesn't natively support multi-threading. There are a few workarounds, but this project was never intended to be that involved. Be aware that while Traitor! is running, you probably won't be able to use Excel for anything else.
* Traitor! doesn't currently have the ability to pause or stop, though this functionality may be added later. This means that once you initiate an attack, it will run until all password candidates are exhausted or you kill the Excel process. For this reason, you may want to limit the number of password attempts and run the script multiple times.
* Unless you know that the person who locked the spreadsheet or workbook creates terrible passwords, you're almost always going to be better off using the dictionary attack mode. The brute force attack mode does work, but it iterates passwords sequentially and lacks any sort of advanced password format features (e.g., John rules or Hashcat masks).

## About / Licensing

If you've followed the installation and use instructions above along with the guidelines in the script and you still ran into a bug or error, please submit an issue on the Traitor! [Issues](https://github.com/TheAirship/Traitor/issues) page.

Questions, comments, feedback, and feature requests can be submitted to infosec@theairship.cloud. Success stories are welcome too, if that ever happens.
