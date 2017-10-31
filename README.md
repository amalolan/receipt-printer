# Receipt Printer
[A Google Sheets Add On](https://chrome.google.com/webstore/detail/receipt-printer/cehbnbbdejohaafojgklmhagflcgeopd?utm_source=permalink) to save a receipt sheet from a spreadsheet as a pdf and email it.

## How does it work?
Receipt Printer uses a template in order to print the receipts. It loops through all the Unique Codes for each receipt one by one , makes a pdf of that receipt, emails them, and finally saves them onto the user's google drive.

So, the template must contain a cell which has that unique code. This unique code is then incremented by 1 in each iteration of the algorithm. To make this work, you would  usually make the template such that the info in that template (First Name, Last Name, Date of Purchase, etc.) depends on that unique code cell to fetch its values (by using VLOOKUP() for example).

Once it prints out that template sheet, for a certain iteration, it emails it to whomever required. And if it needs to be saved into the user's google drive, it does so.

## Inputs

Once you load up the receipt printer with or without mail-merge,  you will see a side bar with 6 input boxes.<br>
NOTE: 'cell' throughout this readme, refers to the cell in the spreadsheet and 'box' or 'input box' refers to the input box on the side bar.<br>

Here is the description of each box:

1.  #### Cell with the amount
  * This takes in a cell (in the format ColumnRow). This cell must contain the amount/price printed using the receipt. **This is an optional input**
  * This is used in the function which converts this amount into words and displays the Amount in Words in the cell from the next input box.
  * If you do not want to convert anything to words or if you already have your own function, enter `null` in this and the next box.
  * Also, currently the amount to words function only supports numbers in the **Indian Numbering System**
  * Example: `A1` , `I5` , `C9`

2. #### Cell to write amount in words
  * This takes in a cell (in the format ColumnRow). This cell's value will be the amount in words from the amount cell given in the previous box. **This is an optional input**
  * A function converts that amount into words and displays it in this box
  * If you do not want to convert anything to words or if you already have your own function, enter `null` in this and the box above.
  * Also, currently the amount to words function only supports numbers in the **Indian Numbering System**
  * Example: `A2` , `I6` , `C10`

3. #### Cell of the unique code
  * This takes in a cell (in the format ColumnRow). This cell must contain the reference unique code. **This is a required input**
  * This receipt printer works by having a reference unique code cell which gets incremented for every receipt. All the values should be fetched from another sheet using that unique code.
  * The receipt printer starts from the current value in this cell and prints till the code entered in the next input box.
  * Please note: **This cell's value should not be empty at the start of the program**
  * Example: `A2` , `I6` , `C10`

4. #### Last entry's unique code
  * This takes in a number. This number is the unique code number up to and including which the add on will print receipts. If you want the whole data to be printed, enter the last unique code in the data sheet.
  **This is a required input**
  * Example: `112`, `55`, `10`

5. #### Destination path
  * This takes in a folder path (in the format main/parent/child). This path is where the pdfs of all the receipts are stored **This is an optional input**
  * The folder in the path is where all receipts are saved. If that path doesn't exist, it will be created.
  * If you do not want to save anything and just email the receipts, enter `null` in this and the next box.
  * **It is recommended that you give a path of a folder instead of leaving it blank if you want to save anything**
  * Example: `Tutorial/Receipts`, `Existing Folder/Hello`, `New Folder`
  * The first folder should be a direct child of the main My Drive folder i.e. when you go to https://drive.google.com/drive/u/0/my-drive it should be present in the folders.
  * If the given folder doesn't exist, it just creates a new folder.

6. #### Structure of filename
  * This takes in cells separated by a comma (in the format ColumnRow, ColumnRow,   ColumnRow ....). This cell must contain the structure of the filename which is how the filename of the pdf will be created. **This is an optional input**
  * Given the cells for the filename, at each run, once the receipt for that unique code is ready, it fetches the values from those cells and adds them together with spaces for the file's name.
  * Example: if I5 contains the unique code and B7 contains the name `I5,B7` would be the input. The filenames would be like so: <br> `ID Name.pdf` or in one case `101 Malolan.pdf`


For a more in-depth tutorial, check out the [Tutorial](https://github.com/amalolan/receipt-printer/blob/master/Tutorial.md) file.

## Limitations

Receipt Printer does suffer from one big limitation: the number of emails sent in a day cannot exceed 100. If you use a CC or BCC, each extra email adds up to that 100. This, unfortunately, is not in our hands as it is a fixed quota set by Google.


## Credits


* Amit Wilson for the [INR()](https://ctrlq.org/code/20098-indian-rupee-lakhs-crores-google-spreadsheet) function
* Amit Agarwal(@labnol) for [paths.gs](https://ctrlq.org/code/19925-google-drive-folder-path)
* @kpgarrod for [pdfMail()](https://gist.github.com/ixhd/3660885) function
