# Read and Write Excel *(Modified)*
This is a highly modified version of [Anthony Sinadinos's Read_and_Write_Excel](http://imagej.net/User:ResultsToExcel) ImageJ plugin. It requires at least ImageJ 1.51p and Java 8. The project itself is in IntelliJ IDEA format, but I've tried to maven-ize it too.

The plugin extracts data from the default ImageJ Results Table and adds it to a page in an .xlsx Excel file. Results Table column headers are added automatically too.

By default, the plugin will use (and create, if necessary) a file named "Rename me after writing is done.xlsx" on the desktop, and put the data into a sheet called "A". If writing to a sheet that already has data in it, the new data will be added adjacent to previous data. 

The defaults can be overridden when calling the plugin from an ImageJ macro using one or more of the parameters below.

## Macro Parameters
This version of the Read_and_Write_Excel plugin supports additional features which make it more flexible for usage in ImageJ macros. These are the supported parameters:
* `no_count_column`: Prevents the plugin from adding a "Count" column automatically.
* `file=`: The path to the excel file to use (uses the default desktop file otherwise)
* `sheet=`: Which sheet in the excel file to put the results in
* `dataset_label=`: The label to write in the cell above the data in the excel file
* `file_mode=`: This should be used if you're going to be writing large amounts of data multiple times to the same file, it will let you keep a file open, queue multiple writes to it, then write and close it. This prevents wasting time by having to reopen the whole Excel file every time you wish to write more to it, which can take *very* long if you have a lot of data in the file. Use it by setting it to one of the following:
    * `read_and_open`: Will just open an excel file (the one you specify with `file=`)...make sure you do `write_and_close` when you're done or you'll have problems.
    * `write_and_close`: Will just write everything you've queued with `queue_write`, then close the excel file.
    * `queue_write`: Will queue something to be written to the excel file you've opened previously with `read_and_open`

## Installation
The easiest way to install the plugin is to use Fiji's built-in updater:
1) Go to Help > Update...
2) Click "Manage update sites"
3) Click "Add update site"
4) Give the new update site a name, and use the URL `http://sites.imagej.net/Bkromhout/`
5) You should now see the plugin available in the updater.

If you can't do that for some reason, you should also be able to download the latest release, unzip it, and copy the plugin's JAR to the ImageJ plugins folder, and the JARs in the "jars" folder to ImageJ's "jars" folder.

## Usage Example
```
run("Read and Write Excel", "file_mode=read_and_open file=[/Users/bkromhout/Desktop/Test.xlsx]");
print("Opened file");

// Put lots of data in the results table...

run("Read and Write Excel", "file_mode=queue_write");
print("Wrote 1");
run("Clear Results");

// Put even more data in the results table...

run("Read and Write Excel", "file_mode=queue_write no_count_column dataset_label=[Test dataset label] sheet=[Sheet Name]");
print("Wrote 2");

run("Read and Write Excel", "file_mode=write_and_close");
print("Closed file");
```