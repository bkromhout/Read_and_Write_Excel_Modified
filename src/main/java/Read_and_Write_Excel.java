import ij.*;
import ij.measure.ResultsTable;
import ij.plugin.PlugIn;
import ij.plugin.filter.Analyzer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**version 1.1.2*/
public class Read_and_Write_Excel implements PlugIn {
    private enum FileHandlingMode { OPEN_CLOSE, READ_OPEN, WRITE_CLOSE, QUEUE}

    private static final String FILE_HOLDER_KEY = "___EXCEL_FILE_HOLDER___";

    // 0-based row numbers for "significant" rows.
    private static final int DATASET_LABEL_ROW = 0;
    private static final int COLUMN_HEADER_ROW = 1;
    private static final int FIRST_DATA_ROW = 2;

    // Default variables, for if the user doesn't supply overrides for them.
    private static final String DEFAULT_SHEET_NAME = "A";
    private static final String DEFAULT_FILE_PATH = System.getProperty("user.home") + File.separator + "Desktop" +
            File.separator + "Rename me after writing is done.xlsx";

    // Options to use during the run. Populated when parseOptions() is called.
    private String filePath;
    private String sheetName;
    private String dataSetLabel;
    private boolean stackResults = false;
    private boolean noCountColumn = false;
    private FileHandlingMode fileHandlingMode = FileHandlingMode.OPEN_CLOSE;

    // Excel file holder.
    private ExcelHolder fileHolder = null;

    // Variables and main() method for testing in IDEs.
    private static String debugOptions = null;
    private static ResultsTable debugTable = null;
    public static void main(String[] args) {
        // To debug, set a CSV file to open, we'll use that in the ResultsTable for the run.
        debugTable = ResultsTable.open2("/path/to/Values.csv");
        new Read_and_Write_Excel().run("");
        debugOptions = "no_count_column";
        new Read_and_Write_Excel().run("");
        debugOptions = "no_count_column dataset_label=[Test dataset label]";
        new Read_and_Write_Excel().run("");
        debugOptions = "no_count_column dataset_label=[Test dataset label] sheet=[Sheet Name]";
        new Read_and_Write_Excel().run("");
        debugOptions = "dataset_label=[Test dataset label 2]";
        new Read_and_Write_Excel().run("");
        debugOptions = "dataset_label=[Test dataset label 2] sheet=[Sheet Name Longer Than 31 Characters]";
        new Read_and_Write_Excel().run("");
        debugOptions = "no_count_column dataset_label=[Test dataset label] sheet=[Sheet Name] file=[/Users/bkromhout/Desktop/Results File.xlsx]";
        new Read_and_Write_Excel().run("");
        debugOptions = "no_count_column file=[/path/to/Results File.xlsx]";
        new Read_and_Write_Excel().run("");
        debugOptions = "stack_results";
    }

    public void run(String arg) {
        // Fill in the options we'll use during this run.
        parseOptions();

        // Try to get the last file holder.
        fileHolder = (ExcelHolder) IJ.getProperty(FILE_HOLDER_KEY);

        // Handle opening or closing the excel file and stop if that's all we want to do.
        if (fileHandlingMode == FileHandlingMode.WRITE_CLOSE) {
            if (fileHolder == null) IJ.handleException(new IllegalStateException("No excel file open to close."));
            else {
                try {
                    fileHolder.writeOutAndCloseWorkbook();
                } catch (IOException e) {
                    IJ.handleException(e);
                } finally {
                    IJ.setProperty(FILE_HOLDER_KEY, null);
                }
            }
            return;
        } else if (fileHandlingMode == FileHandlingMode.READ_OPEN) {
            if (fileHolder != null) IJ.handleException(new IllegalStateException("There's already an excel file open."));
            else {
                try {
                    fileHolder = new ExcelHolder(filePath);
                    fileHolder.readFileAndOpenWorkbook(sheetName);
                    IJ.setProperty(FILE_HOLDER_KEY, fileHolder);
                } catch (IOException e) {
                    IJ.handleException(e);
                } catch (InvalidFormatException e) {
                    IJ.handleException(e);
                }
            }
            return;
        }

        // Get results, the number of columns and rows, the data, and the headings.
        ResultsTable resultsTable = debugTable == null ? Analyzer.getResultsTable() : debugTable;
        int numColumns = resultsTable.getHeadings().length;
        int numRows = resultsTable.size();
        String[] headers = resultsTable.getHeadings();
        String[][] results = new String[numRows][numColumns];

        // Loop over the results. We could use the getRowAsString(), but we'd just have to split it and parse it again,
        // plus it could mess up if there are any empty cells...this is the most sensible in the end.
        for (int row = 0; row < numRows; row++){
            //Solution for handling empty label-column cells, which otherwise causes the plugin to fail. Replace null label cells with "NaN".
            if (resultsTable.getLabel(row) == null){
                resultsTable.setLabel("null", row);
            }
            //Continue with results loop, as described above.
            for (int col = 0; col < numColumns; col++){
                results[row][col] = resultsTable.getStringValue(headers[col], row);		//Column header reference issue. Passing int for some reason calls hidden values. Passing header string works.
            }
        }
            
        // Figure out which holder to use.
        ExcelHolder holderToUse = fileHolder;
        if (fileHandlingMode == FileHandlingMode.OPEN_CLOSE) {
            try {
                holderToUse = new ExcelHolder(filePath);
                holderToUse.readFileAndOpenWorkbook(sheetName);
            } catch (IOException e) {
                IJ.handleException(e);
            } catch (InvalidFormatException e) {
                IJ.handleException(e);
            }
        }

        if (holderToUse == null) {
            IJ.handleException(new IllegalStateException("No ExcelHolder :("));
            IJ.error("No ExcelHolder :(");
        }
        try {
            Workbook wb = holderToUse.wb;
            // Get sheet in workbook. See method for details.
            Sheet sheet = openSheetInWorkbook(wb);

            // Get the first data row (or create it, if it doesn't exist). We'll use this to determine what column
            // index to start writing our data to this time.
            Row row = sheet.getRow(FIRST_DATA_ROW);
            if (row == null) row = sheet.createRow(FIRST_DATA_ROW);
            // The first column index that we want to write to is either the first column (if there aren't any yet), or
            // the column 2 over from the last one with data in it now (so, we'll end up with a blank column between).
            int firstColIdx = row.getLastCellNum() + 1;
            // Determine the index adjustment based on whether or not we want to have a count column added. We'll use
            // this value (either 0 or 1) for doing things like performing an extra loop iteration, shifting the array
            // index we access a value at, etc.
            int idxAdj = noCountColumn ? 0 : 1;

            // Create a bold font style
            CellStyle boldStyle = wb.createCellStyle();
            Font f = wb.createFont();
            f.setColor(Font.COLOR_NORMAL);
            f.setBold(true);
            boldStyle.setFont(f);

            // Write the column header cells. If stackResults is chosen, then only write the header cells if they do not exist already.
            row = sheet.getRow(COLUMN_HEADER_ROW);
            if (row == null) row = sheet.createRow(COLUMN_HEADER_ROW);
            if (stackResults != true) {
                for (int headerIdx = 0; headerIdx < (headers.length + idxAdj); headerIdx++) {
                    int cellCol = firstColIdx + headerIdx;
                    Cell colHeader = row.getCell(cellCol);
                    if (colHeader == null) colHeader = row.createCell(cellCol);
                    // Make sure we write the "Count" column header to the first column, if we want it.
                    if (cellCol == firstColIdx && !noCountColumn) colHeader.setCellValue("Count");
                    else colHeader.setCellValue(headers[headerIdx - idxAdj]);
                    colHeader.setCellStyle(boldStyle);
                }
            } else {
            	if (sheet.getRow(COLUMN_HEADER_ROW).getCell(0) == null){
                	for (int headerIdx = 0; headerIdx < (headers.length + idxAdj); headerIdx++) {
                		int cellCol = firstColIdx + headerIdx;
                		Cell colHeader = row.getCell(cellCol);
                		if (colHeader == null) colHeader = row.createCell(cellCol);
                		// Make sure we write the "Count" column header to the first column, if we want it.
                		if (cellCol == firstColIdx && !noCountColumn) colHeader.setCellValue("Count");
                		else colHeader.setCellValue(headers[headerIdx - idxAdj]);
                		colHeader.setCellStyle(boldStyle);
            	    }
                } 
            }
            
            // Change the first column index to write data to, depending on stack_results status.
            row = sheet.getRow(FIRST_DATA_ROW);
           	int firstDataRow = FIRST_DATA_ROW;
            if (sheet.getLastRowNum()!=2 && stackResults == true) {
            	int lastColIdx = row.getLastCellNum()-1;
            	int lastRowOfLastColIdx = FIRST_DATA_ROW;
            	Row rowL = sheet.getRow(lastRowOfLastColIdx);
            	Cell testCell = row.getCell(lastColIdx);
            	while (testCell != null){	
            		lastRowOfLastColIdx++;
            		rowL = sheet.getRow(lastRowOfLastColIdx);
            		if (rowL != null){ testCell = rowL.getCell(lastColIdx);}
            				else testCell = null;
            	}
            	firstDataRow = lastRowOfLastColIdx;
            	firstColIdx = lastColIdx - headers.length;
            	if (noCountColumn == true) firstColIdx = firstColIdx + 1;
            }

            // Write dataset label to the first row in the sheet.
            row = sheet.getRow(DATASET_LABEL_ROW);
            if (row == null) row = sheet.createRow(DATASET_LABEL_ROW);
            Cell datasetLabelCell = row.getCell(firstColIdx);
            if (datasetLabelCell == null) datasetLabelCell = row.createCell(firstColIdx);
            datasetLabelCell.setCellValue(dataSetLabel);
            datasetLabelCell.setCellStyle(boldStyle);

            // Write the data (and count number, if necessary) to the sheet. All row/col numbers are 0-based indices.
            // We loop over the actual 2D results array here, and figure out the cell row/col based on that.
            row = sheet.getRow(FIRST_DATA_ROW);
            for (int resultRow = 0; resultRow < results.length; resultRow++) {
                // Figure out the cell row index, then get (or create) a row.
                int cellRow = firstDataRow + resultRow;
                row = sheet.getRow(cellRow);
                if (row == null) row = sheet.createRow(cellRow);
                for (int resultCol = 0; resultCol < (results[resultRow].length + idxAdj); resultCol++) {
                    // Figure out the cell column index, then get (or create) a cell.
                    int cellCol = firstColIdx + resultCol;
                    Cell cell = row.getCell(cellCol);
                    if (cell == null) cell = row.createCell(cellCol);

                    // If this is the first cell column, and we want to write the row count, write that (1-based).
                    if (cellCol == firstColIdx && !noCountColumn) cell.setCellValue(resultRow + 1);
                    else {
                        //Check result datatype and format the cell appropriately before writing
                        //Not a perfect checking method, but will work most of the time without too much overhead
                        if (results[resultRow][resultCol - idxAdj].matches(".*[A-Za-z].*") == true){
                            cell.setCellValue(results[resultRow][resultCol - idxAdj]);
                        } else {
                            cell.setCellType(CellType.NUMERIC);
                    		cell.setCellValue(Double.parseDouble(results[resultRow][resultCol - idxAdj]));
                        }
                    }
                }
                IJ.showProgress(resultRow, results.length + ((int)Math.rint(results.length/100)) );
            }

            // Write the output to a file.
            if (fileHandlingMode == FileHandlingMode.OPEN_CLOSE) holderToUse.writeOutAndCloseWorkbook();
        } catch (IOException e) {
            IJ.handleException(e);
        }
        IJ.showProgress(1);
    }

    /**
     * Opens the sheet called {@link #sheetName} in the given workbook. If {@link #sheetName} is equal to {@link
     * #DEFAULT_SHEET_NAME}, then we ignore it and open the last sheet in the workbook (if there are any existing
     * sheets). If a sheet called {@link #sheetName} doesn't already exist, then we'll create a sheet with the name.
     * @param wb Workbook to open sheet in.
     * @return Sheet called {@link #sheetName}.
     */
    private Sheet openSheetInWorkbook(Workbook wb) {
        if (wb.getNumberOfSheets() == 0) {
            // If there are no sheets, just make a sheet with our current sheetName.
            return wb.createSheet(sheetName);
        } else if (sheetName.equals(DEFAULT_SHEET_NAME)) {
            // If there are sheets, but our sheetName is the default, just return the last sheet.
            return wb.getSheetAt(wb.getNumberOfSheets() - 1);
        } else {
            // We know there are existing sheets, and that sheetName isn't the default. Either open or create the one
            // called sheetName.
            Sheet sheet = wb.getSheet(sheetName);
            // This is a shorthand if-statement: "[boolean expression] ? [return if true] : [return if false];"
            return (sheet != null) ? sheet : wb.createSheet(sheetName);
        }
    }

    /**
     * Parses the options string used to run the plugin and extracts options for it. Uses defaults for any which aren't
     * specified.
     */
    private void parseOptions() {
        // Get the options string we were started using.
        String optionsStr = debugOptions == null ? Macro.getOptions() : debugOptions;
        if (optionsStr == null) optionsStr = "";

        // Figure out the file path to use.
        filePath = Macro.getValue(optionsStr, "file", DEFAULT_FILE_PATH);

        // Figure out the file handling mode to use.
        String temp = Macro.getValue(optionsStr, "file_mode", "");
        if (temp.equals("read_and_open")) fileHandlingMode = FileHandlingMode.READ_OPEN;
        else if (temp.equals("write_and_close")) fileHandlingMode = FileHandlingMode.WRITE_CLOSE;
        else if (temp.equals("queue_write")) fileHandlingMode = FileHandlingMode.QUEUE;
        else fileHandlingMode = FileHandlingMode.OPEN_CLOSE;

        // Figure out the sheet name to use.
        sheetName = Macro.getValue(optionsStr, "sheet", DEFAULT_SHEET_NAME);
        // Make sure the sheet name is safe (we don't want to crash Excel!)
        // See JavaDocs for the method below for what qualifies as an "unsafe" sheet name.
        sheetName = WorkbookUtil.createSafeSheetName(sheetName);

        // Figure out if we want to automatically add a "count" column.
        noCountColumn = optionsStr.contains("no_count_column");

        // Figure out the dataset label to use.
        dataSetLabel = Macro.getValue(optionsStr, "dataset_label", "");
        if (dataSetLabel.equals("")) {
            // If the user didn't provide an image name, try to use the title of the currently active open image.
            // (We'll end up keeping empty string as the label if there are no images open.)
            ImagePlus currImage = WindowManager.getCurrentImage();
            if (currImage != null) dataSetLabel = currImage.getTitle();
        }
        
        // Figure out if we want to place new results data into existing columns
        // Stacking data instead of placing it adjacent to existing data in the designated spreadsheet
        stackResults = optionsStr.contains("stack_results");
    }

    private static class ExcelHolder {
        private File excelFile;
        private FileInputStream fileIn;
        private boolean isValid;
        Workbook wb;

        public ExcelHolder(String filePath) {
            this.excelFile = new File(filePath);
            this.isValid = false;
        }

        void readFileAndOpenWorkbook(String defaultSheetName) throws IOException, InvalidFormatException {
            ensureExcelFileExists(excelFile, defaultSheetName);
            fileIn = new FileInputStream(excelFile);
            wb = WorkbookFactory.create(fileIn);
            isValid = true;
        }

        void writeOutAndCloseWorkbook() throws IOException {
            FileOutputStream fileOut = new FileOutputStream(excelFile);
            wb.write(fileOut);
            if (fileIn != null) fileIn.close();
            fileOut.close();
            isValid = false;
        }

        private void ensureExcelFileExists(File excelFile, String defaultSheetName) throws IOException {
            if (!excelFile.exists()) {
                XSSFWorkbook wb = new XSSFWorkbook();
                wb.createSheet(defaultSheetName);
                FileOutputStream tempOut = new FileOutputStream(excelFile);
                wb.write(tempOut);
            }
        }
    }
}
