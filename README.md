# User Manual: TSV/CSV to Excel Converter

**TSV/CSV Converter** is a powerful GUI application designed to convert plain text tabular data (TSV, CSV, TXT) into Excel format (XLSX) or back into CSV. The software is highly optimized for processing large datasets and supports advanced features such as data splitting, filtering, and pivot table generation.

## Key Features

* **Smart Detection:** Automatically detects file encoding (UTF-8, Windows-1251, etc.) and delimiters (comma, semicolon, tab).
* **Data Splitting:** Split your source file into multiple Excel sheets or entirely separate files based on the values in a specific column.
* **Filtering:** Exclude unnecessary rows before conversion to save time and space.
* **Pivot Tables:** Automatically generate pivot tables with custom data aggregation (Sum, Average, Count, Max, Min).
* **Memory Optimization:** Features a dual-mode engine to prevent out-of-memory errors when processing massive files.

## Step-by-Step Guide

**1. Adding Files**

* Drag and drop `.tsv`, `.csv`, or `.txt` files directly into the file list area, or use the add file button. You can load multiple files for batch processing.

**2. Output Configuration**

* **Output Format:** Choose between `XLSX` (Excel) or `CSV`.
* **Default Path:** Specify the destination folder for the converted files. If left blank, the output will be saved in the same directory as the source file.

**3. Advanced Tools (Optional)**

* **Split by Column:** Select a column from the dropdown menu. A dialog will appear allowing you to select specific values. You can choose to split the data containing these values into **different sheets** within a single Excel workbook, or into **separate files** (by checking the corresponding box). Any unselected values will be grouped under an "Others" category.
* **Filter by Column:** Select a column and check only the values you want to keep. Rows not matching these values will be completely ignored, speeding up conversion and reducing file size.
* **Pivot Table:** Open the pivot table settings to group your data by rows and columns, and add calculated fields (e.g., Sum of sales, Count of records).

**4. Program Settings (⚙️)**

* Change the **Theme** (Light / Dark).
* Enable **Auto-open** to launch the file automatically once conversion is finished.
* Enable **Auto-delete** to remove the original source file after a successful conversion (use with caution!).
* Adjust the **RAM Threshold**: if the number of rows in a file exceeds this limit, the program switches to a strict memory-saving mode. This slows down the conversion slightly but ensures the program doesn't crash on huge files.

**5. Starting the Conversion**

* Click the **Start** button.
* During the process, you can monitor the progress bar, processing speed (rows/sec), and Estimated Time of Arrival (ETA). You can abort the process at any time by clicking **Stop**.

**6. Post-Conversion Actions**

* Click **Open File** to view the result immediately.
* Use the **Delete File** button to quickly remove the generated output if you realize you made a mistake in the settings.
* Click **Export Report** to save a detailed HTML log of the conversion process.
