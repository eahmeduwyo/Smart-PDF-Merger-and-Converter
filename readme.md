# **Smart PDF Merger & Converter**

A Jupyter Notebook tool that intelligently groups, converts, and merges documents (PDFs, DOCX) and images (JPG, JPEG) into single PDF files based on filename prefixes.

## **📖 Background & Motivation**

I built this tool during my Teaching Assistantship to handle the massive influx of student submissions downloaded from **Canvas LMS**.

I was faced with over **1,300 files**, a mix of PDFs, Word documents, and images, that needed to be grouped by student and merged into single documents for reserving for upcoming years. I had already graded those. Doing this manually would have been an impossible, daunting task. I created this notebook to automate the entire workflow, saving countless hours of manual labor.

If you are an educator, TA, or just someone dealing with bulk file organization, I hope this saves you as much time as it saved me\!

## **🚀 How It Works**

The notebook scans a target directory for files. It identifies files that belong together by looking at the text before the first underscore (\_) in the filename.

Example Scenario:  
If your folder contains:

1. projectA\_intro.docx  
2. projectA\_data.pdf  
3. projectA\_chart.jpg  
4. reportB\_final.pdf

The script will:

1. **Group** the first three files under projectA.  
2. **Convert** the .docx and .jpg files to temporary PDFs.  
3. **Merge** them all into a single file: Merged\_Output/projectA.pdf.  
4. Leave reportB as is (or merge it if there were other reportB files).  
5. **Clean up** all temporary files automatically.

## **📋 Prerequisites**

* **Operating System:** Windows (Required for docx2pdf conversion which relies on Microsoft Word).  
* **Software:** Microsoft Word must be installed.  
* **Environment:** Jupyter Lab, Jupyter Notebook, or VS Code with Jupyter extension.  
* **Python:** Python 3.x installed.

## **🛠️ Installation & Setup**

1. **Download:** Clone this repository or download the PDF\_Merger\_and\_Renamer.ipynb file.  
2. Install Libraries:  
   Open the notebook. The first cell contains the command to install the necessary libraries. Run it once:  
   \!pip install pypdf docx2pdf pywin32 Pillow

## **💻 Usage**

1. Open the Notebook:  
   Launch Jupyter Lab or Notebook and open PDF\_Merger\_and\_Renamer.ipynb.  
2. Set Your Folder Path:  
   Locate the variable folder\_path in the code (usually near the bottom) and paste the path to the folder containing your files.  
   \# Example  
   folder\_path \= r'C:\\Users\\YourName\\Downloads\\CanvasSubmissions'

3. Run the Cells:  
   Execute the cells in order. The script will process the files and create a new folder named Merged\_Output inside your source directory containing the final PDFs.

## **⚠️ Common Issues**

Error: pywintypes or CoInitialize  
This usually happens when running Word automation loops inside a notebook.

* **Fix:** The code handles this automatically by initializing COM (pythoncom.CoInitialize()). If the error persists, try restarting the Jupyter Kernel (**Kernel** \> **Restart Kernel**).

**Error: Word requires a save...**

* **Fix:** Ensure no dialog boxes are open in Microsoft Word before running the script.

## **🙏 Acknowledgements**

* **Professor Dhawal Jain:** For providing me with this task and the opportunity to optimize this workflow.  
* **Gemini (Google):** For writing the core automation script and helping streamline the development.