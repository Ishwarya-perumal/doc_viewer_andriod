# doc_viewer_andriod
# Word to PDF Converter Android App

## Overview

This application allows users to convert Word documents (.doc, .docx) and Excel spreadsheets (.xls, .xlsx) to PDF format. It leverages the Aspose.Words library for Word to PDF conversion and Apache POI for Excel to PDF conversion.
Features:

Convert Word documents to PDF using Aspose.Words library.
Convert Word documents to PDF using iText library.
Convert Excel spreadsheets to PDF using Apache POI library.
View converted PDF files.

## Usage

1.Clone the repository:
2.Build and run the application on your Android device or emulator.
3.Click on the "Choose File" button to select a Word or Excel document.
4.Once a file is selected, various conversion options will be visible.
5.Click on the desired conversion option to convert the document to PDF.
6.Converted PDFs can be viewed using the respective viewer activities.

## Library Used

Android PDF Viewer by barteksc - Version 3.2.0-beta.1
iText PDF Library - Version 5.5.13
Aspose.Words for Java - Version 20.6
Android PDF Viewer by barteksc- Version 3.2.0-beta.1


## Android Configuration

Compile SDK Version: 30
NDK Version: 20.0.5594570



## Update the AndroidManifest.xml file with the necessary permissions and version information:
<uses-permission android:name="android.permission.READ_EXTERNAL_STORAGE" />
<uses-permission android:name="android.permission.WRITE_EXTERNAL_STORAGE" /> 

## Jar File:
POI Shadow - In your (libs/poishadow-all.jar )

# FilePicker
To enable users to choose Word or Excel file from their device storage, this use “openDocumentFromFileManager” function.  
Intent Action: Intent.ACTION_GET_CONTENT is used to open the file picker.
File Type: intent.setType("application/*") specifies that the file picker should filter files based on the MIME type, allowing users to select files of type "application/*" (e.g., Word or Excel documents).

## Excel to PDF using iText

iText is used to create a PDF document (com.itextpdf.text.Document) and associate it with the FileOutputStream.

Each sheet and each cell in the Excel workbook, adding the cell content as paragraphs to the PDF document.

## Apache POI

Apache POI utilizes the Java API for XML processing. By setting these properties, we override the default XML implementation provided by the Java runtime.

## Usage
To ensure proper XML processing in the Apache POI Library, it is recommended to call the ‘initProperties’ Method

## PDF Viewer Component
The PDFView component utilizes the [Android PDF Viewer library by Barteksc].Provides customizable PDF Viewer for application

implementation 'com.github.barteksc:android-pdf-viewer:3.2.0-beta.1' 