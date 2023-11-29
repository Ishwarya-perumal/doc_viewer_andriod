package com.example.wordtopdfconverter;

import android.annotation.SuppressLint;
import android.annotation.TargetApi;
import android.content.ContentResolver;
import android.content.Intent;
import android.database.Cursor;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.provider.OpenableColumns;
import android.view.View;
import android.webkit.MimeTypeMap;
import android.widget.Button;
import android.widget.TextView;
import android.widget.Toast;

import androidx.annotation.RequiresApi;
import androidx.appcompat.app.AppCompatActivity;

import com.aspose.words.Document;
import com.itextpdf.text.DocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

@TargetApi(Build.VERSION_CODES.FROYO)
public class MainActivity extends AppCompatActivity {

    private static final int PICK_PDF_FILE = 2;

    private final String storageDir = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS) + File.separator;
    private final String outputPDF = storageDir + "Converted_PDF.pdf";
   

    Button btn_webview;
    Button btn_wordtopdfAspose;
    Button btn_wordtopdf;
    Button btn_exceltopdf;
    private Uri selectedDocumentUri;


    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
  
        btn_webview = findViewById(R.id.WebviewReport);
        btn_wordtopdfAspose = findViewById(R.id.WordtopdfAspose);
        btn_wordtopdf = findViewById(R.id.Wordtopdf);
        btn_exceltopdf = findViewById(R.id.Exceltopdf);

        btn_webview.setVisibility(View.INVISIBLE);
        btn_wordtopdfAspose.setVisibility(View.INVISIBLE);
        btn_wordtopdf.setVisibility(View.INVISIBLE);
        btn_exceltopdf.setVisibility(View.INVISIBLE);


        Button btn_chooseFile = findViewById(R.id.ChooseFile);

        btn_chooseFile.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                openDocumentFromFileManager();
            }
        });

        btn_webview.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                if (selectedDocumentUri != null) {
                    String fileExtension = getFileExtension(selectedDocumentUri);
                    if ("doc".equals(fileExtension) || "docx".equals(fileExtension)) {
                        Intent viewDocxIntent = new Intent(MainActivity.this, DocxViewerActivity.class);
                        viewDocxIntent.setData(selectedDocumentUri);
                        startActivity(viewDocxIntent);
                    } else {
                        Toast.makeText(MainActivity.this, "Selected file is not a Word document", Toast.LENGTH_LONG).show();
                    }
                }
            }
        });

        btn_wordtopdfAspose.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                if (selectedDocumentUri != null) {
                    String fileExtension = getFileExtension(selectedDocumentUri);
                    if ("doc".equals(fileExtension) || "docx".equals(fileExtension)) {
                        try (@SuppressLint({"NewApi", "LocalSuppress"}) InputStream inputStream = getContentResolver().openInputStream(selectedDocumentUri)) {
                            Document doc = new Document(inputStream);
                            doc.save(outputPDF);
                            Toast.makeText(MainActivity.this, "File saved in: " + outputPDF, Toast.LENGTH_LONG).show();
                            showPdfInViewerAspose(outputPDF);
                        } catch (FileNotFoundException e) {
                            e.printStackTrace();
                            Toast.makeText(MainActivity.this, "File not found: " + e.getMessage(), Toast.LENGTH_LONG).show();
                        } catch (IOException e) {
                            e.printStackTrace();
                            Toast.makeText(MainActivity.this, e.getMessage(), Toast.LENGTH_LONG).show();
                        } catch (Exception e) {
                            e.printStackTrace();
                            Toast.makeText(MainActivity.this, e.getMessage(), Toast.LENGTH_LONG).show();
                        }
                    } else {
                        Toast.makeText(MainActivity.this, "Selected file is not a Word document", Toast.LENGTH_LONG).show();
                    }
                }
            }
        });

        btn_wordtopdf.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                if (selectedDocumentUri != null) {
                    String fileExtension = getFileExtension(selectedDocumentUri);
                    if ("doc".equals(fileExtension) || "docx".equals(fileExtension)) {
                        convertWordToPdf(selectedDocumentUri);
                    } else {
                        Toast.makeText(MainActivity.this, "Selected file is not a Word document", Toast.LENGTH_LONG).show();
                    }
                }
            }
        });


        btn_exceltopdf.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                if (selectedDocumentUri != null) {
                    String fileExtension = getFileExtension(selectedDocumentUri);

                    if ("xls".equals(fileExtension) || "xlsx".equals(fileExtension)) {
                        convertExcelToPdf(selectedDocumentUri);
                    } else {
                        Toast.makeText(MainActivity.this, "Selected file is not an Excel file", Toast.LENGTH_LONG).show();
                    }
                }
            }
        });
    }
    private String getFileExtension(Uri uri) {
        String extension = null;

        if (uri.getScheme().equals("content")) {
            ContentResolver contentResolver = getContentResolver();
            MimeTypeMap mimeTypeMap = MimeTypeMap.getSingleton();
            String type = contentResolver.getType(uri);
            if (type != null) {
                extension = mimeTypeMap.getExtensionFromMimeType(type);
            }
        } else {
            String path = uri.getLastPathSegment();
            if (path != null) {
                int dotIndex = path.lastIndexOf(".");
                if (dotIndex >= 0) {
                    extension = path.substring(dotIndex + 1);
                }
            }
        }
        return extension;
    }


    private static void initProperties() {
        System.setProperty("org.apache.poi.javax.xml.stream.XMLInputFactory", "com.fasterxml.aalto.stax.InputFactoryImpl");
        System.setProperty("org.apache.poi.javax.xml.stream.XMLOutputFactory", "com.fasterxml.aalto.stax.OutputFactoryImpl");
        System.setProperty("org.apache.poi.javax.xml.stream.XMLEventFactory", "com.fasterxml.aalto.stax.EventFactoryImpl");
    }

    private void openDocumentFromFileManager() {
        Intent intent = new Intent(Intent.ACTION_GET_CONTENT);
        intent.setType("application/*");
        startActivityForResult(intent, PICK_PDF_FILE);
    }

    @RequiresApi(api = Build.VERSION_CODES.KITKAT)
    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        initProperties();

        if (requestCode == PICK_PDF_FILE && resultCode == RESULT_OK && data != null && data.getData() != null) {
            Uri uri = data.getData();
            String fileName = getFileName(uri);
            displaySelectedFileName(fileName);
            selectedDocumentUri = uri;
            showAllButtons();
        }
    }


    private String getFileName(Uri uri) {
        String result = null;
        if (uri.getScheme().equals("content")) {
            try (@SuppressLint({"NewApi", "LocalSuppress"}) Cursor cursor = getContentResolver().query(uri, null, null, null, null)) {
                if (cursor != null && cursor.moveToFirst()) {
                    result = cursor.getString(cursor.getColumnIndex(OpenableColumns.DISPLAY_NAME));
                }
            }
        }
        if (result == null) {
            result = uri.getPath();
            int cut = result.lastIndexOf('/');
            if (cut != -1) {
                result = result.substring(cut + 1);
            }
        }
        return result;
    }

    private void displaySelectedFileName(String fileName) {
        TextView selectedFileTextView = findViewById(R.id.textViewSelectedFile);
        selectedFileTextView.setText("Selected File: " + fileName);
    }

    private void showAllButtons() {
        btn_webview.setVisibility(View.VISIBLE);
        btn_wordtopdfAspose.setVisibility(View.VISIBLE);
        btn_wordtopdf.setVisibility(View.VISIBLE);
        btn_exceltopdf.setVisibility(View.VISIBLE);
    }

    private void convertExcelToPdf(Uri excelDocumentUri) {
        try {
            InputStream inputStream = getContentResolver().openInputStream(excelDocumentUri);
            Workbook workbook = WorkbookFactory.create(inputStream);

            // Generate a temporary PDF file
            File cacheDir = getCacheDir();
            String pdfFileName = "output_excel.pdf";
            File pdfFile = new File(cacheDir, pdfFileName);
            FileOutputStream fos = new FileOutputStream(pdfFile);

            // Use iText to convert Excel to PDF
            com.itextpdf.text.Document pdfDocument = new com.itextpdf.text.Document();
            com.itextpdf.text.pdf.PdfWriter.getInstance(pdfDocument, fos);
            pdfDocument.open();

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                for (Row row : sheet) {
                    for (Cell cell : row) {
                        pdfDocument.add(new com.itextpdf.text.Paragraph(cell.toString()));
                    }
                }
            }

            pdfDocument.close();
            fos.close();
            showPdfInViewer(pdfFileName);

        } catch (IOException | DocumentException | InvalidFormatException e) {
            e.printStackTrace();
            Toast.makeText(MainActivity.this, e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }

    private void convertWordToPdf(Uri wordDocumentUri) {
        try {
            InputStream inputStream = getContentResolver().openInputStream(wordDocumentUri);
            XWPFDocument document = new XWPFDocument(inputStream);

            // Generate a temporary PDF file
            File cacheDir = getCacheDir();
            String pdfFileName = "output.pdf";
            File pdfFile = new File(cacheDir, pdfFileName);
            FileOutputStream fos = new FileOutputStream(pdfFile);

            // Use iText to convert Word to PDF
            com.itextpdf.text.Document pdfDocument = new com.itextpdf.text.Document();
            com.itextpdf.text.pdf.PdfWriter.getInstance(pdfDocument, fos);
            pdfDocument.open();

            for (XWPFParagraph paragraph : document.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    pdfDocument.add(new com.itextpdf.text.Paragraph(run.getText(0)));
                }
            }

            pdfDocument.close();
            fos.close();
            showPdfInViewer(pdfFileName);

        } catch (IOException | DocumentException e) {
            e.printStackTrace();
            Toast.makeText(MainActivity.this, e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }

    private void showPdfInViewer(String pdfFileName) {
        Intent intent = new Intent(MainActivity.this, PdfView.class);
        intent.putExtra("pdfFileName", pdfFileName);
        startActivity(intent);
    }


    private void showPdfInViewerAspose(String pdfFileName) {
        Intent intent = new Intent(MainActivity.this, AsposeWordToPdf.class);
        intent.putExtra("pdfFileName", pdfFileName);
        startActivity(intent);
    }

    
}
