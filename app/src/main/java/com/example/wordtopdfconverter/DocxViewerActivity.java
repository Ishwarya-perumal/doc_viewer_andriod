package com.example.wordtopdfconverter;

import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;
import android.os.Bundle;
import android.annotation.SuppressLint;
import android.net.Uri;
import android.webkit.WebSettings;
import android.webkit.WebView;



import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.IOException;
import java.io.InputStream;

public class DocxViewerActivity extends AppCompatActivity {

    private WebView webView;

    @SuppressLint({"SetJavaScriptEnabled", "MissingInflatedId"})
    @Override
    protected void onCreate(@Nullable Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_docx_viewer);

        webView = findViewById(R.id.docxWebView);
        WebSettings webSettings = webView.getSettings();
        webSettings.setJavaScriptEnabled(true);
        Uri documentUri = getIntent().getData();

        try {
            InputStream inputStream = getContentResolver().openInputStream(documentUri);
            XWPFDocument docx = new XWPFDocument(inputStream);
            String htmlContent = generateHtmlContent(docx);
            webView.loadDataWithBaseURL(null, htmlContent, "text/html", "utf-8", null);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private String generateHtmlContent(XWPFDocument docx) {
        StringBuilder htmlBuilder = new StringBuilder();

        htmlBuilder.append("<html><head><style>body{font-size:14px;}</style></head><body>");
        for (XWPFParagraph paragraph : docx.getParagraphs()) {
            htmlBuilder.append("<p>").append(paragraph.getText()).append("</p>");
        }
        htmlBuilder.append("</body></html>");
        return htmlBuilder.toString();
    }


}
