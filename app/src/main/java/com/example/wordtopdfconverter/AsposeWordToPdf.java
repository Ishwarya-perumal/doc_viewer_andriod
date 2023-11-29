// PdfViewActivity.java
package com.example.wordtopdfconverter;

import android.annotation.SuppressLint;
import android.os.Bundle;
import android.view.View;
import android.widget.ImageView;

import androidx.appcompat.app.AppCompatActivity;

import com.github.barteksc.pdfviewer.PDFView;

import java.io.File;

public class AsposeWordToPdf extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_pdf_view);
        @SuppressLint({"MissingInflatedId", "LocalSuppress"}) ImageView backArrow = findViewById(R.id.backArrow);

        backArrow.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                finish();
            }
        });
        String pdfFileName = getIntent().getStringExtra("pdfFileName");
        PDFView pdfView = findViewById(R.id.pdfView);
        pdfView.fromFile(new File(pdfFileName)).load();
    }
}
