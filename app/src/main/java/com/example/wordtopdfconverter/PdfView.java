package com.example.wordtopdfconverter;



import androidx.appcompat.app.AppCompatActivity;
import com.github.barteksc.pdfviewer.util.FitPolicy;
import android.annotation.SuppressLint;
import android.content.Intent;
import android.os.Bundle;
import android.view.MenuItem;
import android.view.View;
import android.widget.ImageView;
import android.widget.Toolbar;
import com.github.barteksc.pdfviewer.util.FitPolicy;


import com.github.barteksc.pdfviewer.PDFView;

import java.io.File;

public class PdfView extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_pdf_view);

        @SuppressLint({"MissingInflatedId", "LocalSuppress"}) PDFView pdfView = findViewById(R.id.pdfView);

        @SuppressLint({"MissingInflatedId", "LocalSuppress"}) ImageView backArrow = findViewById(R.id.backArrow);

        backArrow.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                finish();
            }
        });
        Intent intent = getIntent();
        if (intent != null && intent.hasExtra("pdfFileName")) {
            String pdfFileName = intent.getStringExtra("pdfFileName");

            File pdfFile = new File(getCacheDir(), pdfFileName);
            pdfView.fromFile(pdfFile)
                    .pageFitPolicy(FitPolicy.WIDTH)
                    .load();
        }

    }

}

