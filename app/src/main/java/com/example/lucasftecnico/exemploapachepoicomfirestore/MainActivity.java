package com.example.lucasftecnico.exemploapachepoicomfirestore;

import android.content.Intent;
import android.os.Bundle;
import android.support.annotation.Nullable;
import android.support.v7.app.AppCompatActivity;
import android.view.View;
import android.widget.Button;

import java.io.FileNotFoundException;
import java.io.OutputStream;

public class MainActivity extends AppCompatActivity {
    private Button btnCriarDocumento;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        btnCriarDocumento = (Button) findViewById(R.id.btnCriarDocumento);

        btnCriarDocumento.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intent = new Intent(Intent.ACTION_CREATE_DOCUMENT);
                intent.addCategory(Intent.CATEGORY_OPENABLE);
                intent.putExtra(Intent.EXTRA_TITLE, "Teste.xlsx");
                intent.setType("*/*");
                startActivityForResult(intent, 2);
            }
        });
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, @Nullable Intent data) {
        super.onActivityResult(requestCode, resultCode, data);

        if (resultCode == RESULT_OK) {
            switch (requestCode) {
                case 2:
                    requestCreateFile(data);
                    break;
            }
        }
    }


    public void requestCreateFile(Intent intent) {
        if (intent.getData() != null) {
            try {
                OutputStream outputStream = getApplicationContext().getContentResolver().openOutputStream(intent.getData());

                ManagerXLSX managerXLSX = new ManagerXLSX();
                managerXLSX.criar(outputStream);

            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
        }
    }
}
