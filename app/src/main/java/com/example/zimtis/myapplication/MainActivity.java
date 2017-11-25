package com.example.zimtis.myapplication;

import android.os.Bundle;
import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Toast;

import java.io.File;
import java.io.IOException;

import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class MainActivity extends AppCompatActivity implements View.OnClickListener {

    public static String FOLDER = "testFolder";

    private EditText editText;
    private Button   save;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        try {
            // Hier wird geguckt, ob der Ordner bereits vorhanden ist, wenn nicht, wird er erstellt.
            File newFolder = new File(Environment.getExternalStorageDirectory(), FOLDER);
            if (!newFolder.exists()) {
                newFolder.mkdir();
            }
        } catch (Exception e) {
            Log.e("error", e.getMessage());
        }

        // referenzen auf die editText und button holen
        this.editText = findViewById(R.id.editText);
        this.save = findViewById(R.id.button);

        // on click listener einfügen, this kann nur genommen werden, wenn diese Klasse View.OnClickListener implemewntiert, Siehe implements View.OnClickListener
        this.save.setOnClickListener(this);

    }

    @Override
    public void onClick(View view) {
        // wert auslesen
        String value = this.editText.getText().toString();

        // Excel datei öffnen

        File folder = new File(Environment.getExternalStorageDirectory(), FOLDER);

        /// wir gucken ob wir die vorlage haben

        File preset = new File(folder, "Vorlage.xls");
        if (preset.exists()) {

            // hier setzen wir den dateinamen der neuen datei
            File newFile = new File(folder, "Protokoll.xls");

            try {
                Workbook workbook = Workbook.getWorkbook(preset);
                WritableWorkbook newWorkbook = Workbook.createWorkbook(newFile, workbook);

                // wir holen uns die erste Tabelle aus der Datei
                WritableSheet sheet = newWorkbook.getSheet(0);

                // wir ersetllen ein neues label welches an der stelle 0,0 angezeigt werden soll, also in der ersten spalte und ersten reihe
                // hier setzten wir auch gleich den Wert für die neue Zelle
                Label newValue = new Label(0, 0, value);
                // wir fühgen die neue Zelle in die Tabell ein
                sheet.addCell(newValue);

                // new Exceldatei schreiben
                newWorkbook.write();
                // und abschließen
                newWorkbook.close();
                // alte tabelle schließen
                workbook.close();

            } catch (IOException e) {
                e.printStackTrace();
            } catch (BiffException e) {
                e.printStackTrace();
            } catch (WriteException e) {
                e.printStackTrace();
            }
        } else {
            Toast.makeText(this, "Vorlage datei wurde nicht gefunden", Toast.LENGTH_SHORT);
        }

    }
}
