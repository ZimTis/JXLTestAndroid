package com.example.zimtis.myapplication;

import android.Manifest;
import android.annotation.SuppressLint;
import android.os.Bundle;
import android.os.Environment;
import android.support.annotation.NonNull;
import android.support.v7.app.AppCompatActivity;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Toast;

import java.io.File;
import java.io.IOException;

import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import me.zhanghai.android.effortlesspermissions.EffortlessPermissions;
import pub.devrel.easypermissions.AfterPermissionGranted;
import pub.devrel.easypermissions.EasyPermissions;

public class MainActivity extends AppCompatActivity implements View.OnClickListener {

    public static       String FOLDER                   = "testFolder";
    public static final int    REQUEST_EXTERNAL_STORAGE = 500;

    private EditText editText;
    private Button   save;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        // Wir fragen nach den Benötigten Berechtigungen.
        this.checkPermission();

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

    @AfterPermissionGranted(REQUEST_EXTERNAL_STORAGE)
    private void checkPermission() {
        String[] permissions = {Manifest.permission.WRITE_EXTERNAL_STORAGE, Manifest.permission.READ_EXTERNAL_STORAGE};
        if (EffortlessPermissions.hasPermissions(this, permissions)) {
        } else {
            EffortlessPermissions.requestPermissions(this, "Um zu funktionieren, müssen wir einige Berechtigungen erfragen.", REQUEST_EXTERNAL_STORAGE, permissions);
        }
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);

        EasyPermissions.onRequestPermissionsResult(requestCode, permissions, grantResults, this);
    }

    @Override
    public void onClick(View view) {

        // wert auslesen
        String value = this.editText.getText().toString();

        // wir holen uns eine Referenz auf den Ordner i ndem die Protokolle gespeichert werden sollen

        File folder = new File(Environment.getExternalStorageDirectory(), FOLDER);

        /// wir gucken ob wir die Vorlage.xls haben

        File preset = new File(folder, "Vorlage.xls");
        if (preset.exists()) {

            // hier setzen wir den dateinamen der neuen datei
            File newFile = new File(folder, "Protokoll.xls");

            try {
                Workbook workbook = Workbook.getWorkbook(preset);
                WritableWorkbook newWorkbook = Workbook.createWorkbook(newFile, workbook);

                // wir holen uns die erste Tabelle aus der Datei
                WritableSheet sheet = newWorkbook.getSheet(0);

                //- Den Nachfolgenden Bereich müssen wir für jeden Wert wiederholen, den wir schreiben wollen. Natürlich mit den jeweiligen Daten

                // wir ersetllen ein neues label welches an der stelle 0,0 angezeigt werden soll, also in der ersten spalte der ersten Reihe
                // hier setzten wir auch gleich den Wert für die neue Zelle
                Label newValue = new Label(0, 0, value);
                // wir fühgen die neue Zelle in die Tabell ein
                sheet.addCell(newValue);

                //-

                // neue Exceldatei schreiben
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
