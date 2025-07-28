import java.io.FileInputStream;             // Import der Klasse für den Dateizugriff
import java.util.ArrayList;// Import der Klasse für die ArrayList, um dynamische Arrays zu erstellen
import java.util.List; // Import der Klasse für die List, um eine Liste von Spalteninhalten zu erstellen
import org.apache.poi.ss.usermodel.Cell;  //Import der Klasse für den Zugriff auf Zelleninhalte
import org.apache.poi.ss.usermodel.Row;      // Import der Klasse für den Zugriff auf Zeileninhalte
import org.apache.poi.ss.usermodel.Sheet;// Import der Klasse für den Zugriff auf Tabellenblätter
import org.apache.poi.ss.usermodel.Workbook;// Import der Klasse für die Arbeitsmappe
import org.apache.poi.util.IOUtils;// Import der Klasse für IOUtils, um das ByteArray-Limit zu setzen
import org.apache.poi.xssf.usermodel.XSSFWorkbook;// Import der Klasse für die Verarbeitung von Excel-Dateien im XLSX-Format

public class Einmalige_Werte_Benchmark {

    public static void main(String[] args) throws Exception {
        
        IOUtils.setByteArrayMaxOverride(200_000_000);
        // Setze das ByteArray-Limit auf 200 MB, um große Excel-Dateien zu verarbeiten

        String Pfad = "C:\\Users\\marti\\Desktop\\Datensätze\\Hotelbuchungen.xlsx"; 
        // Pfad zur Excel-Datei, die verarbeitet werden soll

        int Anzahl_Durchlaeufe = 3;                   
        // Anzahl der Iterationen für den Benchmark-Test, um die Stabilität der Ergebnisse zu prüfen
        long Gesamtlaufzeit = 0;               
        // Summe der Laufzeiten in Nanosekunden
        long Gesamtspeicher = 0;                   
        // Summe des belegten Speichers in KB

        for (int Durchlauf = 1; Durchlauf <= Anzahl_Durchlaeufe; Durchlauf++) {
            //Durchführung der Benchmark-Tests für die angegebene Anzahl an Durchläufen
            System.gc();                             
            //Garbage Collector aufrufen, um den Speicher zu bereinigen und den Speicherbedarf zu minimieren
            Thread.sleep(100);                      
             // Kurze Pause, um sicherzustellen, dass der Garbage Collector seine Arbeit abgeschlossen hat
            long Speicher_vorher = getVerwendeterSpeicher(); 
          //Speicherbedarf in KB vor der Ausführung des Codes ermitteln

            long Startzeit = System.nanoTime();      
            // Zeit in Nanosekunden vor Ausführung des Codes erfassen

            FileInputStream fis = new FileInputStream(Pfad); 
            // Excel Datei mittels FileInputStream öffnen
            Workbook Arbeitsmappe = new XSSFWorkbook(fis);        
            // Arbeitsmappe aus der Datei abrufen
            Sheet Tabelle = Arbeitsmappe.getSheetAt(0);           
            // Tabellenblatt 0 (erstes Blatt) der Arbeitsmappe abrufen

            int spaltenAnzahl = Tabelle.getRow(0).getLastCellNum(); 
            // Ermittlung der Anzahl der Spalten im ersten Blatt
            List<List<String>> Spaltendaten = new ArrayList<>();   
             // Erstellung einer Liste von Listen, um die Spalteninhalte zu speichern

            // Erstellung einer Leeren Liste für jede Spalte
            for (int i = 0; i < spaltenAnzahl; i++) {
                // Erstellung der Spaltenlisten
                Spaltendaten.add(new ArrayList<>());
                // Jede Spalte erhält eine eigene Liste
            }

            // Befüllung der Spaltenlisten mit den Werten aus der Spalte
            for (Row Zeile : Tabelle) { 
                // Iteration über alle Zeilen in der Tabelle
                for (int s = 0; s < spaltenAnzahl; s++) {
                    // Iteration über alle Spalten der Zeile
                    Cell Zelle = Zeile.getCell(s);                     
                    // Aufruf der Zelle in der aktuellen Zeile und Spalte
                    String Wert = (Zelle != null) ? Zelle.toString() : ""; 
                    // Aufruf des Zelleninhaltes auch wenn die Zelle leer ist
                    Spaltendaten.get(s).add(Wert);                     
                    // Hinzufügen des Wertes zur entsprechenden Spaltenliste
                }
            }

            //  Durchlauf durch die Spaltendaten um die eindeutigen Werte zu zählen
            for (int s = 0; s < Spaltendaten.size(); s++) {
                List<String> Spalte = Spaltendaten.get(s); 
                // Spalte die aktuell bearbeitet wird
                int Anzahl_einmalig = 0;                   
                 // Zählervariable für eindeutige Werte in der Spalte

                // Prüfung für jeden Eintrag, ob dieser nur einmal in der Spalte vorkommt
                for (int i = 0; i < Spalte.size(); i++) {
                    // Durchlauf aller Einträge in der Spalte
                    String aktueller_Wert = Spalte.get(i);
                    //Wert für den aktuellen Prüfungseintrag
                    int Haeufigkeit = 0;
                    // Zählervariable für die Häufigkeit des aktuellen Wertes

                    // Durchlauf aller Einträge in der Spalte, um die Häufigkeit des aktuellen Wertes zu zählen
                    for (int j = 0; j < Spalte.size(); j++) {
                        // Durchlauf aller Einträge in derselben Spalte
                        if (Spalte.get(j).equals(aktueller_Wert)) {
                            // Prüfung, ob der aktuelle Wert mit dem Vergleichswert übereinstimmt
                            Haeufigkeit++; 
                            // Erhöhung der Häufigkeit, wenn der Wert übereinstimmt
                        }
                    }

                    // Prüfung, ob der aktuelle Wert nur einmal in der Spalte vorkommt
                    if (Haeufigkeit == 1) {
                        //Wenn Prüfung erfolgreich, dann ist der Wert eindeutig
                    Anzahl_einmalig++;
                    // Erhöhung des Zählers für eindeutige Werte
                    }
                }

                // Ausgabe der Anzahl der eindeutigen Werte für die aktuelle Spalte
                System.out.println("Durchlauf " + Durchlauf + " – Spalte " + (s + 1) + ": " + Anzahl_einmalig + " eindeutige Werte");
            }

            Arbeitsmappe.close(); 
            // Schließen der Arbeitsmappe, um Ressourcen freizugeben
            fis.close();
            // Schließen des FileInputStream, um Ressourcen freizugeben

            long Endzeit = System.nanoTime();
             // Abruf der Zeit in Nanosekunden nach der Ausführung des Codes
            long LaufzeitMillis = (Endzeit - Startzeit) / 1_000_000; 
            // Berechnung der Gesamtlaufzeit in Millisekunden
            long SpeicherNachher = getVerwendeterSpeicher(); 
            // Abruf des aktuell belegten Speichers in KB nach der Ausführung des Codes
            long BelegterSpeicherKB = (SpeicherNachher - Speicher_vorher) / 1024; 
            // Berechnung des belegten Speichers in KB

            // Aufaddieren der Laufzeit und des belegten Speichers für die Gesamtauswertung
            Gesamtlaufzeit += LaufzeitMillis;
            // Addition der Laufzeit in Millisekunden
            Gesamtspeicher += BelegterSpeicherKB;
            // Addition des belegten Speichers in KB

            // Ausgabe der Ergebnisse für den aktuellen Durchlauf
            System.out.println(" Laufzeit Durchlauf " + Durchlauf + ": " + LaufzeitMillis + " ms");
            // Ausgabe der Laufzeit in Millisekunden
            System.out.println("Speicherverbrauch Durchlauf " + Durchlauf + ": " + BelegterSpeicherKB + " KB");
            // Ausgabe des Speicherverbrauchs in KB
            System.out.println("---------------------------------------------");
        }

        // Gesamtauswertung nach allen Wiederholungen
        System.out.println("Durchschnittliche Laufzeit: " + (Gesamtlaufzeit / Anzahl_Durchlaeufe) + " ms");
        // Durchschnittliche Laufzeit in Millisekunden
        System.out.println(" Durchschnittlicher Speicherverbrauch: " + (Gesamtspeicher / Anzahl_Durchlaeufe) + " KB");
        // Durchschnittlicher Speicherverbrauch in KB
    }

    // Hilfsmethode: berechnet aktuell verwendeten Heap-Speicher
    public static long getVerwendeterSpeicher() {
        Runtime laufzeit = Runtime.getRuntime();
        return laufzeit.totalMemory() - laufzeit.freeMemory();
    }
}
