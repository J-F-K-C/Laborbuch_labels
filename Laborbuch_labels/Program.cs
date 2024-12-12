using System;
using System.IO;
using System.Drawing;
using System.Drawing.Printing;

class ZebraPrinter
{
    static void Main()
    {
        string directoryPath = @"\\filesvr001\Users\in"; // Das Verzeichnis
        string filePattern = "*.gdt"; // Wildcard für .gdt Dateien

        // Suche nach Dateien im Verzeichnis, die dem Muster entsprechen
        string[] files = Directory.GetFiles(directoryPath, filePattern);

        if (files.Length == 0)
        {
            Console.WriteLine("Keine .gdt-Datei im angegebenen Verzeichnis gefunden.");
            return;
        }

        // Nehme die erste gefundene Datei
        string filePath = files[0];
        Console.WriteLine($"Verarbeite Datei: {filePath}");

        // Zeilen, aus denen wir die relevanten Daten extrahieren wollen (8, 9, 12, 13, 16)
        int[] relevantLines = { 8, 9, 12, 13, 16 };

        // Variablen, um die extrahierten Daten zu speichern
        string nachname = null;
        string vorname = null;
        DateTime? datum = null;
        string uhrzeit = null;
        string auftragsnummer = null;

        // Lese den Inhalt der Datei und analysiere jede Zeile
        try
        {
            string[] lines = File.ReadAllLines(filePath);

            // Durchlaufe nur die relevanten Zeilen
            foreach (var lineIndex in relevantLines)
            {
                if (lineIndex - 1 < lines.Length) // Überprüfen, ob die Zeile existiert
                {
                    string line = lines[lineIndex - 1]; // Zeilen sind nullbasiert, also -1

                    // Extrahiere ab der 8. Spalte (Index 7) der Zeile
                    string extractedData = line.Length > 7 ? line.Substring(7).Trim() : string.Empty;

                    if (lineIndex == 8) // Zeile 8: Nachname
                    {
                        nachname = extractedData;
                    }
                    else if (lineIndex == 9) // Zeile 9: Vorname
                    {
                        vorname = extractedData;
                    }
                    else if (lineIndex == 12) // Zeile 12: Datum (tt.mm.jjjj)
                    {
                        if (DateTime.TryParseExact(extractedData, "ddMMyyyy", null, System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                        {
                            datum = parsedDate;
                        }
                    }
                    else if (lineIndex == 13) // Zeile 13: Uhrzeit (hhmmss)
                    {
                        if (extractedData.Length == 6) // Erwartet  Zeichen z.B. 113820
                        {
                            string hours = extractedData.Substring(0, 2);
                            string minutes = extractedData.Substring(2, 2);

                            // Parse nur Stunden und Minuten (hhmm)
                            if (int.TryParse(hours, out int hour) && int.TryParse(minutes, out int minute))
                            {
                                uhrzeit = $"{hour:D2}:{minute:D2}";
                            }
                        }
                    }
                    else if (lineIndex == 16) // Zeile 16: Auftragsnummer
                    {
                        auftragsnummer = extractedData;
                    }
                }
            }

            // Formatieren des ZPL-Drucks
            string zplCommand = ""; // Start des ZPL Drucks

            if (nachname != null)
                zplCommand += $"Nachname:{nachname}\n";
            if (vorname != null)
                zplCommand += $"Vorname:{vorname}\n";
            if (datum.HasValue)
                zplCommand += $"{datum.Value:dd.MM.yyyy}\n";
            if (uhrzeit != null)
                zplCommand += $"{uhrzeit} Uhr\n";
            if (auftragsnummer != null)
                zplCommand += $"{auftragsnummer}\n";

            // Druckauftrag an den Zebra-Drucker senden
            string printerName = "ZDesigner ZD410-203dpi ZPL"; // Der Name des Zebra-Druckers laut Systemsteuerung

            // Erstelle das PrintDocument-Objekt
            PrintDocument printDoc = new PrintDocument();
            printDoc.PrinterSettings.PrinterName = printerName;

            // Beim Drucken ZPL-Befehl an den Drucker senden
            printDoc.PrintPage += (sender, e) =>
            {
                // ZPL-Befehl direkt auf dem Drucker drucken
                e.Graphics.DrawString(zplCommand, new Font("Arial", 7), Brushes.Black, 10, 10);
            };

            try
            {
                // Druckauftrag ausführen
                printDoc.Print();
                Console.WriteLine("Druckauftrag erfolgreich gesendet!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fehler: " + ex.Message);
            }

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler beim Verarbeiten der Datei: {ex.Message}");
        }
    }
}
