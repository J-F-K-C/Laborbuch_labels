using System;
using System.IO;
using System.Text;
using System.Drawing;
using System.Drawing.Printing;

class ZebraPrinter
{
    static void Main()
    {
        string directoryPath = @"U:\IN"; // Das Verzeichnis
        string filePattern = "*.gdt"; // Wildcard für .gdt Dateien

        // Suche nach Dateien im Verzeichnis, die dem Muster entsprechen
        string[] files = Directory.GetFiles(directoryPath, filePattern);

        if (files.Length == 0)
        {
            Console.WriteLine("Keine .gdt-Datei im angegebenen Verzeichnis gefunden.");
            return;
        }

        foreach (var filePath in files)
        {
            Console.WriteLine($"Verarbeite Datei: {filePath}");

            // Datei mit ISO-8859-15-Kodierung lesen
            string[] lines;
            using (var reader = new StreamReader(filePath, Encoding.GetEncoding("ISO-8859-1")))
            {
                lines = reader.ReadToEnd().Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            }

            string nachname = null;
            string vorname = null;
            DateTime? datum = null;
            string uhrzeit = null;
            string auftragsnummer = null;

            // Zeilen, aus denen wir die relevanten Daten extrahieren wollen (8, 9, 16, 17, 19)
            int[] relevantLines = { 8, 9, 16, 17, 19 };

            foreach (var lineIndex in relevantLines)
            {
                if (lineIndex - 1 < lines.Length)
                {
                    string line = lines[lineIndex - 1];
                    string extractedData = line.Length > 7 ? line.Substring(7).Trim() : string.Empty;

                    if (lineIndex == 8)
                        nachname = extractedData;
                    else if (lineIndex == 9)
                        vorname = extractedData;
                    else if (lineIndex == 16 && DateTime.TryParseExact(extractedData, "ddMMyyyy", null, System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                        datum = parsedDate;
                    else if (lineIndex == 17 && extractedData.Length >= 4)
                        uhrzeit = $"{extractedData.Substring(0, 2)}:{extractedData.Substring(2, 2)}";
                    else if (lineIndex == 19)
                        auftragsnummer = extractedData;
                }
            }

            // Formatieren des Drucktextes
            string zplCommand = "";
            if (nachname != null) zplCommand += $"Nachname: {nachname}\n";
            if (vorname != null) zplCommand += $"Vorname: {vorname}\n";
            if (datum.HasValue) zplCommand += $"{datum.Value:dd.MM.yyyy}\n";
            if (uhrzeit != null) zplCommand += $"{uhrzeit} Uhr\n";
            if (auftragsnummer != null) zplCommand += $"{auftragsnummer}\n";



            // Druckauftrag an den Drucker senden
            string printerName = "ZDesigner ZD410-203dpi ZPL (Kopie 2)";
            PrintDocument printDoc = new PrintDocument();
            printDoc.PrinterSettings.PrinterName = printerName;
            printDoc.PrintPage += (sender, e) =>
            {
                e.Graphics.DrawString(zplCommand, new Font("Arial", 7), Brushes.Black, 10, 10);
            };

            try
            {
                printDoc.Print();
                Console.WriteLine("Druckauftrag erfolgreich gesendet!");
                File.Delete(filePath);
                Console.WriteLine($"Datei {filePath} erfolgreich gelöscht.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fehler: " + ex.Message);
            }
        }
    }
}
