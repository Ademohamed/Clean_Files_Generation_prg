/*
 * Ce projet utilise la bibliothèque EPPlus (version 4.5.3)
 * pour manipuler des fichiers Excel. EPPlus est sous licence LGPL.
 * Pour plus d'informations, consultez : https://github.com/EPPlusSoftware/EPPlus
 */



using System;
using System.IO;
using OfficeOpenXml;
using System.Linq;
using System.ComponentModel;

class Program
{
    static void Main(string[] args)
    {
        //if (args.Length < 3)
        {
            Console.WriteLine("les arguments sont pas complets ...!!");
            return;
        }
        //version test

        string cheminFichier = args[0];   // Chemin vers le fichier Excel de données
        string cheminTranco = args[1];      // Chemin vers le fichier Excel de transco
        string nbCol = args[2];         // Cellule à modifier (ex: A1)
        string separateur = ";";

        //var cheminFichier = "C:\\Users\\mohamed-n-ade\\OneDrive - MMC\\Documents\\Mercer\\Projets\\Programme Modif entete Generation\\Consommation_GENERATION entete original.csv";   // Chemin vers le fichier Excel de données
        //var cheminTranco = "C:\\Users\\mohamed-n-ade\\OneDrive - MMC\\Documents\\Mercer\\Projets\\Programme Modif entete Generation\\transco variable genration vers SAS.xlsx";      // Chemin vers le fichier Excel de transco
        //string nbCol = "60";         // Cellule à modifier (ex: A1)
        //string separateur = ";";

        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;


        // Lire les données de la colonne A du fichier Excel
        string[] columnAData;
        using (var package = new ExcelPackage(new FileInfo(cheminTranco)))
        {
            var worksheet = package.Workbook.Worksheets[0]; // Accéder à la première feuille de calcul
            int rowCount = int.Parse(nbCol); // Obtenir le nombre total de lignes

            columnAData = new string[rowCount];

            // Parcourir toutes les lignes de la colonne A
            for (int row = 1; row <= rowCount; row++)
            {
                columnAData[row - 1] = worksheet.Cells[row, 1].Text; // Lire le texte de la colonne A (colonne 1)
            }
        }

        

        // Lire tout le contenu du fichier CSV en une seule fois
        string csvContent = File.ReadAllText(cheminFichier);

        // Diviser le contenu en lignes
        string[] csvLines = csvContent.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

        // Remplacer la première ligne par les données de la colonne A
        csvLines[0] = string.Join(separateur, columnAData);

        // Réécrire tout le contenu dans le fichier CSV avec la première ligne modifiée
        File.WriteAllText(cheminFichier, string.Join(Environment.NewLine, csvLines));
        Console.WriteLine("Les valeurs ont été remplacées avec succès !");
    }
}
