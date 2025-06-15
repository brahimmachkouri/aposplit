// BM 2025
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace aposplit
{
    // Abstraction de la gestion PDF au cas où on voudrait changer de bibliothèques PDF
    public interface IPdfHandler
    {
        int GetPageCount(string pdfPath);
        string GetPageText(string pdfPath, int pageIndex);
        void AddPageToDocument(PdfSharp.Pdf.PdfDocument document, string pdfPath, int pageIndex);
    }

    /// <summary>
    /// Interface pour la gestion des PDF, permettant de séparer les responsabilités.
    /// Implémentée par PdfHandler qui utilise PdfSharp et PdfPig pour les opérations PDF.
    /// </summary>
    public class PdfHandler : IPdfHandler
    {
        /// <summary>
        /// Retourne le nombre de pages d'un PDF en utilisant PdfSharp.
        /// </summary>
        /// <param name="pdfPath"></param>
        /// <returns></returns>
        public int GetPageCount(string pdfPath)
        {
            using (var doc = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import))
                return doc.PageCount;
        }

        /// <summary>
        /// Extraction d'un texte d'une page spécifique d'un PDF en utilisant PdfPig.
        /// </summary>
        /// <param name="pdfPath"></param>
        /// <param name="pageIndex"></param>
        /// <returns></returns>
        public string GetPageText(string pdfPath, int pageIndex)
        {
            using (var pigDoc = UglyToad.PdfPig.PdfDocument.Open(pdfPath))
            {
                var page = pigDoc.GetPage(pageIndex + 1); // PdfPig pages are 1-based
                return page.Text;
            }
        }

        /// <summary>
        /// Ajoute une page d'un PDF source à un document PDF existant, en utilisant PdfSharp.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="pdfPath"></param>
        /// <param name="pageIndex"></param>
        public void AddPageToDocument(PdfSharp.Pdf.PdfDocument document, string pdfPath, int pageIndex)
        {
            using (var sourceDoc = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import))
            {
                document.AddPage(sourceDoc.Pages[pageIndex]);
            }
        }
    }

    /// <summary>
    /// Divise un PDF de relevés de notes en un fichier par étudiant, en utilisant le numéro étudiant comme clé.
    /// Les fichiers sont créés dans un sous-dossier nommé <nom_original_per_student>.
    /// </summary>
    public static class StudentPdfSplitter
    {
        // Constantes pour regex/chaînes magiques
        private const string StudentNumberRegex = @"N°\s*Etudiant\s*[:\-]?\s*(\d+)";
        private const string NameLineMarker = "Page :/";

        private static readonly Regex ReNumEtu = new(StudentNumberRegex, RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private record StudentPageInfo(string? Name, string? StudentNumber);

        /// <summary>
        /// Division d'un PDF de relevés de notes en un fichier par étudiant.
        /// </summary>
        /// <param name="pdfPath"></param>
        /// <param name="outputBaseDir"></param>
        public static void SplitByStudent(string pdfPath, string outputBaseDir)
        {
            ValidateInputPaths(pdfPath, outputBaseDir);

            string originalFileNameWithoutExt = Path.GetFileNameWithoutExtension(pdfPath);
            string sanitizedBaseName = SanitizeForFilename(originalFileNameWithoutExt);
            string studentOutputDir = Path.Combine(outputBaseDir, $"{sanitizedBaseName}_per_student");

            Directory.CreateDirectory(studentOutputDir);

            var pdfHandler = new PdfHandler();
            var studentDocuments = ExtractStudentDocuments(pdfPath, pdfHandler);

            SaveStudentDocuments(studentDocuments, sanitizedBaseName, studentOutputDir);

            Console.WriteLine($"Traitement terminé : {studentDocuments.Count} fichiers créés dans {studentOutputDir}");
        }

        /// <summary>
        /// Validation des chemins d'entrée et de sortie.
        /// </summary>
        /// <param name="pdfPath"></param>
        /// <param name="outputBaseDir"></param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="FileNotFoundException"></exception>
        private static void ValidateInputPaths(string pdfPath, string outputBaseDir)
        {
            if (string.IsNullOrEmpty(pdfPath)) throw new ArgumentNullException(nameof(pdfPath));
            if (!File.Exists(pdfPath)) throw new FileNotFoundException("Le fichier PDF source n'a pas été trouvé.", pdfPath);
            if (string.IsNullOrEmpty(outputBaseDir)) throw new ArgumentNullException(nameof(outputBaseDir));
        }

        /// <summary>
        /// Extraction des pages des relevés de notes pour chaque étudiant à partir du PDF source.
        /// </summary>
        /// <param name="pdfPath"></param>
        /// <param name="pdfHandler"></param>
        /// <returns></returns>
        private static List<(PdfSharp.Pdf.PdfDocument Pdf, StudentPageInfo Info)> ExtractStudentDocuments(
            string pdfPath, IPdfHandler pdfHandler)
        {
            var results = new List<(PdfSharp.Pdf.PdfDocument, StudentPageInfo)>();
            StudentPageInfo? currentStudentInfo = null;
            PdfSharp.Pdf.PdfDocument? currentStudentPdf = null;
            
            // On suppose que la première page est une page de garde et on l'ignore
            int pageCount = pdfHandler.GetPageCount(pdfPath);
            for (int i = 1; i < pageCount; i++) // On ignore la première page (page de garde)
            {
                // Extraction du texte de la page actuelle
                string pageText = pdfHandler.GetPageText(pdfPath, i);
                StudentPageInfo extractedInfo = ParseStudentInfoFromText(pageText);

                // Vérification si on a déjà un document pour l'étudiant actuel
                bool isNewStudent = currentStudentInfo == null ||
                                    (extractedInfo.StudentNumber != null && extractedInfo.StudentNumber != currentStudentInfo.StudentNumber) ||
                                    (extractedInfo.StudentNumber == null && currentStudentInfo.StudentNumber != null);
                // Si on a un nouveau numéro étudiant ou si on n'a pas de numéro étudiant actuel
                if (isNewStudent)
                {
                    if (currentStudentPdf != null && currentStudentInfo != null && currentStudentPdf.PageCount > 0)
                        results.Add((currentStudentPdf, currentStudentInfo));

                    currentStudentInfo = extractedInfo;
                    currentStudentPdf = new PdfSharp.Pdf.PdfDocument();
                }

                pdfHandler.AddPageToDocument(currentStudentPdf!, pdfPath, i);
            }

            // Ajout du dernier étudiant
            if (currentStudentPdf != null && currentStudentInfo != null && currentStudentPdf.PageCount > 0)
                results.Add((currentStudentPdf, currentStudentInfo));

            return results;
        }

        /// <summary>
        /// Enregistrement des documents PDF des étudiants dans le répertoire de sortie spécifié.
        /// </summary>
        /// <param name="studentDocuments"></param>
        /// <param name="baseName"></param>
        /// <param name="outputDir"></param>
        private static void SaveStudentDocuments(
            List<(PdfSharp.Pdf.PdfDocument Pdf, StudentPageInfo Info)> studentDocuments,
            string baseName, string outputDir)
        {
            foreach (var (pdf, info) in studentDocuments)
            {
                SaveStudentPdf(pdf, info, baseName, outputDir);
            }
        }

        /// <summary>
        /// Enregistrement du PDF d'un étudiant dans un fichier avec un nom basé sur son nom et son numéro étudiant.
        /// </summary>
        /// <param name="studentPdf"></param>
        /// <param name="studentInfo"></param>
        /// <param name="baseFileName"></param>
        /// <param name="outputDir"></param>
        private static void SaveStudentPdf(PdfSharp.Pdf.PdfDocument studentPdf, StudentPageInfo studentInfo, string baseFileName, string outputDir)
        {
            string safeStudentName = SanitizeForFilename(studentInfo.Name ?? "NomNonTrouve");
            string safeStudentNumber = SanitizeForFilename(studentInfo.StudentNumber ?? "NumNonTrouve");
            string fileName = $"{baseFileName}_{safeStudentName}_{safeStudentNumber}.pdf";
            string fullPath = Path.Combine(outputDir, fileName);

            try
            {
                studentPdf.Save(fullPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur lors de la sauvegarde du fichier {fullPath}: {ex.Message}");
            }
            finally
            {
                studentPdf.Close(); // Libération des ressources
            }
        }

        /// <summary>
        /// Analyse du texte d'une page pour extraire les informations de l'étudiant (nom et numéro étudiant).
        /// </summary>
        /// <param name="pageText"></param>
        /// <returns></returns>
        private static StudentPageInfo ParseStudentInfoFromText(string pageText)
        {
            var lines = pageText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            string? studentName = null;
            string? studentNumber = null;

            foreach (string line in lines)
            {
                string trimmedLine = line.Trim();
                
                if (studentName == null)
                {
                    // Recherche du nom de l'étudiant à partir de la ligne contenant le marqueur "Page :/"
                    int markerIndex = trimmedLine.IndexOf(NameLineMarker, StringComparison.OrdinalIgnoreCase);
                    if (markerIndex >= 0)
                    {
                        // On extrait le texte après le marqueur et on tente de trouver le nom
                        string textAfterMarker = trimmedLine[(markerIndex + NameLineMarker.Length)..].TrimStart();
                        var nameBuilder = new StringBuilder();
                        foreach (char c in textAfterMarker)
                        {
                            if (char.IsUpper(c) || char.IsWhiteSpace(c) || c == '\'' || c == '-')
                                nameBuilder.Append(c);
                            else break;
                        }

                        // On nettoie le nom potentiel et on vérifie s'il contient une initiale
                        string potentialNameWithInitial = nameBuilder.ToString().Trim();
                        if (!string.IsNullOrWhiteSpace(potentialNameWithInitial))
                        {
                            int lastSpaceIdx = potentialNameWithInitial.LastIndexOf(' ');
                            if (lastSpaceIdx > 0 &&
                                (potentialNameWithInitial.Length - lastSpaceIdx - 1 == 1) &&
                                char.IsUpper(potentialNameWithInitial[lastSpaceIdx + 1]))
                            {
                                studentName = potentialNameWithInitial.Substring(0, lastSpaceIdx).Trim();
                            }
                            else
                            {
                                studentName = potentialNameWithInitial;
                            }
                        }

                        // Si le nom est vide après nettoyage, on le met à null pour éviter de l'utiliser
                        if (string.IsNullOrWhiteSpace(studentName)) studentName = null;
                    }
                }

                // Recherche du numéro étudiant
                if (studentNumber == null)
                {
                    Match numMatch = ReNumEtu.Match(trimmedLine);
                    if (numMatch.Success)
                        studentNumber = numMatch.Groups[1].Value;
                }

                if (studentName != null && studentNumber != null)
                    break;
            }
            return new StudentPageInfo(studentName, studentNumber);
        }

        /// <summary>
        /// Nettoyoie une chaîne pour qu'elle soit utilisable comme nom de fichier.
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string SanitizeForFilename(string input)
        {
            if (string.IsNullOrEmpty(input))
                return "chaine_vide";
            string normalizedString = input.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(normalizedString.Length);
            foreach (char c in normalizedString)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                    sb.Append(c);
            }
            string lowerCaseString = sb.ToString().Normalize(NormalizationForm.FormC).ToLowerInvariant();
            string replaced = Regex.Replace(lowerCaseString, @"[^a-z0-9]", "_", RegexOptions.Compiled);
            string cleaned = Regex.Replace(replaced, @"_+", "_", RegexOptions.Compiled);
            cleaned = cleaned.Trim('_');
            if (string.IsNullOrEmpty(cleaned))
                return "nettoyage_vide";
            return cleaned;
        }
    }
}