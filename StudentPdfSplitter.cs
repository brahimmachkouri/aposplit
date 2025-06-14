// BM 2025
using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using PdfSharp.Pdf.IO; // Nécessaire pour PdfReader
using PdfSharpDoc = PdfSharp.Pdf.PdfDocument;    // alias pour PdfSharp
using PigDoc = UglyToad.PdfPig.PdfDocument;      // alias pour PdfPig
using UglyToad.PdfPig.Content;                   // pour Page
using System.Globalization;

namespace aposplit
{
    /// <summary>
    /// Divise un PDF de relevés de notes en un fichier par étudiant, en utilisant le numéro étudiant comme clé.
    /// Les fichiers sont créés dans un sous-dossier nommé <nom_original_per_student>.
    /// </summary>
    public static class StudentPdfSplitter
    {
        // Regex pour extraire le numéro étudiant (ex: "N° Etudiant : 12345")
        private static readonly Regex ReNumEtu = new(@"N°\s*Etudiant\s*[:\-]?\s*(\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Record pour stocker les informations extraites d'une page
        private record StudentPageInfo(string? Name, string? StudentNumber);

        /// <summary>
        /// Divise le fichier PDF spécifié en plusieurs fichiers PDF, un par étudiant.
        /// </summary>
        /// <param name="pdfPath">Chemin vers le fichier PDF source.</param>
        /// <param name="outputBaseDir">Dossier de base où sera créé le sous-dossier de sortie.</param>
        public static void SplitByStudent(string pdfPath, string outputBaseDir)
        {
            if (string.IsNullOrEmpty(pdfPath)) throw new ArgumentNullException(nameof(pdfPath));
            if (!File.Exists(pdfPath)) throw new FileNotFoundException("Le fichier PDF source n'a pas été trouvé.", pdfPath);
            if (string.IsNullOrEmpty(outputBaseDir)) throw new ArgumentNullException(nameof(outputBaseDir));

            string originalFileNameWithoutExt = Path.GetFileNameWithoutExtension(pdfPath);
            string sanitizedBaseName = SanitizeForFilename(originalFileNameWithoutExt);
            string studentOutputDir = Path.Combine(outputBaseDir, $"{sanitizedBaseName}_per_student");
            Directory.CreateDirectory(studentOutputDir); // Crée le dossier s'il n'existe pas

            using PdfSharpDoc sourcePdfSharpDoc = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import);
            using PigDoc sourcePigDoc = PigDoc.Open(pdfPath);

            StudentPageInfo? currentStudentInfo = null;
            PdfSharpDoc? currentStudentPdf = null;
            int filesCreatedCount = 0;

            // On présume que la première page (index 0 pour PdfSharp, page 1 pour PdfPig) est une page de garde et on l'ignore.
            // PdfSharp est 0-basé pour les pages, PdfPig est 1-basé.
            for (int i = 1; i < sourcePdfSharpDoc.PageCount; i++)
            {
                Page pigPage = sourcePigDoc.GetPage(i + 1); // PdfPig page numbers are 1-based
                string pageText = pigPage.Text;
                StudentPageInfo extractedInfo = ParseStudentInfoFromText(pageText);

                bool isNewStudent = currentStudentInfo == null ||
                                    (extractedInfo.StudentNumber != null && extractedInfo.StudentNumber != currentStudentInfo.StudentNumber) ||
                                    (extractedInfo.StudentNumber == null && currentStudentInfo.StudentNumber != null); // Gère le cas où un numéro disparaîtrait

                if (isNewStudent)
                {
                    if (currentStudentPdf != null && currentStudentInfo != null && currentStudentPdf.PageCount > 0)
                    {
                        SaveStudentPdf(currentStudentPdf, currentStudentInfo, sanitizedBaseName, studentOutputDir);
                        filesCreatedCount++;
                    }
                    currentStudentInfo = extractedInfo;
                    currentStudentPdf = new PdfSharpDoc();
                }

                // S'assurer que currentStudentPdf est initialisé avant d'ajouter une page
                currentStudentPdf?.AddPage(sourcePdfSharpDoc.Pages[i]);
            }

            // Sauvegarder le PDF du dernier étudiant traité
            if (currentStudentPdf != null && currentStudentInfo != null && currentStudentPdf.PageCount > 0)
            {
                SaveStudentPdf(currentStudentPdf, currentStudentInfo, sanitizedBaseName, studentOutputDir);
                filesCreatedCount++;
            }

            Console.WriteLine($"Traitement terminé : {filesCreatedCount} fichiers créés dans {studentOutputDir}");
        }

        /// <summary>
        /// Extrait le nom et le numéro de l'étudiant à partir du texte d'une page.
        /// </summary>
        private static StudentPageInfo ParseStudentInfoFromText(string pageText)
        {
            var lines = pageText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            string? studentName = null;
            string? studentNumber = null;

            // Le nom de famille est indiqué par une ligne contenant "Page :/" suivi du nom de l'étudiant. 
            // (différent de la bibliothèque Python, qui ne récupère pas les mêmes séquences textuelles) 
            const string nameLineMarker = "Page :/";

            foreach (string line in lines)
            {
                string trimmedLine = line.Trim();

                if (studentName == null)
                {
                    int markerIndex = trimmedLine.IndexOf(nameLineMarker, StringComparison.OrdinalIgnoreCase);
                    if (markerIndex >= 0)
                    {
                        string textAfterMarker = trimmedLine.Substring(markerIndex + nameLineMarker.Length).TrimStart();
                        var nameBuilder = new StringBuilder();
                        foreach (char c in textAfterMarker)
                        {
                            // Capture les majuscules, espaces, apostrophes, tirets
                            // S'arrête au premier caractère qui ne correspond pas (ex: minuscule du prénom complet, chiffre, etc.)
                            if (char.IsUpper(c) || char.IsWhiteSpace(c) || c == '\'' || c == '-')
                            {
                                nameBuilder.Append(c);
                            }
                            else
                            {
                                break;
                            }
                        }

                        string potentialNameWithInitial = nameBuilder.ToString().Trim();

                        if (!string.IsNullOrWhiteSpace(potentialNameWithInitial))
                        {
                            int lastSpaceIdx = potentialNameWithInitial.LastIndexOf(' ');
                            // Vérifier si la partie après le dernier espace est une seule lettre majuscule (l'initiale)
                            if (lastSpaceIdx > 0 &&
                                (potentialNameWithInitial.Length - lastSpaceIdx - 1 == 1) &&
                                char.IsUpper(potentialNameWithInitial[lastSpaceIdx + 1]))
                            {
                                studentName = potentialNameWithInitial.Substring(0, lastSpaceIdx).Trim(); // Prendre la partie avant l'initiale
                            }
                            else
                            {
                                studentName = potentialNameWithInitial; // Pas d'initiale ou format différent, on garde ce qu'on a.
                            }
                        }
                        if (string.IsNullOrWhiteSpace(studentName)) studentName = null;
                    }
                }

                if (studentNumber == null)
                {
                    Match numMatch = ReNumEtu.Match(trimmedLine);
                    if (numMatch.Success)
                    {
                        studentNumber = numMatch.Groups[1].Value;
                    }
                }

                if (studentName != null && studentNumber != null)
                {
                    break;
                }
            }
            // Si le nom n'est pas trouvé, on peut utiliser une valeur par défaut ou laisser null
            // Idem pour le numéro étudiant. La gestion des nulls se fera dans SaveStudentPdf
            return new StudentPageInfo(studentName, studentNumber);
        }

        /// <summary>
        /// Sauvegarde le PdfDocument PdfSharp pour l'étudiant courant.
        /// </summary>
        private static void SaveStudentPdf(PdfSharpDoc studentPdf, StudentPageInfo studentInfo, string baseFileName, string outputDir)
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
                // Log l'erreur ou la remonter. Pour l'instant, on l'affiche en console.
                Console.WriteLine($"Erreur lors de la sauvegarde du fichier {fullPath}: {ex.Message}");
            }
            finally
            {
                studentPdf.Close(); // Important de fermer le document pour libérer les ressources
            }
        }

        /// <summary>
        /// Nettoie une chaîne de caractères pour la rendre utilisable comme partie d'un nom de fichier.
        /// Supprime les accents, passe en minuscules, et remplace les caractères non alphanumériques par des underscores.
        /// </summary>
        public static string SanitizeForFilename(string input)
        {
            if (string.IsNullOrEmpty(input))
                return "chaine_vide"; // Retourne une chaîne par défaut si l'entrée est vide ou nulle

            // 1) Normalisation pour décomposer les caractères accentués (ex: "é" -> "e" + "'")
            string normalizedString = input.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(normalizedString.Length);
            foreach (char c in normalizedString)
            {
                // Conserver uniquement les caractères de base (lettres, chiffres), ignorer les marques diacritiques
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(c);
                }
            }
            // Recomposer la chaîne (FormC) après avoir retiré les accents, puis passage en minuscules.
            string lowerCaseString = sb.ToString().Normalize(NormalizationForm.FormC).ToLowerInvariant();

            // 2) Remplacement de tout ce qui n'est pas a-z ou 0-9 par '_'
            string replaced = Regex.Replace(lowerCaseString, @"[^a-z0-9]", "_", RegexOptions.Compiled);

            // 3) Écrasement des '_' multiples
            string cleaned = Regex.Replace(replaced, @"_+", "_", RegexOptions.Compiled);

            // 4) Trim des '_' en début/fin
            cleaned = cleaned.Trim('_');

            // 5) Si la chaîne est vide après nettoyage (ex: "!!!"), retourner une valeur par défaut.
            if (string.IsNullOrEmpty(cleaned))
                return "nettoyage_vide";

            return cleaned;
        }
    } // Fin de la classe StudentPdfSplitter
} // Fin du namespace aposplit
