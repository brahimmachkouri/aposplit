// (c) Brahim MACHKOURI 2025
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
    /// Les fichiers sont créés dans un sous-dossier nommé <nom_original_par_etudiant>.
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
            string studentOutputDir = Path.Combine(outputBaseDir, $"{sanitizedBaseName}_par_etudiant");

            Directory.CreateDirectory(studentOutputDir);

            var pdfHandler = new PdfHandler();
            ExtractAndSaveStudentDocuments(pdfPath, pdfHandler, studentOutputDir);
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
        /// Extrait les pages des relevés de notes de chaque étudiant à partir d’un fichier PDF source, 
        /// et enregistre chaque relevé individuel dans un fichier PDF distinct.
        /// </summary>
        /// <param name="pdfPath">Chemin vers le fichier PDF source contenant tous les relevés de notes.</param>
        /// <param name="pdfHandler">Instance de gestionnaire PDF permettant de lire et manipuler les pages du document.</param>
        /// <param name="outputDir">Chemin du répertoire dans lequel les fichiers PDF des étudiants seront enregistrés.</param>
        /// <remarks>
        /// Chaque étudiant est identifié par un numéro étudiant extrait du texte de la page. Lorsqu’un changement d’étudiant est détecté,
        /// les pages précédentes sont enregistrées dans un nouveau fichier PDF, et un nouveau document est créé.
        /// </remarks>
        /// <returns>Ne retourne rien mais affiche dans la console le nombre de fichiers PDF créés à la fin du traitement.</returns>
        private static void ExtractAndSaveStudentDocuments(
    string pdfPath, IPdfHandler pdfHandler, string outputDir)
        {
            StudentPageInfo? currentStudentInfo = null;
            PdfSharp.Pdf.PdfDocument? currentStudentPdf = null;
            int savedCount = 0;

            int pageCount = pdfHandler.GetPageCount(pdfPath);
            for (int i = 1; i < pageCount; i++)
            {
                string pageText = pdfHandler.GetPageText(pdfPath, i);
                StudentPageInfo extractedInfo = ParseStudentInfoFromText(pageText);

                bool isNewStudent = currentStudentInfo == null ||
                                    (extractedInfo.StudentNumber != null && extractedInfo.StudentNumber != currentStudentInfo.StudentNumber) ||
                                    (extractedInfo.StudentNumber == null && currentStudentInfo.StudentNumber != null);

                if (isNewStudent)
                {
                    // Sauvegarder le PDF précédent avant de passer au suivant
                    if (currentStudentPdf != null && currentStudentInfo != null && currentStudentPdf.PageCount > 0)
                    {
                        SaveCurrentStudentPdf(currentStudentPdf, currentStudentInfo, outputDir);
                        savedCount++;
                        currentStudentPdf.Dispose(); // Libérer la mémoire
                    }

                    currentStudentInfo = extractedInfo;
                    currentStudentPdf = new PdfSharp.Pdf.PdfDocument();
                }

                pdfHandler.AddPageToDocument(currentStudentPdf!, pdfPath, i);
            }

            // Sauvegarder le dernier étudiant
            if (currentStudentPdf != null && currentStudentInfo != null && currentStudentPdf.PageCount > 0)
            {
                SaveCurrentStudentPdf(currentStudentPdf, currentStudentInfo, outputDir);
                savedCount++;
                currentStudentPdf.Dispose();
            }

            Console.WriteLine($"Traitement terminé : {savedCount} fichiers créés.");
        }

        /// <summary>
        /// Enregistre le document PDF d’un étudiant dans le répertoire de sortie spécifié, 
        /// en générant un nom de fichier sûr à partir de son nom et de son numéro d'étudiant.
        /// </summary>
        /// <param name="pdf">Document PDF à sauvegarder, correspondant à un étudiant.</param>
        /// <param name="info">Informations de l'étudiant (nom, numéro étudiant, etc.).</param>
        /// <param name="outputDir">Chemin absolu vers le répertoire dans lequel le fichier PDF sera enregistré.</param>
        /// <remarks>
        /// Le nom du fichier est généré sous la forme : {NomEtudiant}_{NumeroEtudiant}.pdf,
        /// après nettoyage des caractères potentiellement invalides pour le système de fichiers.
        /// </remarks>
        private static void SaveCurrentStudentPdf(
            PdfSharp.Pdf.PdfDocument pdf, StudentPageInfo info, string outputDir)
        {
            string safeStudentName = SanitizeForFilename(info.Name ?? "NomNonTrouve");
            string safeStudentNumber = SanitizeForFilename(info.StudentNumber ?? "NumNonTrouve");
            string fileName = $"{safeStudentName}_{safeStudentNumber}.pdf";
            string filePath = Path.Combine(outputDir, fileName);

            pdf.Save(filePath);
            Console.WriteLine($"Sauvegardé : {fileName}");
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

        private static readonly Regex ReAttestationEdition = new Regex(@"edition\s*d['’]\s*attestations\s*de\s*r[ée]ussite", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        //private static bool IsAttestationPage(Page page) => ReAttestationEdition.IsMatch(Normalize(page.Text ?? ""));
        private static readonly Regex ReNameLine = new Regex(@"(Monsieur|Madame)\s+(.+?)a\s+été\s+décern[ée]+\s+à", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        //private static readonly Regex ReFormationLine = new Regex(@"\bsp[ée]cialit[ée]\s+(.+?)\s*parcours", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex ReNumEtuAttestation = new Regex(@"\)(\d+)N°\s*[ée]tudiant\s*:", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static Tuple<string, string> ParseMeta(Page page)
        {
            string nom = "";
            string ine = "";
            string fullText = page.Text ?? "";

            // Recherche du nom dans tout le texte
            var nameMatch = ReNameLine.Match(fullText);
            if (nameMatch.Success)
            {
                var fullName = nameMatch.Groups[2].Value.Trim();

                // On coupe seulement au premier espace
                var firstSpaceIndex = fullName.IndexOf(' ');
                if (firstSpaceIndex > 0)
                {
                    nom = Clean(fullName.Substring(firstSpaceIndex + 1).Trim());
                }
                else
                {
                    // Cas particulier : un seul mot
                    nom = fullName;
                }
            }

            // Recherche du numéro étudiant dans tout le texte
            var ineMatch = ReNumEtuAttestation.Match(fullText);
            if (ineMatch.Success)
            {
                ine = Clean(ineMatch.Groups[1].Value);
            }


            // Recherche du numéro étudiant ligne par ligne
            var lines = fullText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Select(l => l.Trim());
            foreach (var line in lines)
            {
                if (ine == "")
                {
                    var m = ReNumEtuAttestation.Match(line);
                    if (m.Success)
                    {
                        var tokens = m.Groups[2].Value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        if (tokens.Length > 0) ine = Clean(tokens[tokens.Length - 1]);
                    }
                }

                if (ine != "") break;
            }

            // Valeurs par défaut
            if (nom == "") nom = "nan";
            if (ine == "") ine = "nan";

            Console.WriteLine($"Résultat final: {nom}, {ine}");
            return Tuple.Create(
                Normalize(nom),
                Normalize(ine)
            );
        }

        // ------------------------- Helpers utils ---------------------------------
        private static string Normalize(string s)
        {
            var d = s.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(d.Length);
            foreach (var c in d)
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                    sb.Append(c);
            return sb.ToString().Normalize(NormalizationForm.FormC).ToLowerInvariant();
        }

        private static string Clean(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "x";
            foreach (var c in Path.GetInvalidFileNameChars()) s = s.Replace(c, '_');
            s = Regex.Replace(s, @"[^\w\-]", "_");
            s = Regex.Replace(s, "_+", "_").Trim('_');
            return s == "" ? "x" : s;
        }

        private static void SplitByAttestation(string pdfPath)
        {
            string outDir = Path.Combine(Path.GetDirectoryName(pdfPath) ?? Environment.CurrentDirectory, $"{Path.GetFileNameWithoutExtension(pdfPath)}_attestations");
            Directory.CreateDirectory(outDir);

            UglyToad.PdfPig.PdfDocument pdfDocument = UglyToad.PdfPig.PdfDocument.Open(pdfPath);
            UglyToad.PdfPig.PdfDocument pig = pdfDocument;
            PdfSharp.Pdf.PdfDocument sharp = PdfSharp.Pdf.IO.PdfReader.Open(pdfPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);

            int saved = 0;
            for (int i = 0; i < pig.NumberOfPages; i++)
            {
                if (i > 0)
                {
                    var pagePig = pig.GetPage(i + 1);
                    var (nom, ine) = ParseMeta(pagePig);
                    string fileName = $"{nom}-{ine}.pdf";
                    string fullPath = Path.Combine(outDir, fileName);

                    using (var single = new PdfSharp.Pdf.PdfDocument())
                    {
                        single.AddPage(sharp.Pages[i]); // PdfSharp pages = 0-based
                        single.Save(fullPath);
                    }

                    Console.WriteLine($"✓ Page {i + 1} → {fileName}");
                    saved++;
                }
            }

            Console.WriteLine(saved > 0
                ? $"Terminé : {saved} attestation(s) enregistrée(s) dans « {outDir} »."
                : "Aucune page d’attestation détectée.");
        }

        /// <summary>
        /// Détecte automatiquement le type de document (relevé de notes ou attestation)
        /// et effectue la séparation adaptée.
        /// </summary>
        /// <param name="pdfPath">Chemin du fichier PDF source</param>
        /// <param name="outputBaseDir">Répertoire de sortie</param>
        public static void AutoSplit(string pdfPath, string outputBaseDir)
        {
            if (string.IsNullOrEmpty(pdfPath)) throw new ArgumentNullException(nameof(pdfPath));
            if (!File.Exists(pdfPath)) throw new FileNotFoundException("Le fichier PDF source n'a pas été trouvé.", pdfPath);

            var pdfHandler = new PdfHandler();
            // On lit le texte de la première page pour détecter le type de document
            string firstPageText = pdfHandler.GetPageText(pdfPath, 0);

            // On utilise la regex d'attestation déjà présente
            if (ReAttestationEdition.IsMatch(Normalize(firstPageText)))
            {
                SplitByAttestation(pdfPath);
            }
            else
            {
                // Par défaut, on considère qu'il s'agit d'un relevé de notes
                SplitByStudent(pdfPath, outputBaseDir);
            }
        }

    }
}
