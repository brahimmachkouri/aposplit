// (c) Brahim MACHKOURI 2025 — version simplifiée

using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using UglyToad.PdfPig;

namespace aposplit
{
    // Session légère qui combine PdfPig (texte) + PdfSharp (import de pages)
    internal sealed class PdfSession : IDisposable
    {
        private readonly UglyToad.PdfPig.PdfDocument _pig;
        private readonly PdfSharp.Pdf.PdfDocument _sharp;

        public PdfSession(string path)
        {
            _pig = UglyToad.PdfPig.PdfDocument.Open(path, new ParsingOptions
            {
                UseLenientParsing = true,
                SkipMissingFonts = true
            });
            _sharp = PdfReader.Open(path, PdfDocumentOpenMode.Import); // Import (pas Modify)
        }

        public int PageCount => _sharp.PageCount; // ou _pig.NumberOfPages

        public string GetPageText(int pageIndex0) => _pig.GetPage(pageIndex0 + 1).Text ?? string.Empty; // PdfPig = 1‑based

        public void AddPageToDocument(PdfSharp.Pdf.PdfDocument dst, int pageIndex0) => dst.AddPage(_sharp.Pages[pageIndex0]);

        public void Dispose()
        {
            _sharp?.Dispose();
            _pig?.Dispose();
        }
    }

    /// <summary>
    /// Divise un PDF de relevés/attestations en fichiers individuels.
    /// </summary>
    public static class StudentPdfSplitter
    {
        // ----------- Constantes & Regex -----------
        private const string NameLineMarker = "Page :/"; // marqueur de la ligne contenant NOM Prénom I
        private static readonly Regex ReNumEtu = new(@"N°\s*Etudiant\s*[:\-]?\s*(\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex ReAttestationEdition = new(@"edition\s*d['’]\s*attestations\s*de\s*r[ée]ussite", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex ReNameLine = new(@"(Monsieur|Madame)\s+(.+?)a\s+été\s+décern[ée]+\s+à", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex ReNumEtuAttestation = new(@"\)(\d+)N°\s*[ée]tudiant\s*:", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private record StudentPageInfo(string? Name, string? StudentNumber);

        /// <summary>
        /// Détecte automatiquement le type de document contenu dans <paramref name="pdfPath"/>
        /// (relevés de notes ou attestations) et lance le découpage approprié.
        /// </summary>
        /// <param name="pdfPath">Chemin vers le fichier PDF source. Doit exister.</param>
        /// <param name="outputBaseDir">
        /// Répertoire de base où les fichiers résultants seront écrits pour le mode
        /// « relevés par étudiant ». Doit être un chemin non vide.
        /// </param>
        public static void AutoSplit(string pdfPath, string outputBaseDir)
        {
            ValidateInputPaths(pdfPath, outputBaseDir);
            using var session = new PdfSession(pdfPath);

            string firstPageText = Normalize(session.GetPageText(0));
            if (ReAttestationEdition.IsMatch(firstPageText))
            {
                SplitByAttestation(pdfPath);
            }
            else
            {

                SplitByStudent(pdfPath, outputBaseDir);
            }
        }

        /// <summary>
        /// Prépare et lance le découpage du PDF en fichiers par étudiant.
        /// </summary>
        /// <param name="pdfPath">Chemin complet du fichier PDF source. Doit exister.</param>
        /// <param name="outputBaseDir">
        /// Répertoire de base dans lequel sera créé un sous-dossier nommé
        /// "{nom_du_fichier_sanitized}_par_etudiant" contenant les PDF par étudiant.
        /// </param>
        private static void SplitByStudent(string pdfPath, string outputBaseDir)
        {
            ValidateInputPaths(pdfPath, outputBaseDir);

            string baseName = Path.GetFileNameWithoutExtension(pdfPath);
            string studentDir = Path.Combine(outputBaseDir, $"{SanitizeForFilename(baseName)}_par_etudiant");
            Directory.CreateDirectory(studentDir);

            ExtractAndSaveStudentDocuments(pdfPath, studentDir);
        }

        // ----------- Cœur « relevés par étudiant » -----------
        /// <summary>
        /// Parcourt les pages du PDF source et crée un fichier PDF par étudiant.
        /// </summary>
        /// <param name="pdfPath">Chemin complet vers le fichier PDF source (par ex. un relevés de notes).</param>
        /// <param name="outputDir">
        /// Répertoire dans lequel seront écrits les fichiers PDF individuels.
        /// Le répertoire est créé par l'appelant (<see cref="SplitByStudent"/>) si nécessaire.
        /// </param>
        private static void ExtractAndSaveStudentDocuments(string pdfPath, string outputDir)
        {
            using var session = new PdfSession(pdfPath); // ouverture unique

            StudentPageInfo? current = null;
            PdfSharp.Pdf.PdfDocument? currentPdf = null;
            int savedCount = 0;

            for (int i = 1; i < session.PageCount; i++)
            {
                string pageText = session.GetPageText(i);
                var info = ParseStudentInfoFromText(pageText);

                bool isNewStudent = current == null || !string.Equals(info.StudentNumber, current.StudentNumber, StringComparison.Ordinal);
                if (isNewStudent)
                {
                    if (currentPdf is { PageCount: > 0 } && current != null)
                    {
                        SaveCurrentStudentPdf(currentPdf, current, outputDir);
                        savedCount++;
                        currentPdf.Dispose();
                    }

                    current = info;
                    currentPdf = new PdfSharp.Pdf.PdfDocument();
                    // Option CPU vs taille :
                    // currentPdf.Options.NoCompression = true;
                    // currentPdf.Options.CompressContentStreams = false;
                }

                session.AddPageToDocument(currentPdf!, i);
            }

            if (currentPdf is { PageCount: > 0 } && current != null)
            {
                SaveCurrentStudentPdf(currentPdf, current, outputDir);
                savedCount++;
                currentPdf.Dispose();
            }

            Console.WriteLine($"Traitement terminé : {savedCount} fichiers créés.");
        }

        /// <summary>
        /// Représente les informations extraites d'une page PDF relatives à un étudiant.
        /// </summary>
        /// <param name="Name">
        /// Nom (ou portion du nom) détecté sur la page. Peut être null si le nom n'a pas été trouvé.
        /// </param>
        /// <param name="StudentNumber">
        /// Numéro étudiant extrait de la page. Peut être null si non trouvé.
        /// </param>
        private static StudentPageInfo ParseStudentInfoFromText(string pageText)
        {
            var lines = pageText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            string? name = null;
            string? number = null;

            foreach (var raw in lines)
            {
                var line = raw.Trim();

                // Nom à partir de la ligne contenant le marqueur « Page :/ »
                if (name == null)
                {
                    int idx = line.IndexOf(NameLineMarker, StringComparison.OrdinalIgnoreCase);
                    if (idx >= 0)
                    {
                        string tail = line[(idx + NameLineMarker.Length)..].TrimStart();
                        var sb = new StringBuilder();
                        foreach (char c in tail)
                        {
                            if (char.IsUpper(c) || char.IsWhiteSpace(c) || c == '\'' || c == '-') sb.Append(c); else break;
                        }
                        string candidate = sb.ToString().Trim();
                        if (!string.IsNullOrWhiteSpace(candidate))
                        {
                            int lastSpaceIdx = candidate.LastIndexOf(' ');
                            if (lastSpaceIdx > 0 && (candidate.Length - lastSpaceIdx - 1 == 1) && char.IsUpper(candidate[lastSpaceIdx + 1]))
                                name = candidate[..lastSpaceIdx].Trim();
                            else
                                name = candidate;
                        }
                        if (string.IsNullOrWhiteSpace(name)) name = null;
                    }
                }

                // Numéro étudiant
                if (number == null)
                {
                    var m = ReNumEtu.Match(line);
                    if (m.Success) number = m.Groups[1].Value;
                }

                if (name != null && number != null) break;
            }
            return new StudentPageInfo(name, number);
        }

        /// <summary>
        /// Sauvegarde le document PDF accumulé pour un étudiant dans le répertoire spécifié.
        /// </summary>
        /// <param name="pdf">Document PdfSharp contenant les pages de l'étudiant. Ne doit pas être null.</param>
        /// <param name="info">Informations extraites de la page (nom et numéro). Les valeurs null sont remplacées par des valeurs par défaut.</param>
        /// <param name="outputDir">Répertoire de sortie où sera écrit le fichier. Doit exister ou être géré par l'appelant.</param>
        private static void SaveCurrentStudentPdf(PdfSharp.Pdf.PdfDocument pdf, StudentPageInfo info, string outputDir)
        {
            string safeName = SanitizeForFilename(info.Name ?? "NomNonTrouve");
            string safeNum = SanitizeForFilename(info.StudentNumber ?? "NumNonTrouve");
            string fileName = $"{safeName}_{safeNum}.pdf";
            string path = Path.Combine(outputDir, fileName);
            pdf.Save(path);
            Console.WriteLine($"Sauvegardé : {fileName}");
        }

        // ----------- Mode « attestations » -----------
        /// <summary>
        /// Extrait et enregistre chaque page d'attestation comme fichier PDF individuel.
        /// </summary>
        /// <param name="pdfPath">Chemin complet du fichier PDF source contenant les attestations.</param>
        private static void SplitByAttestation(string pdfPath)
        {
            string outDir = Path.Combine(Path.GetDirectoryName(pdfPath) ?? Environment.CurrentDirectory,
                                         $"{Path.GetFileNameWithoutExtension(pdfPath)}_attestations");
            Directory.CreateDirectory(outDir);

            using var session = new PdfSession(pdfPath);
            int saved = 0;

            for (int i = 0; i < session.PageCount; i++)
            {
                if (i > 0) // on conserve le « skip page 1 » tel quel
                {
                    string fullText = session.GetPageText(i);
                    var (nom, ine) = ParseMeta(fullText);
                    string fileName = $"{nom}-{ine}.pdf";
                    string fullPath = Path.Combine(outDir, fileName);

                    using var single = new PdfSharp.Pdf.PdfDocument();
                    session.AddPageToDocument(single, i);
                    single.Save(fullPath);

                    Console.WriteLine($"✓ Page {i + 1} → {fileName}");
                    saved++;
                }
            }

            Console.WriteLine(saved > 0
                ? $"Terminé : {saved} attestation(s) enregistrée(s) dans « {outDir} »."
                : "Aucune page d’attestation détectée.");
        }

        /// <summary>
        /// Extrait le nom et le numéro étudiant (INE / numéro) depuis le texte complet d'une page d'attestation.
        /// </summary>
        /// <param name="fullText">Texte complet de la page (peut contenir plusieurs lignes).</param>
        /// <returns>
        /// Tuple (nom, ine) :
        /// - nom : valeur nettoyée et normalisée extraite via <see cref="ReNameLine"/> (ou "nan" si introuvable).
        /// - ine : numéro étudiant extrait via <see cref="ReNumEtuAttestation"/> (ou "nan" si introuvable).
        /// </returns>
        private static (string nom, string ine) ParseMeta(string fullText)
        {
            string nom = "";
            string ine = "";

            // Nom sur tout le texte
            var nameMatch = ReNameLine.Match(fullText);
            if (nameMatch.Success)
            {
                var fullName = nameMatch.Groups[2].Value.Trim();
                int firstSpace = fullName.IndexOf(' ');
                nom = firstSpace > 0 ? Clean(fullName[(firstSpace + 1)..].Trim()) : Clean(fullName);
            }

            // N° étudiant sur tout le texte
            var ineMatch = ReNumEtuAttestation.Match(fullText);
            if (ineMatch.Success)
            {
                ine = Clean(ineMatch.Groups[1].Value);
            }

            if (string.IsNullOrWhiteSpace(nom)) nom = "nan";
            if (string.IsNullOrWhiteSpace(ine)) ine = "nan";

            return (Normalize(nom), Normalize(ine));
        }

        // ----------- Helpers -----------

        /// <summary>
        /// Valide les chemins utilisés par les opérations de découpage.
        /// </summary>
        /// <param name="pdfPath">Chemin complet du fichier PDF source. Ne doit pas être null/empty et doit pointer vers un fichier existant.</param>
        private static void ValidateInputPaths(string pdfPath, string outputBaseDir)
        {
            if (string.IsNullOrEmpty(pdfPath)) throw new ArgumentNullException(nameof(pdfPath));
            if (!File.Exists(pdfPath)) throw new FileNotFoundException("Le fichier PDF source n'a pas été trouvé.", pdfPath);
            if (string.IsNullOrEmpty(outputBaseDir)) throw new ArgumentNullException(nameof(outputBaseDir));
        }

        /// <summary>
        /// Nettoie et normalise une chaîne pour qu'elle soit utilisable comme nom de fichier.
        /// </summary>
        /// <param name="input">Chaîne d'entrée (par ex. nom d'étudiant ou titre de fichier).</param>
        /// <returns>
        /// Chaîne sécurisée contenant uniquement des caractères ASCII minuscules [a‑z0‑9] et des underscores.
        /// - Supprime les diacritiques (normalisation Unicode FormD + suppression des NonSpacingMark).
        /// - Recompose en FormC puis transforme en minuscules invariant culture.
        /// - Remplace tout caractère non [a‑z0‑9] par '_' et réduit les suites de '_' à un seul '_'.
        /// - Supprime les '_' en début/fin.
        /// - Si le résultat est vide, retourne "nettoyage_vide". Si l'entrée est null/empty, retourne "chaine_vide".
        /// </returns>
        private static string SanitizeForFilename(string input)
        {
            if (string.IsNullOrEmpty(input)) return "chaine_vide";

            string normalized = input.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(normalized.Length);
            foreach (char c in normalized)
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark) sb.Append(c);

            string s = sb.ToString().Normalize(NormalizationForm.FormC).ToLowerInvariant();
            s = Regex.Replace(s, @"[^a-z0-9]", "_", RegexOptions.Compiled);
            s = Regex.Replace(s, @"_+", "_", RegexOptions.Compiled).Trim('_');
            return string.IsNullOrEmpty(s) ? "nettoyage_vide" : s;
        }

        /// <summary>
        /// Normalise une chaîne en supprimant les diacritiques et en la convertissant en minuscules.
        /// </summary>
        /// <param name="s">Chaîne d'entrée (doit être non null). La méthode décompose les caractères Unicode, supprime
        /// les marques diacritiques (NonSpacingMark) puis recombine la chaîne.</param>
        /// <returns>
        /// Chaîne recomposée en NormalizationForm.FormC et convertie en minuscules invariant culture,
        /// sans les marques diacritiques.
        /// </returns>
        private static string Normalize(string s)
        {
            var d = s.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(d.Length);
            foreach (var c in d)
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark) sb.Append(c);
            return sb.ToString().Normalize(NormalizationForm.FormC).ToLowerInvariant();
        }

        /// <summary>
        /// Nettoie une chaîne pour la rendre sûre en tant que composant de nom de fichier.
        /// </summary>
        /// <param name="s">Chaîne d'entrée (peut contenir espaces, caractères invalides, accents, etc.).</param>
        /// <returns>
        /// Chaîne nettoyée :
        /// - tous les caractères invalides pour un nom de fichier sont remplacés par '_',
        /// - les séquences de '_' sont réduites à un seul '_',
        /// - les '_' en début/fin sont supprimés,
        /// - si l'entrée est vide ou ne contient aucun caractère valide, retourne "x".
        /// </returns>
        private static string Clean(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "x";
            foreach (var c in Path.GetInvalidFileNameChars()) s = s.Replace(c, '_');
            s = Regex.Replace(s, @"[^\w\-]", "_");
            s = Regex.Replace(s, "_+", "_").Trim('_');
            return s == "" ? "x" : s;
        }
    }
}
