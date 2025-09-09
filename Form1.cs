using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace aposplit
{
    public partial class Form1 : Form
    {
        private bool _busy;
        private static readonly string[] AllowedExts = { ".pdf" };
        private static readonly TimeSpan StatusBarDelay = TimeSpan.FromMilliseconds(1000);
        private const string PdfOnlyMessage = "Seuls les fichiers PDF sont acceptés.";

        public Form1()
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;

            AllowDrop = true;
            DragLeave += OnDragLeave;
        }

        private void Form1_Load(object? sender, EventArgs e)
        {
            lblFichier.Text = string.Empty;
            var version = $"v{Application.ProductVersion}";
            var firstSpaceIndex = version.IndexOf('+');
            if (firstSpaceIndex > 0)
            {
                version = version.Substring(0, firstSpaceIndex).Trim();
            }
            lblVersion.Text = version;
            ResetStatusBar();
        }

        private void Form1_DragEnter(object? sender, DragEventArgs e)
        {
            ResetStatusBar();

            if (e.Data?.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
            {
                bool allPdf = files.All(filePath =>
                    !string.IsNullOrEmpty(filePath) &&
                    AllowedExts.Contains(Path.GetExtension(filePath), StringComparer.OrdinalIgnoreCase));

                if (allPdf)
                {
                    e.Effect = DragDropEffects.Copy;
                    return;
                }
            }

            e.Effect = DragDropEffects.None;
            toolTip1.Show(
                PdfOnlyMessage,
                this,
                PointToClient(Cursor.Position),
                (int)StatusBarDelay.TotalMilliseconds
            );
        }

        private async void Form1_DragDrop(object? sender, DragEventArgs e)
        {
            if (_busy) return;

            if (e.Data?.GetData(DataFormats.FileDrop) is not string[] files || files.Length == 0)
            {
                UpdateStatusBar(false);
                await DelayAndResetAsync();
                return;
            }

            string pdfFilePath = files[0];
            if (string.IsNullOrEmpty(pdfFilePath))
            {
                Debug.WriteLine("Un chemin de fichier vide ou nul a été reçu de DragDrop.");
                UpdateStatusBar(false);
                await DelayAndResetAsync();
                return;
            }

            _busy = true;
            AllowDrop = false;
            Cursor = Cursors.WaitCursor;
            ResetStatusBar();

            lblFichier.Text = Path.GetFileName(pdfFilePath) ?? string.Empty;

            try
            {
                string? outputDirectory = Path.GetDirectoryName(pdfFilePath);
                if (string.IsNullOrEmpty(outputDirectory))
                {
                    throw new InvalidOperationException($"Impossible de déterminer le dossier de sortie pour '{pdfFilePath}'.");
                }

                //await Task.Run(() => StudentPdfSplitter.SplitByStudent(pdfFilePath, outputDirectory));
                await Task.Run(() => StudentPdfSplitter.AutoSplit(pdfFilePath, outputDirectory));
                UpdateStatusBar(true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Erreur lors du traitement de '{pdfFilePath}': {ex}");
                MessageBox.Show(this, $"Une erreur est survenue lors du traitement du fichier :\n{ex.Message}", "Erreur de traitement", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatusBar(false);
            }
            finally
            {
                await DelayAndResetAsync();
            }
        }

        private void OnDragLeave(object? sender, EventArgs e)
        {
            Cursor = Cursors.Default;
            toolTip1.Hide(this);
        }

        private void ResetStatusBar()
        {
            pnlStatusBar.BackColor = SystemColors.Control;
        }

        private void UpdateStatusBar(bool success)
        {
            pnlStatusBar.BackColor = success ? Color.Green : Color.Red;
        }

        private async Task DelayAndResetAsync()
        {
            await Task.Delay(StatusBarDelay);
            ResetStatusBar();
            ResetUI();
        }

        private void ResetUI()
        {
            Cursor = Cursors.Default;
            AllowDrop = true;
            _busy = false;
            lblFichier.Text = string.Empty;
        }

       
    }
}
