using System;
using System.Windows.Forms;
using AutoUpdaterDotNET;

namespace aposplit
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                // Configuration de l'auto-updater
                AutoUpdater.ReportErrors = false;
                AutoUpdater.Synchronous = false;
                AutoUpdater.Mandatory = false;

                AutoUpdater.CheckForUpdateEvent += (args) =>
                {
                    if (args.Error != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"AutoUpdater error: {args.Error.Message}");
                        return;
                    }

                    if (args.IsUpdateAvailable)
                    {
                        // On fait l'appel UI dans le thread principal
                        if (Application.OpenForms.Count > 0)
                        {
                            var mainForm = Application.OpenForms[0];
                            mainForm.Invoke(new Action(() =>
                            {
                                AutoUpdater.ShowUpdateForm(args);
                            }));
                        }
                        else
                        {
                            // Si aucun formulaire n'est encore ouvert, fallback : téléchargement direct
                            AutoUpdater.DownloadUpdate(args);
                        }
                    }
                    // sinon : ne rien afficher
                };

                // URL vers le manifest JSON (branche main > updates/aposplit.xml)
                AutoUpdater.Start("https://raw.githubusercontent.com/brahimmachkouri/aposplit/main/updates/aposplit.xml");
            }
            catch (Exception ex)
            {
                // En cas d'erreur réseau ou parsing, on ignore pour ne pas bloquer l'appli
                System.Diagnostics.Debug.WriteLine($"AutoUpdater failed: {ex.Message}");
            }
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }
    }
}
