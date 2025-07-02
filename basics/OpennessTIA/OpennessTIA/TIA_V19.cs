using Microsoft.Win32;
using Siemens.Engineering;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace OpennessTIA
{
    public class TIA_V19
    {
        public static TiaPortal instTIA;
        public Project CurrentProject { get; private set; } // Added property to hold the opened project

        /// <summary>
        /// Open a new instance of TIA portal with / without user interface
        /// </summary>
        /// <param name="guiTIA"></param>
        public void CreateTIAinstance(bool guiTIA)
        {
            // set whitelist entry
            SetWhitelist(System.Diagnostics.Process.GetCurrentProcess().ProcessName, System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);
            // open new tia instance with user interface
            if (guiTIA)
            {
                instTIA = new TiaPortal(TiaPortalMode.WithUserInterface);
            }
            // open tia instance without user interface
            else
            {
                instTIA = new TiaPortal(TiaPortalMode.WithoutUserInterface);
            }
        }

        /// <summary>
        /// Opens an existing TIA Portal project.
        /// </summary>
        /// <param name="projectFilePath">The full path to the .ap19 project file.</param>
        /// <returns>The opened Project object, or null if an error occurred.</returns>
        public Project OpenProject(string projectFilePath)
        {
            if (instTIA == null)
            {
                Console.WriteLine("TIA Portal instance is not created. Call CreateTIAinstance() first.");
                return null;
            }

            try
            {
                FileInfo projectFile = new FileInfo(projectFilePath);

                if (!projectFile.Exists)
                {
                    Console.WriteLine($"Error: Project file not found at '{projectFilePath}'");
                    return null;
                }

                Console.WriteLine($"Attempting to open project: {projectFile.Name}...");
                CurrentProject = instTIA.Projects.Open(projectFile);
                Console.WriteLine($"Successfully opened project: {CurrentProject.Name}");
                return CurrentProject;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to open project: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
                }
                CurrentProject = null; // Ensure CurrentProject is null on failure
                return null;
            }
        }


        /// <summary>
        /// set whitelist for tia portal registry
        /// </summary>
        /// <param name="ApplicationName"></param>
        /// <param name="ApplicationStartupPath"></param>
        static void SetWhitelist(string ApplicationName, string ApplicationStartupPath)
        {

            RegistryKey key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            RegistryKey software = null;
            try
            {
                software = key.OpenSubKey(@"SOFTWARE\Siemens\Automation\Openness")
                    .OpenSubKey("19.0") // This matches your TIA_V19 class name, assuming TIA Portal V19
                    .OpenSubKey("Whitelist")
                    .OpenSubKey(ApplicationName + ".exe")
                    .OpenSubKey("Entry", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryRights.FullControl);
            }
            catch (Exception)
            {

                //Eintrag in der Whitelist ist nicht vorhanden
                //Entry in whitelist is not available
                software = key.CreateSubKey(@"SOFTWARE\Siemens\Automation\Openness")
                    .CreateSubKey("19.0") // This matches your TIA_V19 class name, assuming TIA Portal V19
                    .CreateSubKey("Whitelist")
                    .CreateSubKey(ApplicationName + ".exe")
                    .CreateSubKey("Entry", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryOptions.None);
            }


            string lastWriteTimeUtcFormatted = String.Empty;
            DateTime lastWriteTimeUtc;
            HashAlgorithm hashAlgorithm = SHA256.Create();
            FileStream stream = File.OpenRead(ApplicationStartupPath);
            byte[] hash = hashAlgorithm.ComputeHash(stream);
            // this is how the hash should appear in the .reg file
            string convertedHash = Convert.ToBase64String(hash);
            software.SetValue("FileHash", convertedHash);
            lastWriteTimeUtc = new FileInfo(ApplicationStartupPath).LastWriteTimeUtc;
            // this is how the last write time should be formatted
            lastWriteTimeUtcFormatted = lastWriteTimeUtc.ToString(@"yyyy/MM/dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
            software.SetValue("DateModified", lastWriteTimeUtcFormatted);
            software.SetValue("Path", ApplicationStartupPath);
        }
    }
}