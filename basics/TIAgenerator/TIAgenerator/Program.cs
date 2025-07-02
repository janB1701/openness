using OpennessTIA;
using Siemens.Engineering;
using Siemens.Engineering.HW; // <--- ADD THIS USING DIRECTIVE
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TIAgenerator
{ 
    internal class Program
    {
        static void Main(string[] args)
        {
            TIA_V19 tiaAutomation = new TIA_V19();

            try
            {
                tiaAutomation.CreateTIAinstance(true);

                string projectDirectory = @"C:\BMW_Daten\05E6GA125_20250610_1225";
                string projectName = "05E6GA125_20250610_1225.ap19";
                string fullProjectPath = Path.Combine(projectDirectory, projectName);

                Project myTiaProject = tiaAutomation.OpenProject(fullProjectPath);

                if (myTiaProject != null)
                {
                    Console.WriteLine($"Project '{myTiaProject.Name}' is now open and accessible.");



                    Console.WriteLine("\nDevices in the project:");
                    foreach (Device device in myTiaProject.Devices)
                    {
                        foreach (DeviceItem di in device.DeviceItems)
                        {
                            Console.WriteLine($"- {di.Name} nameeeeee");
                            
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Failed to open the TIA Portal project. Check console for details.");
                }

                Console.WriteLine("\nPress any key to exit and close TIA Portal instance...");
                Console.ReadKey();

                if (myTiaProject != null)
                {
                    myTiaProject.Close();
                    Console.WriteLine("Project closed.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred in Main: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
                }
            }
            finally
            {
                if (TIA_V19.instTIA != null)
                {
                    TIA_V19.instTIA.Dispose();
                    Console.WriteLine("TIA Portal instance disposed.");
                }
            }
        }
    }
}