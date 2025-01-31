using System;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace SlideMasterRemover
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("Usage: PowerPointMasterCleanup <path to PowerPoint file>");
                return 1;
            }

            string pptFilePath = args[0];

            if (!File.Exists(pptFilePath))
            {
                Console.WriteLine($"Error: The specified file \"{pptFilePath}\" does not exist.");
                return 1;
            }

            string ext = Path.GetExtension(pptFilePath).ToLower();
            if (ext != ".pptx" && ext != ".pptm" && ext != ".ppt")
            {
                Console.WriteLine($"Warning: The extension \"{ext}\" may not be supported. Continue? (Y/N)");
                var key = Console.ReadKey(intercept: true);
                Console.WriteLine();
                if (key.KeyChar != 'Y' && key.KeyChar != 'y')
                {
                    return 1;
                }
            }

            Application pptApp = null;
            Presentation presentation = null;

            try
            {
                pptApp = new Application();
                presentation = pptApp.Presentations.Open(
                    pptFilePath,
                    ReadOnly: MsoTriState.msoFalse,
                    Untitled: MsoTriState.msoFalse,
                    WithWindow: MsoTriState.msoFalse
                );

                RemoveUnusedDesigns(presentation);
                presentation.Save();

                Console.WriteLine("Unused slide masters have been removed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                return 1;
            }
            finally
            {
                if (presentation != null)
                {
                    presentation.Close();
                    Marshal.ReleaseComObject(presentation);
                }
                if (pptApp != null)
                {
                    pptApp.Quit();
                    Marshal.ReleaseComObject(pptApp);
                }
            }

            return 0;
        }

        static void RemoveUnusedDesigns(Presentation presentation)
        {
            var usedLayouts = new HashSet<CustomLayout>();

            foreach (Slide slide in presentation.Slides)
            {
                usedLayouts.Add(slide.CustomLayout);
            }

            for (int i = presentation.Designs.Count; i >= 1; i--)
            {
                Design design = presentation.Designs[i];
                bool isDesignUsed = false;

                for (int j = design.SlideMaster.CustomLayouts.Count; j >= 1; j--)
                {
                    CustomLayout layout = design.SlideMaster.CustomLayouts[j];
                    if (usedLayouts.Contains(layout))
                    {
                        isDesignUsed = true;
                        Console.WriteLine($"Layout \"{layout.Name}\" is used.");
                    }
                    else
                    {
                        Console.WriteLine($"Delete layout: {layout.Name}");
                        layout.Delete();
                    }
                }

                if (!isDesignUsed)
                {
                    Console.WriteLine($"Delete design: {design.Name}");
                    design.Delete();
                }
            }
        }
    }
}
