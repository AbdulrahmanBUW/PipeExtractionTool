using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;

namespace PipeExtractionTool
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Application : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                CreateRibbonTab(application);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to initialize Pipe Extraction Tool: {ex.Message}");
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        private void CreateRibbonTab(UIControlledApplication application)
        {
            string tabName = "Pipe Extraction";

            // Create ribbon tab
            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch (ArgumentException)
            {
                // Tab already exists, ignore
            }

            // Create ribbon panel
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Pipe Tools");

            // Get assembly path for button image
            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            string assemblyDir = Path.GetDirectoryName(assemblyPath);

            // Create push button data
            PushButtonData buttonData = new PushButtonData(
                "PipeExtractionCommand",
                "Extract Pipes\nto Excel",
                assemblyPath,
                "PipeExtractionTool.PipeExtractionCommand"
            );

            buttonData.ToolTip = "Extract pipe data from selected drawings to Excel file";
            buttonData.LongDescription = "This tool extracts all pipes with their SPEC_POSITION parameter values from selected drawings and exports them to an Excel file.";

            // Add button to panel
            PushButton pushButton = panel.AddItem(buttonData) as PushButton;

            // Set button image (32x32 for large icon)
            try
            {
                string imagePath = Path.Combine(assemblyDir, "Images", "pipe_icon_32.png");
                if (File.Exists(imagePath))
                {
                    pushButton.LargeImage = new BitmapImage(new Uri(imagePath));
                }
            }
            catch
            {
                // Image not found, use default
            }
        }
    }
}