using System;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace WorkbookUpdate
{
    class Program
    {
        static void Main(string[] args)
        {
            foreach (string arg in args)
            {
                if (arg.Length > 4 && arg.Substring(arg.Length - 4, 4).ToLower().Equals("xlsx") && File.Exists(arg))
                {
                    var fiCopy = new FileInfo(Guid.NewGuid() + ".xlsx");
                    File.Copy(arg, fiCopy.FullName);
                    using (Package package = Package.Open(fiCopy.FullName, FileMode.Open, FileAccess.ReadWrite))
                    {
                        OpenSettings openSettings = new OpenSettings();
                        openSettings.MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2013);
                        using (SpreadsheetDocument doc = SpreadsheetDocument.Open(package, openSettings))
                        {
                            int count = 0;
                            OpenXmlValidator validator = new OpenXmlValidator(FileFormatVersions.Office2013);
                            foreach (var error in validator.Validate(doc))
                            {
                                count++;
                                Debug.WriteLine("Error " + count);
                                Debug.WriteLine("Description: " + error.Description);
                                Debug.WriteLine("ErrorType: " + error.ErrorType);
                                Debug.WriteLine("Node: " + error.Node);
                                Debug.WriteLine("Path: " + error.Path.XPath);
                                Debug.WriteLine("Part: " + error.Part.Uri);
                                Debug.WriteLine("-------------------------------------------");
                            }

                            Debug.WriteLine("count={0}", count);
                        }
                        fiCopy.Delete();
                    }                    
                }
            }
        }
    }
}
