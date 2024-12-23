
## Export DataTable to Excel using NPOI and MvvmLight

This guide will walk you through the process of exporting a `DataTable` to an Excel file using the NPOI library and following the MvvmLight framework conventions.

### Prerequisites

- Install the NuGet packages:
  - NPOI: Handles Excel file operations.
  - MvvmLight: Provides a simple MVVM pattern implementation for WPF applications.

### Step 1: Install NuGet Packages

Open your Package Manager Console and run the following commands:

```powershell
Install-Package NPOI
Install-Package MvvmLightLibs
```

### Step 2: Export excel

Create a ViewModel that will handle the logic for exporting the DataTable.

```csharp
            svar saveDialog = new SaveFileDialog
            {
                DefaultExt = ".xlsx",
                Title = "Save Excel (.xlsx)",
                Filter = "Excel Files|*.xlsx|All Files|*",
                CheckPathExists = true,
                OverwritePrompt = true,
                RestoreDirectory = true
            };
            if (saveDialog.ShowDialog() != true) return;

            // Create DataTable
            DataTable dt = CreateFakeDataTable(30);

            // Export to Excel file
            string filePath = "FakeData"; // You can modify the file path and name as needed
            NPOIHelper.ExportExcel(dt, saveDialog.FileName, filePath);
```

