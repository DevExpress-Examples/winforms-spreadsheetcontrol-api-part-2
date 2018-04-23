using System;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using System.Diagnostics;

namespace SpreadsheetControl_API
{
    public partial class Form1 : Form
    {

        IWorkbook workbook;

        public Form1()
        {
            InitializeComponent();

            // Access a workbook.
            workbook = spreadsheetControl1.Document;

            InitTreeListControl();

        }

        private void InitTreeListControl()
        {
            GroupsOfSpreadsheetExamples examples = new GroupsOfSpreadsheetExamples();
            InitData(examples);
            DataBinding(examples);
        }

        private void InitData(GroupsOfSpreadsheetExamples examples)
        {
            #region GroupNodes
            examples.Add(new SpreadsheetNode("Shapes"));
            examples.Add(new SpreadsheetNode("Custom Functions"));
            examples.Add(new SpreadsheetNode("Tables"));
            examples.Add(new SpreadsheetNode("Protection"));
            examples.Add(new SpreadsheetNode("Sort"));
            examples.Add(new SpreadsheetNode("Search"));
            examples.Add(new SpreadsheetNode("Export"));
            #endregion

            #region ExampleNodes
 
            // Add nodes to the "Shapes" group of examples.
            examples[0].Groups.Add(new SpreadsheetExample("Insert picture", ShapeActions.InsertShapeAction));
            examples[0].Groups.Add(new SpreadsheetExample("Insert picture from URI", ShapeActions.InsertShapeFromUriAction));
            examples[0].Groups.Add(new SpreadsheetExample("Modify picture", ShapeActions.ModifyShapeAction));
            
            // Add nodes to the "Custom Functions" group of examples.
            examples[1].Groups.Add(new SpreadsheetExample("Add UDF(user defined function)", CustomFunctionActions.SphereMassAction ));

            // Add nodes to the "Tables" group of examples.
            examples[2].Groups.Add(new SpreadsheetExample("Create table", TableActions.CreateTableAction));
            examples[2].Groups.Add(new SpreadsheetExample("Apply custom style", TableActions.CustomTableStyleAction));

            // Add nodes to the "Protection" group of examples.
            examples[3].Groups.Add(new SpreadsheetExample("Protect workbook", ProtectionActions.ProtectWorkbookAction));
            examples[3].Groups.Add(new SpreadsheetExample("Protect worksheet", ProtectionActions.ProtectWorksheetAction));
            examples[3].Groups.Add(new SpreadsheetExample("Protect range", ProtectionActions.ProtectRangeAction));

            // Add nodes to the "Sort" group of examples.
            examples[4].Groups.Add(new SpreadsheetExample("Simple sort", SortActions.SimpleSortAction));
            examples[4].Groups.Add(new SpreadsheetExample("Sort in descending order", SortActions.DescendingOrderAction));
            examples[4].Groups.Add(new SpreadsheetExample("Sort using custom comparer", SortActions.SelectComparerAction));
            examples[4].Groups.Add(new SpreadsheetExample("Sort by column", SortActions.SortBySpecifiedColumnAction));
            examples[4].Groups.Add(new SpreadsheetExample("Sort by multiple columns", SortActions.SortByMultipleColumnsAction));

            // Add nodes to the "Search" group of examples.
            examples[5].Groups.Add(new SpreadsheetExample("Simple search", SearchActions.SimpleSearchAction));
            examples[5].Groups.Add(new SpreadsheetExample("Search with options", SearchActions.AdvancedSearchAction));

            // Add nodes to the "Export" group of examples.
            examples[6].Groups.Add(new SpreadsheetExample("Export to HTML", ExportActions.ExportToHTMLAction));
            #endregion
        }

        void DataBinding(GroupsOfSpreadsheetExamples examples)
        {
            treeList1.DataSource = examples;
            treeList1.ExpandAll();
            treeList1.BestFitColumns();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            LoadDocumentFromFile();
            SpreadsheetExample example = treeList1.GetDataRecordByNode(treeList1.FocusedNode) as SpreadsheetExample;
            if (example == null)
                return;
            Action<IWorkbook> action = example.Action;
            action(workbook);
            this.spreadsheetControl1.Refresh();
            SaveDocumentToFile();
        }

        // ------------------- Load and Save a Document -------------------
        private void LoadDocumentFromFile() {
            #region #LoadDocumentFromFile
            // Load a workbook from the file.
            workbook.LoadDocument("Documents\\Document.xlsx", DocumentFormat.OpenXml);
            #endregion #LoadDocumentFromFile
        }

        private void SaveDocumentToFile() {
            #region #SaveDocumentToFile
            // Save the modified document to the file.
            workbook.SaveDocument("Documents\\SavedDocument.xlsx", DocumentFormat.OpenXml);
            #endregion #SaveDocumentToFile
        }
    }
}
