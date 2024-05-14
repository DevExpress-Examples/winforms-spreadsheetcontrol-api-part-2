using System;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using System.Diagnostics;
using System.Net;

namespace SpreadsheetControl_API
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {

        IWorkbook workbook;

        public Form1()
        {
            System.Net.ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
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
            examples.Add(new SpreadsheetNode("Pictures"));
            examples.Add(new SpreadsheetNode("Custom Functions"));
            examples.Add(new SpreadsheetNode("Tables"));
            examples.Add(new SpreadsheetNode("Protection"));
            examples.Add(new SpreadsheetNode("Sort"));
            examples.Add(new SpreadsheetNode("Search"));
            examples.Add(new SpreadsheetNode("Export"));
            examples.Add(new SpreadsheetNode("Group Data"));
            examples.Add(new SpreadsheetNode("Filter Data"));
            examples.Add(new SpreadsheetNode("Document Properties"));
            #endregion

            #region ExampleNodes
 
            // Add nodes to the "Pictures" group of examples.
            examples[0].Groups.Add(new SpreadsheetExample("Insert picture", PictureActions.InsertPictureAction));
            examples[0].Groups.Add(new SpreadsheetExample("Insert picture from URI", PictureActions.InsertPictureFromUriAction));
            examples[0].Groups.Add(new SpreadsheetExample("Move picture", PictureActions.MovePictureAction));
            examples[0].Groups.Add(new SpreadsheetExample("Rotate picture", PictureActions.RotatePictureAction));
            examples[0].Groups.Add(new SpreadsheetExample("Bring picture to front", PictureActions.ChangeZOrderAction));
            examples[0].Groups.Add(new SpreadsheetExample("Add hyperlink", PictureActions.InsertHyperlinkAction));

            
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
            examples[4].Groups.Add(new SpreadsheetExample("Sort by fill", SortActions.SortByFillColorAction));
            examples[4].Groups.Add(new SpreadsheetExample("Sort by font color", SortActions.SortByFontColorAction));

            // Add nodes to the "Search" group of examples.
            examples[5].Groups.Add(new SpreadsheetExample("Simple search", SearchActions.SimpleSearchAction));
            examples[5].Groups.Add(new SpreadsheetExample("Search with options", SearchActions.AdvancedSearchAction));

            // Add nodes to the "Export" group of examples.
            examples[6].Groups.Add(new SpreadsheetExample("Export to HTML", ExportActions.ExportToHTMLAction));

            // Add nodes to the "Group Data" group of examples.
            examples[7].Groups.Add(new SpreadsheetExample("Group Rows", GroupAndOutlineActions.GroupRowsAction));
            examples[7].Groups.Add(new SpreadsheetExample("Group Columns", GroupAndOutlineActions.GroupColumnsAction));
            examples[7].Groups.Add(new SpreadsheetExample("Unroup Rows", GroupAndOutlineActions.UngroupRowsAction));
            examples[7].Groups.Add(new SpreadsheetExample("Unroup Columns", GroupAndOutlineActions.UngroupColumnsAction));
            examples[7].Groups.Add(new SpreadsheetExample("Create an Auto Outline", GroupAndOutlineActions.AutoOutlineAction));
            examples[7].Groups.Add(new SpreadsheetExample("Insert Subtotals", GroupAndOutlineActions.SubtotalAction));

            // Add nodes to the "Filter Data" group of examples.
            examples[8].Groups.Add(new SpreadsheetExample("Enable filtering", AutoFilterActions.ApplyFilterAction));
            examples[8].Groups.Add(new SpreadsheetExample("Sort by single column", AutoFilterActions.FilterAndSortBySingleColumnAction));
            examples[8].Groups.Add(new SpreadsheetExample("Sort by multiple columns", AutoFilterActions.FilterAndSortByMultipleColumnsAction));
            examples[8].Groups.Add(new SpreadsheetExample("Custom number filter", AutoFilterActions.FilterNumericByConditionAction));
            examples[8].Groups.Add(new SpreadsheetExample("Custom text filter", AutoFilterActions.FilterTextByConditionAction));
            examples[8].Groups.Add(new SpreadsheetExample("Custom date filter", AutoFilterActions.FilterDatesByConditionAction));
            examples[8].Groups.Add(new SpreadsheetExample("Filter by single value", AutoFilterActions.FilterByValuesAction));
            examples[8].Groups.Add(new SpreadsheetExample("Filter by multiple values", AutoFilterActions.FilterByMultipleValuesAction));
            examples[8].Groups.Add(new SpreadsheetExample("Filter mixed data types by values", AutoFilterActions.FilterMixedDataTypesByValuesAction));
            examples[8].Groups.Add(new SpreadsheetExample("Apply Top 10 filter", AutoFilterActions.Top10FilterAction));
            examples[8].Groups.Add(new SpreadsheetExample("Apply dynamic filter", AutoFilterActions.DynamicFilterAction));
            examples[8].Groups.Add(new SpreadsheetExample("Sort and filter by color", AutoFilterActions.FilterAndSortByColorAction));
            examples[8].Groups.Add(new SpreadsheetExample("Filter by font color", AutoFilterActions.FilterByFontColorAction));
            examples[8].Groups.Add(new SpreadsheetExample("Filter by fill color", AutoFilterActions.FilterByFillColorAction));
            examples[8].Groups.Add(new SpreadsheetExample("Filter by background color", AutoFilterActions.FilterByBackgroundColorAction));
            examples[8].Groups.Add(new SpreadsheetExample("Reapply filter", AutoFilterActions.ReapplyFilterAction));
            examples[8].Groups.Add(new SpreadsheetExample("Clear filter", AutoFilterActions.ClearFilterAction));
            examples[8].Groups.Add(new SpreadsheetExample("Disable filtering", AutoFilterActions.DisableFilterAction));

            // Add nodes to the "Document Properties" group of examples.
            examples[9].Groups.Add(new SpreadsheetExample("Built-in properties", DocumentPropertiesActions.BuiltInPropertiesAction));
            examples[9].Groups.Add(new SpreadsheetExample("Custom properties", DocumentPropertiesActions.CustomPropertiesAction));

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
            workbook.Options.Culture = System.Globalization.CultureInfo.CurrentCulture; 
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
