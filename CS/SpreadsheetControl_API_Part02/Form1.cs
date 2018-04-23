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
            #endregion

            #region ExampleNodes
 
            // Add nodes to the "Shapes" group of examples.
            examples[0].Groups.Add(new SpreadsheetExample("Insert a picture", ShapeActions.InsertShapeAction));
            examples[0].Groups.Add(new SpreadsheetExample("Insert a picture from URI", ShapeActions.InsertShapeFromUriAction));
            examples[0].Groups.Add(new SpreadsheetExample("Modify a picture", ShapeActions.ModifyShapeAction));
            
            // Add nodes to the "Cuistom Function" group of examples.
            examples[1].Groups.Add(new SpreadsheetExample("Add a SPHEREMASS function", CustomFunctionActions.SphereMassAction ));

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
