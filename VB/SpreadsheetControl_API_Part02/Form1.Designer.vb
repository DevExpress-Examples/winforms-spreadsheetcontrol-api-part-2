Namespace SpreadsheetControl_API
    Partial Public Class Form1
        ''' <summary>
        ''' Required designer variable.
        ''' </summary>
        Private components As System.ComponentModel.IContainer = Nothing

        ''' <summary>
        ''' Clean up any resources being used.
        ''' </summary>
        ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso (components IsNot Nothing) Then
                components.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        #Region "Windows Form Designer generated code"

        ''' <summary>
        ''' Required method for Designer support - do not modify
        ''' the contents of this method with the code editor.
        ''' </summary>
        Private Sub InitializeComponent()
            Me.spreadsheetControl1 = New DevExpress.XtraSpreadsheet.SpreadsheetControl()
            Me.treeList1 = New DevExpress.XtraTreeList.TreeList()
            Me.treeListColumn1 = New DevExpress.XtraTreeList.Columns.TreeListColumn()
            Me.LayoutControl1 = New DevExpress.XtraLayout.LayoutControl()
            Me.Root = New DevExpress.XtraLayout.LayoutControlGroup()
            Me.LayoutControlItem1 = New DevExpress.XtraLayout.LayoutControlItem()
            Me.LayoutControlItem2 = New DevExpress.XtraLayout.LayoutControlItem()
            Me.button1 = New DevExpress.XtraEditors.SimpleButton()
            Me.LayoutControlItem3 = New DevExpress.XtraLayout.LayoutControlItem()
            Me.SplitterItem1 = New DevExpress.XtraLayout.SplitterItem()
            CType(Me.treeList1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.LayoutControl1.SuspendLayout()
            CType(Me.Root, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.SplitterItem1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'spreadsheetControl1
            '
            Me.spreadsheetControl1.Location = New System.Drawing.Point(307, 12)
            Me.spreadsheetControl1.Name = "spreadsheetControl1"
            Me.spreadsheetControl1.Options.Culture = New System.Globalization.CultureInfo("en-US")
            Me.spreadsheetControl1.Options.Export.Csv.Culture = New System.Globalization.CultureInfo("")
            Me.spreadsheetControl1.Options.Export.Txt.Culture = New System.Globalization.CultureInfo("")
            Me.spreadsheetControl1.Options.Export.Txt.ValueSeparator = Global.Microsoft.VisualBasic.ChrW(44)
            Me.spreadsheetControl1.Options.Import.Csv.Culture = New System.Globalization.CultureInfo("")
            Me.spreadsheetControl1.Options.Import.Txt.Culture = New System.Globalization.CultureInfo("")
            Me.spreadsheetControl1.Size = New System.Drawing.Size(1046, 862)
            Me.spreadsheetControl1.TabIndex = 0
            Me.spreadsheetControl1.Text = "spreadsheetControl1"
            '
            'treeList1
            '
            Me.treeList1.Appearance.FocusedCell.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
            Me.treeList1.Appearance.FocusedCell.ForeColor = System.Drawing.Color.Blue
            Me.treeList1.Appearance.FocusedCell.Options.UseFont = True
            Me.treeList1.Appearance.FocusedCell.Options.UseForeColor = True
            Me.treeList1.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.treeListColumn1})
            Me.treeList1.Location = New System.Drawing.Point(12, 12)
            Me.treeList1.Name = "treeList1"
            Me.treeList1.OptionsBehavior.Editable = False
            Me.treeList1.OptionsView.ShowColumns = False
            Me.treeList1.OptionsView.ShowIndicator = False
            Me.treeList1.Size = New System.Drawing.Size(281, 797)
            Me.treeList1.TabIndex = 1
            '
            'treeListColumn1
            '
            Me.treeListColumn1.Caption = "Name"
            Me.treeListColumn1.FieldName = "Name"
            Me.treeListColumn1.Name = "treeListColumn1"
            Me.treeListColumn1.Visible = True
            Me.treeListColumn1.VisibleIndex = 0
            Me.treeListColumn1.Width = 92
            '
            'LayoutControl1
            '
            Me.LayoutControl1.Controls.Add(Me.button1)
            Me.LayoutControl1.Controls.Add(Me.spreadsheetControl1)
            Me.LayoutControl1.Controls.Add(Me.treeList1)
            Me.LayoutControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.LayoutControl1.Location = New System.Drawing.Point(0, 0)
            Me.LayoutControl1.Name = "LayoutControl1"
            Me.LayoutControl1.OptionsCustomizationForm.DesignTimeCustomizationFormPositionAndSize = New System.Drawing.Rectangle(627, 98, 650, 400)
            Me.LayoutControl1.Root = Me.Root
            Me.LayoutControl1.Size = New System.Drawing.Size(1365, 886)
            Me.LayoutControl1.TabIndex = 4
            Me.LayoutControl1.Text = "LayoutControl1"
            '
            'Root
            '
            Me.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.[True]
            Me.Root.GroupBordersVisible = False
            Me.Root.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.LayoutControlItem1, Me.LayoutControlItem2, Me.LayoutControlItem3, Me.SplitterItem1})
            Me.Root.Name = "Root"
            Me.Root.Size = New System.Drawing.Size(1365, 886)
            Me.Root.TextVisible = False
            '
            'LayoutControlItem1
            '
            Me.LayoutControlItem1.Control = Me.treeList1
            Me.LayoutControlItem1.Location = New System.Drawing.Point(0, 0)
            Me.LayoutControlItem1.Name = "LayoutControlItem1"
            Me.LayoutControlItem1.Size = New System.Drawing.Size(285, 801)
            Me.LayoutControlItem1.TextSize = New System.Drawing.Size(0, 0)
            Me.LayoutControlItem1.TextVisible = False
            '
            'LayoutControlItem2
            '
            Me.LayoutControlItem2.Control = Me.spreadsheetControl1
            Me.LayoutControlItem2.Location = New System.Drawing.Point(295, 0)
            Me.LayoutControlItem2.Name = "LayoutControlItem2"
            Me.LayoutControlItem2.Size = New System.Drawing.Size(1050, 866)
            Me.LayoutControlItem2.TextSize = New System.Drawing.Size(0, 0)
            Me.LayoutControlItem2.TextVisible = False
            '
            'button1
            '
            Me.button1.Location = New System.Drawing.Point(12, 813)
            Me.button1.Name = "button1"
            Me.button1.Size = New System.Drawing.Size(281, 61)
            Me.button1.StyleController = Me.LayoutControl1
            Me.button1.TabIndex = 4
            Me.button1.Text = "Run"
            '
            'LayoutControlItem3
            '
            Me.LayoutControlItem3.Control = Me.button1
            Me.LayoutControlItem3.Location = New System.Drawing.Point(0, 801)
            Me.LayoutControlItem3.MaxSize = New System.Drawing.Size(0, 65)
            Me.LayoutControlItem3.MinSize = New System.Drawing.Size(30, 65)
            Me.LayoutControlItem3.Name = "LayoutControlItem3"
            Me.LayoutControlItem3.Size = New System.Drawing.Size(285, 65)
            Me.LayoutControlItem3.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom
            Me.LayoutControlItem3.TextSize = New System.Drawing.Size(0, 0)
            Me.LayoutControlItem3.TextVisible = False
            '
            'SplitterItem1
            '
            Me.SplitterItem1.AllowHotTrack = True
            Me.SplitterItem1.Location = New System.Drawing.Point(285, 0)
            Me.SplitterItem1.Name = "SplitterItem1"
            Me.SplitterItem1.Size = New System.Drawing.Size(10, 866)
            '
            'Form1
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(1365, 886)
            Me.Controls.Add(Me.LayoutControl1)
            Me.Name = "Form1"
            Me.Text = "Form1"
            CType(Me.treeList1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.LayoutControl1.ResumeLayout(False)
            CType(Me.Root, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.SplitterItem1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)

        End Sub

#End Region

        Private spreadsheetControl1 As DevExpress.XtraSpreadsheet.SpreadsheetControl
        Private treeList1 As DevExpress.XtraTreeList.TreeList
        Private treeListColumn1 As DevExpress.XtraTreeList.Columns.TreeListColumn
        Friend WithEvents LayoutControl1 As DevExpress.XtraLayout.LayoutControl
        Friend WithEvents button1 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents Root As DevExpress.XtraLayout.LayoutControlGroup
        Friend WithEvents LayoutControlItem1 As DevExpress.XtraLayout.LayoutControlItem
        Friend WithEvents LayoutControlItem2 As DevExpress.XtraLayout.LayoutControlItem
        Friend WithEvents LayoutControlItem3 As DevExpress.XtraLayout.LayoutControlItem
        Friend WithEvents SplitterItem1 As DevExpress.XtraLayout.SplitterItem
    End Class
End Namespace

