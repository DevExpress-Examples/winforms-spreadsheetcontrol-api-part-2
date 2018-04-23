Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace SpreadsheetControl_API
	Friend NotInheritable Class Program
		''' <summary>
		''' The main entry point for the application.
		''' </summary>
		Private Sub New()
		End Sub
		<STAThread> _
		Shared Sub Main()
			System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("ru-RU")
			System.Threading.Thread.CurrentThread.CurrentUICulture = New System.Globalization.CultureInfo("ru-RU")
			Application.EnableVisualStyles()
			Application.SetCompatibleTextRenderingDefault(False)
			Application.Run(New Form1())


		End Sub
	End Class
End Namespace
