VERSION 5.00
Begin VB.Form frmWeb 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "ExcelChartAsWebPage"
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
End
Attribute VB_Name = "frmWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function BuildChart(chartTitle As Variant, categoryAxisTitle As Variant, valueAxisTitle As Variant, seriesTitleA As Variant, categoryAxis As Variant, seriesDataA As Variant, seriesDataB As Variant, folderPathForWebPage As String, fileNameForWebPage As String) As Boolean

 On Error GoTo ErrorHandler
 
    Dim webPageCreationResults As Boolean
    webPageCreationResults = False
    
    Dim oExcel As Excel.Application
    Dim oBook As Excel.Workbook
    Dim oBooks As Excel.Workbooks

    'Start Excel and open the workbook
    Set oExcel = CreateObject("Excel.Application")
    Set oBooks = oExcel.Workbooks
    Set oBook = oBooks.Open(App.Path & "\ExcelChart.xls")
       
    oExcel.Visible = True
    
    'Calls the Excel Macro to make Chart's & WebPage Creation
     webPageCreationResults = oExcel.Run("DrawBarChart", chartTitle, categoryAxisTitle, valueAxisTitle, seriesTitleA, categoryAxis, seriesDataA, folderPathForWebPage, fileNameForWebPage)
     
    'Close workbook and exit the Excel
    oBook.Close (False)
    Set oBook = Nothing
    Set oBooks = Nothing
    oExcel.Quit
    Set oExcel = Nothing
    
    'Returns the results WebPage Creation
    BuildChart = webPageCreationResults
 
 Exit Function
   
ErrorHandler:

    'Close workbook and exit the Excel
    oBook.Close (False)
    Set oBook = Nothing
    Set oBooks = Nothing
    oExcel.Quit
    Set oExcel = Nothing
    
    'Returns the results WebPage Creation
    BuildChart = webPageCreationResults

End Function

Private Sub Command1_Click()

Dim categoryAxis(5), seriesDataA(5), seriesDataB(5)

categoryAxis(0) = "10:00 AM"
categoryAxis(1) = "12:00 PM"
categoryAxis(2) = "2:00 PM"
categoryAxis(3) = "4:00 PM"
categoryAxis(4) = "6:00 PM"

seriesDataA(0) = "10"
seriesDataA(1) = "60"
seriesDataA(2) = "30"
seriesDataA(3) = "100"
seriesDataA(4) = "80"


seriesDataB(0) = "100"
seriesDataB(1) = "160"
seriesDataB(2) = "130"
seriesDataB(3) = "1000"
seriesDataB(4) = "180"

webFlag = BuildChart("Forecast Data", "Time Slots", "No.Of Calls", "No.Of Calls", categoryAxis, seriesDataA, seriesDataB, App.Path, "myChart.htm")

If (webFlag) Then
    MsgBox "[PASS] : Creating Excel Chart As WebPage!", vbInformation
Else
    MsgBox "[FAIL] : Creating Excel Chart As WebPage!", vbCritical
End If

End Sub
