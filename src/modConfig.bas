Attribute VB_Name = "modConfig"
Option Explicit

'===========================================
' QA Automation Framework Configuration
'===========================================

Public Const MATRIX_SHEET As String = "Matrix"
Public Const REPORT_SHEET As String = "Processing_Report"

Public Const VERSION As String = "QA Automation Framework v1.0"

'Folders
Public Const INPUT_FOLDER As String = "\Input\"
Public Const OUTPUT_FOLDER As String = "\Output\"
Public Const SPLIT_FOLDER As String = "\Split\"
Public Const REPORT_FOLDER As String = "\Reports\"
Public Const BACKUP_FOLDER As String = "\Backup\"

'Paths
Public FrameworkPath As String
Public InputPath As String
Public OutputPath As String
Public SplitPath As String
Public ReportPath As String
Public BackupPath As String

'Execution Time
Public ProcessingStart As Date
Public ProcessingEnd As Date

'Workbook Count
Public TotalTC As Long
Public PassedTC As Long
Public FailedTC As Long
Public FailedTCList As String
Public FailedTCListText As String

'Current Workbook Information
Public InputWorkbookName As String
Public ReportFileName As String
Public TimeStamp As String
