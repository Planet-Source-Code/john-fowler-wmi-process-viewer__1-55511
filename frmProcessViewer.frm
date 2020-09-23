VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcessViewer 
   Caption         =   "Process Viewer"
   ClientHeight    =   10410
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   ScaleHeight     =   10410
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstProperties 
      Height          =   10005
      Left            =   6000
      TabIndex        =   1
      Top             =   360
      Width           =   5910
   End
   Begin MSComctlLib.TreeView tvProcesses 
      Height          =   10035
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   17701
      _Version        =   393217
      Indentation     =   0
      LineStyle       =   1
      Style           =   7
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin VB.Label lblProperties 
      Caption         =   "Process Properties"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   60
      Width           =   5895
   End
   Begin VB.Label lblProcesses 
      Caption         =   "System Processes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   60
      Width           =   5895
   End
End
Attribute VB_Name = "frmProcessViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Locator As Object ' SWbemLocator
Dim Service As Object ' SWbemServices
Dim Processes As Object ' SWbemObjectSet
Dim Process As Object ' SWbemObject
Dim Property As Object ' SWbemProperty

Private Sub Form_Load()
    
    Me.Caption = App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision & "   " & App.LegalCopyright
    Set Locator = CreateObject("WbemScripting.SWbemLocator")
    Set Service = Locator.ConnectServer
    Set Processes = Service.InstancesOf("Win32_Process")
    
    LoadProcesses

End Sub

Private Sub LoadProcesses()

    tvProcesses.Nodes.Clear
    
    'add all system processes starting with root processes (parent process ID = 0)
    AddChildProcesses 0
    
End Sub

Private Sub AddChildProcesses(ProcessID As Long)

    'add child processes - recursively
    For Each Process In Processes
        With Process
            If .ParentProcessID = ProcessID Then
                On Error Resume Next
                Err.Clear
                If ProcessID = 0 Then
                    tvProcesses.Nodes.Add , , "PID" & .ProcessID, .Name
                Else
                    tvProcesses.Nodes.Add "PID" & .ParentProcessID, tvwChild, "PID" & .ProcessID, .Name
                    tvProcesses.Nodes("PID" & .ParentProcessID).Expanded = True
                End If
                Select Case Err
                    Case 0
                        AddChildProcesses .ProcessID
                    Case 35602
                        'key already added
                    Case Else
                        MsgBox "Error #" & Err & " - " & Err.Description & vbCrLf & vbCrLf & _
                            "Process ID: " & .ProcessID & vbCrLf & _
                            "Process Name: " & .Name & vbCrLf & _
                            "Parent Process ID: " & .ParentProcessID, vbCritical, "Error"
                End Select
            End If
        End With
    Next
End Sub

Private Sub Form_Resize()
    Me.tvProcesses.Height = Me.Height - 390
    Me.lstProperties.Height = Me.tvProcesses.Height
    Me.tvProcesses.Width = (Me.Width - 135) / 2
    Me.lblProcesses.Width = Me.tvProcesses.Width
    Me.lstProperties.Left = Me.tvProcesses.Left + Me.tvProcesses.Width + 25
    Me.lblProperties.Left = Me.lstProperties.Left
    Me.lstProperties.Width = Me.tvProcesses.Width
    Me.lblProperties.Width = Me.lstProperties.Width
End Sub

Private Sub tvProcesses_NodeClick(ByVal Node As MSComctlLib.Node)
    ShowProperties Val(Mid$(Node.Key, 4))
End Sub

Private Sub ShowProperties(ProcessID)

    lstProperties.Clear
    
    'find process by ID and show properties
    For Each Process In Processes
        With Process
            Debug.Print .ProcessID
            If .ProcessID = ProcessID Then
                For Each Property In .Properties_
                    Me.lstProperties.AddItem Property.Name & ": " & Property.Value
                Next Property
                Exit For
            End If
        End With
    Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set Process = Nothing
    Set Processes = Nothing
    Set Service = Nothing
    Set Locator = Nothing

End Sub

