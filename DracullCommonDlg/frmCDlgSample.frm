VERSION 5.00
Begin VB.Form frmCDlgSample 
   Caption         =   "DracullSoft CommonDlg"
   ClientHeight    =   5400
   ClientLeft      =   5070
   ClientTop       =   2550
   ClientWidth     =   8385
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCDlgSample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSAVE 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2580
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   1260
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00BC741B&
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   900
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2400
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu mnuFileTOP 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Open..."
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save..."
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Page Setup"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Print"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Print Thumbnail"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&About"
         Index           =   10
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   20
      End
   End
   Begin VB.Menu mnuOTop 
      Caption         =   "&Options"
      Begin VB.Menu mnuOption 
         Caption         =   "Select Font"
         Index           =   21
      End
   End
   Begin VB.Menu mnuZoomTop 
      Caption         =   "&Zoom"
      Index           =   1
      Begin VB.Menu mnuZoom 
         Caption         =   "10%"
         Index           =   10
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "20%"
         Index           =   20
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "30%"
         Index           =   30
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "40%"
         Index           =   40
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "50%"
         Index           =   50
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "60%"
         Index           =   60
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "70%"
         Index           =   70
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "80%"
         Index           =   80
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "90%"
         Index           =   90
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "100%"
         Index           =   100
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "125%"
         Index           =   125
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "150%"
         Index           =   150
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "175%"
         Index           =   175
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "200%"
         Index           =   200
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "300%"
         Index           =   300
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "400%"
         Index           =   400
      End
   End
   Begin VB.Menu mnuSettingTop 
      Caption         =   "&Settings"
      Begin VB.Menu mnuSettings 
         Caption         =   "Examples of FullVersion"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmCDlgSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ======================================================================================
' File  :     cCommonDialog
' Author:     DracullSoft (DiceSix)
' --------------------------------------------------------------------------------------
' Copyright © 2008 DracullSoft aka DiceSix
' Author or copyright holders can not be held responsible for anything.
' Example based on : Clint LaFever  vb Calendar Maker
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=42227&lngWId=1
' --------------------------------------------------------------------------------------
' Purpose: Examples of using cCommonDlg
' Date  :  05-05-2008
'
' Date  :  12-05-2008
'          Now just showing the picture and printing it.. to make the use Simpler to
'          understand.. also no comments and no good feedback came about the Calendar
'          as an application. Later we plan to release it as a Free Application via
'           http://www.digiapp.com
'          If you enjoy game development with VB6 then checkout
'           http://gamedev.digiapp.com
' ======================================================================================
Option Explicit

Private Declare Sub InitCommonControls Lib "Comctl32" ()

'------------------------------------------------------------
' Stores the last image we picked for the calendar
'------------------------------------------------------------
Private m_LastImage As String

Private m_HeadTextFont      As New StdFont
Private m_ZoomFactor        As Single

Private m_Paper             As EPaperSize


'Private m_p As Picture
Private Sub Form_Initialize()
  '__ Must be done in Main() in a bas module or in the Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
Dim cc As cCommonDlg
    Set cc = New cCommonDlg
    Me.BackColor = vbWhite
    Me.AutoRedraw = True
    Me.picSAVE.AutoRedraw = True
    m_Paper = epsLegal
    '------------------------------------------------------------
    ' Prepare the form making it the 80 percent the
    ' size of a standard sheet of paper [US standard paper that is]
    '------------------------------------------------------------
    
    
    Label2.Top = Label1.Top + 10
    Label2.Left = Label1.Left + 10
    
    '------------------------------------------------------------
    ' Need to set AutoRedraw to true since we are drawing
    ' to these things.
    '------------------------------------------------------------
    
    CopyStdFont m_HeadTextFont, Me.Font
    m_HeadTextFont.Size = 16
    Set Label1.Font = m_HeadTextFont
    Set Label2.Font = m_HeadTextFont
    Label1.Caption = "Common Dialogs with Thumbnails"
    Label2.Caption = Label1.Caption
    
    
    m_LastImage = App.Path & "\indiana_jones_art_harrison_ford.jpg" '"\jessica_alba_002.jpg"
    Set Image1.Picture = LoadPicture(m_LastImage)
    

    
    mnuZoom_Click 50 ' m_ZoomFactor = 0.5
    
    
End Sub

Private Sub CopyStdFont(toFont As StdFont, fromFont As StdFont)
  toFont = fromFont
  toFont.Bold = fromFont.Bold
  toFont.Charset = fromFont.Charset
  toFont.Italic = fromFont.Italic
  toFont.Size = fromFont.Size
  toFont.Strikethrough = fromFont.Strikethrough
  toFont.Underline = fromFont.Underline
  toFont.Weight = fromFont.Weight
End Sub

Private Sub mnuOption_Click(Index As Integer)
  Dim cc   As cCommonDlg
  Dim aColor As OLE_COLOR
  Dim myfont As New StdFont
  
  Select Case Index
    
   Case 21 ' Header text font
    Set cc = New cCommonDlg
    aColor = 1
    CopyStdFont myfont, m_HeadTextFont
    If cc.VBChooseFont(myfont, , Me.hwnd, aColor) Then
      CopyStdFont m_HeadTextFont, myfont
      
      Set Label2.Font = m_HeadTextFont
      Set Label1.Font = m_HeadTextFont
      Label2.ForeColor = aColor
    End If
   
 End Select
 
End Sub



Private Sub Form_Resize()
   On Error Resume Next
   
   DrawGrid Me
   DoEvents
   
End Sub


Private Sub mnuSettings_Click(Index As Integer)
  Select Case Index
  Case 2:
    
     Call DoShellExecute("open", App.Path & "\dracullCalendar.pdf", vbNullString, vbNullString, vbNormalFocus)

  End Select
End Sub

Private Sub mnuZoom_Click(Index As Integer)
  On Error Resume Next
  Dim x As Menu
  For Each x In mnuZoom
    x.Checked = False
  Next
  
  mnuZoom(Index).Checked = True

  m_ZoomFactor = Index
  m_ZoomFactor = m_ZoomFactor / 100
  SetPaperAndZoom
End Sub

Private Sub SetPaperAndZoom()
    On Error GoTo hell:
    Dim cc As cCommonDlg
  Set cc = New cCommonDlg
  
  Me.Width = ScaleX(cc.GetPaperSizeX(m_Paper), cc.GetPaperMeasure(m_Paper), vbTwips) * m_ZoomFactor
  Me.Height = ScaleY(cc.GetPaperSizeY(m_Paper), cc.GetPaperMeasure(m_Paper), vbTwips) * m_ZoomFactor
hell:
End Sub


Private Sub mnuFile_Click(Index As Integer)
Dim cc As cCommonDlg
Dim sFile As String
   Select Case Index
   Case 0
      
      DoOPENImage
   Case 1
      DoSAVE
'      Set cc = New cCommonDlg
'      If cc.VBGetSaveFileName(sFile, , , "Text Files (*.txt)|*.txt|All Files (*.*)|*.*", , , , "TXT", Me.hwnd) Then
'         WriteFileText sFile, txtDoc.Text
'      End If
   
   Case 2
'      Set cc = New cCommonDlg
'      If cc.VBPageSetupDlg Then
'
'      End If
      Dim aPaper As EPaperSize
      Set cc = New cCommonDlg
      If cc.VBPageSetupDlg(Me.hwnd, True, False, False, False, , , , , , , , , aPaper) Then
        m_Paper = aPaper
        Debug.Print "gotit"
        SetPaperAndZoom
      End If
  Case 3
      DoPRINT


  Case 4
      DoPRINT 25


   Case 10
      MsgBox "Common Dialogs Example by DracullSoft" & vbCr & vbCr & _
      "Based on work by Steve McMahon and Bruce McKinney, Randy Birch of VBnet and Clint LaFever.", , "Dracull Common Dlg"
   Case 20
      Unload Me
    
   End Select
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DoPRINT
' Purpose   : Printing the current calendar view
'   lThumbnailPct:  allow you to set a value between 0 and 100 to scale the print
'
'---------------------------------------------------------------------------------------
Private Sub DoPRINT(Optional ByVal lThumbnailPct As Single = 100)
    On Error Resume Next
    Dim cc As cCommonDlg, p As Object
    '------------------------------------------------------------
    ' Set our hidden picture box to the size of paper
    ' with quarter inch margins, the draw the calendar
    ' to it then paint that image to the printer.
    '------------------------------------------------------------
    Set cc = New cCommonDlg
    Set p = Printer
    cc.VBPrintDlg Me.hdc, , , , , , , True, False, , , , Me.hwnd, p
    If cc.SelectedPrinterName <> "" Then
        For Each p In Printers
            If p.DeviceName = cc.SelectedPrinterName Then
                Set Printer = p
                Exit For
            End If
        Next
    End If
    If cc.SelectedPrinterName <> "" Then
        Screen.MousePointer = vbHourglass
        With Me.picSAVE
            .Cls
          Debug.Print "papersize: " & Printer.PaperSize
          
          .Width = ScaleX(cc.GetPaperSizeX(Printer.PaperSize), cc.GetPaperMeasure(Printer.PaperSize), vbTwips)
          .Height = ScaleY(cc.GetPaperSizeY(Printer.PaperSize), cc.GetPaperMeasure(Printer.PaperSize), vbTwips)
            
       '     .Width = ScaleX(8, vbInches, vbTwips)
       '     .Height = ScaleX(10.5, vbInches, vbTwips)
            DrawGrid picSAVE, True
        End With
        Printer.PaintPicture picSAVE.Image, 0, 0, Printer.ScaleWidth * lThumbnailPct / 100, Printer.ScaleHeight * lThumbnailPct / 100
        Printer.EndDoc
        Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub DoSAVE()
    On Error Resume Next
    Dim fName As String, cc As cCommonDlg
    '------------------------------------------------------------
    ' Using the CDLG class, show  Common Dialog
    ' box asking for a place to save to, if a file
    ' name is given then draw our calendar to the hidden
    ' picture box which is the size of a stanard piece
    ' of paper with quarter inch margins.  Then save
    ' that image to the file given.
    '------------------------------------------------------------
    Set cc = New cCommonDlg
    cc.VBGetSaveFileName fName, , , "BMP Files|*.bmp", , CurDir, "Save to", "*.bmp", Me.hwnd
    If fName <> "" Then
        Screen.MousePointer = vbHourglass
        With Me.picSAVE
            .Cls
            .Width = ScaleX(8, vbInches, vbTwips)
            .Height = ScaleY(10.5, vbInches, vbTwips)
            DrawGrid picSAVE
            SavePicture .Image, fName
        End With
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub DoOPENImage()
On Error Resume Next
    Dim fName As String
    Dim cc   As cCommonDlg
    '------------------------------------------------------------
    ' using the CDLG class, show the Common Dialog
    ' box asking for a file to load.  If one is picked,
    ' set m_LastImage and call the draw function again.
    '------------------------------------------------------------
    Set cc = New cCommonDlg
    cc.VBGetOpenFileName fName, , , , , , "Image Files|*.bmp;*.jpg;*.gif|All Files|*.*", , CurDir, "Load Image", , Me.hwnd, OFN_HIDEREADONLY, False, SHVIEW_THUMBNAIL
    If fName <> "" Then
        m_LastImage = fName
        Me.Cls
        Set Image1.Picture = LoadPicture(m_LastImage)
        
        DrawGrid Me
    End If
End Sub


Private Sub DrawGrid(DrawTo As Object, Optional ByVal isPrinting As Boolean = False)
    On Error Resume Next
 
    If Not Image1.Picture Is Nothing Then
      With Image1.Picture
       If .Width > 0 And .Height > 0 Then
             DrawTo.PaintPicture Image1.Picture, 0, 0, DrawTo.ScaleWidth, DrawTo.ScaleHeight
       End If
      End With
    End If
End Sub


