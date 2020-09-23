VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl msEditor2 
   Alignable       =   -1  'True
   ClientHeight    =   4920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10215
   ScaleHeight     =   4920
   ScaleWidth      =   10215
   ToolboxBitmap   =   "msEditor2.ctx":0000
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5550
      Top             =   1470
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4665
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Text            =   "Ln "
            TextSave        =   "Ln "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Text            =   "Col"
            TextSave        =   "Col"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   706
            MinWidth        =   706
            Text            =   "Tot"
            TextSave        =   "Tot"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   10689
            MinWidth        =   1058
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox Richinsert 
      Height          =   435
      Left            =   5520
      TabIndex        =   2
      Top             =   3300
      Visible         =   0   'False
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   767
      _Version        =   393217
      TextRTF         =   $"msEditor2.ctx":0312
   End
   Begin MSComctlLib.ImageList ilsToolbar 
      Left            =   5760
      Top             =   2610
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   27
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":0396
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":04A8
            Key             =   "Font"
            Object.Tag             =   "Font"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":05BA
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":06CC
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":07DE
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":08F0
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":0A02
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":0B14
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":0C26
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":0D38
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":0E4A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":0F5C
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":106E
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":1180
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":1292
            Key             =   "StrikeThru"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":13A4
            Key             =   "Color"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":16C6
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":17D8
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":18EA
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":19FC
            Key             =   "DecreaseIndent"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":1E06
            Key             =   "IncreaseIndent"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":2210
            Key             =   "Printer"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":2752
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":2C94
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":2D2D
            Key             =   "InsertP"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":3105
            Key             =   "Object"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "msEditor2.ctx":34A4
            Key             =   "Replace"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5610
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilsToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   30
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open a file"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "InsertPicture"
            Object.ToolTipText     =   "Insert Picture"
            ImageKey        =   "InsertP"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "InsertText"
            Object.ToolTipText     =   "Insert Text file"
            ImageKey        =   "Insert"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "InsertObject"
            Object.ToolTipText     =   "Insert an object"
            ImageKey        =   "Object"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print WYSIWYG"
            ImageKey        =   "Printer"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Replace"
            Object.ToolTipText     =   "Find And Replace"
            ImageKey        =   "Replace"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageKey        =   "Undo"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Object.ToolTipText     =   "Redo"
            ImageKey        =   "Redo"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "StrikeThru"
            Object.ToolTipText     =   "StrikeThru"
            ImageKey        =   "StrikeThru"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Color"
            Object.ToolTipText     =   "Font Color"
            ImageKey        =   "Color"
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font"
            Object.ToolTipText     =   "Font type and Size"
            ImageKey        =   "Font"
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            Object.ToolTipText     =   "Align Left"
            ImageKey        =   "Left"
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageKey        =   "Center"
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            Object.ToolTipText     =   "Align Right"
            ImageKey        =   "Right"
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DecreaseIndent"
            Object.ToolTipText     =   "Decrease Indent"
            ImageKey        =   "DecreaseIndent"
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "IncreaseIndent"
            Object.ToolTipText     =   "Increase Indent"
            ImageKey        =   "IncreaseIndent"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox msEdit 
      Height          =   2205
      Left            =   510
      TabIndex        =   1
      Top             =   1260
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3889
      _Version        =   393217
      BackColor       =   12648447
      ScrollBars      =   3
      TextRTF         =   $"msEditor2.ctx":3578
   End
End
Attribute VB_Name = "msEditor2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type Rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type CharRange
  cpMin As Long     ' First character of range (0 for start of doc)
  cpMax As Long     ' Last character of range (-1 for end of doc)
End Type

Private Type FormatRange
  hdc As Long       ' Actual DC to draw on
  hdcTarget As Long ' Target DC for determining text formatting
  rc As Rect        ' Region of the DC to draw to (in twips)
  rcPage As Rect    ' Region of the entire DC (page size) (in twips)
  chrg As CharRange ' Range of text to draw (see above declaration)
End Type


Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, lp As Any) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As Long, ByVal lpInitData As Long) As Long

Private Const AnInch As Long = 1440   '1440 twips per inch
Private Const QuarterInch As Long = 360
Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113

Private Const EM_LINEINDEX = &HBB
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_GETLINECOUNT = &HBA

Private Const WM_PASTE = &H302

Private mnuBold As Boolean
Private mnuItalic As Boolean
Private mnuUnderline As Boolean
Private mnuStrikeThru As Boolean

Private DocumentName As String
Private DefaultFont As String
Private DefaultFontSize As Integer
Private DefaultTextColor As String
Private DefaultBackgroundColor As String
Private DefaultBold As Boolean
Private DefaultItalic As Boolean
Private DefaultUnderline As Boolean
Private DefaultStrikeThru As Boolean
Private DefaultAlignment As String
Private IndentSize As Integer

Private mDocumentName As String
Private mDefaultFont As String
Private mDefaultFontSize As String
Private mDefaultTextColor As String
Private mDefaultBackgroundColor As String
Private mDefaultBold As String
Private mDefaultItalic As String
Private mDefaultUnderline As String
Private mDefaultStrikeThru As String
Private mDefaultAlignment As String
Private mItentSize As String

Private Cancelled As Boolean
Private Saved As Boolean

Private msCancel As Boolean
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(1000) As String

'Const m_def_Enabled = True
Const m_def_ForeColor = 0
Const m_def_FontUnderline = 0
Const m_def_FontStrikethru = 0
Const m_def_FontSize = 0
Const m_def_FontName = ""
Const m_def_FontItalic = 0
Const m_def_FontBold = 0
Const m_def_TextRTF = "0"

Const m_def_Open_Visible = True
Const m_def_Save_Visible = True
Const m_def_New_Visible = True
Const m_def_Print_Visible = True

'Property Variables:

'Dim m_Enabled As Boolean
Dim m_ForeColor As Long
Dim m_FontUnderline As Boolean
Dim m_FontStrikethru As Boolean
Dim m_FontSize As Single
Dim m_FontName As String
Dim m_FontItalic As Boolean
Dim m_FontBold As Boolean
Dim m_TextRTF As String

Private WithEvents mReplaceCallback As frmReplace
Attribute mReplaceCallback.VB_VarHelpID = -1

Private Sub Timer1_Timer()
  ShowStatus
End Sub

Private Sub UserControl_Initialize()
              
    mDocumentName = "NewPage"
    mDefaultFont = "Tahoma"
    mDefaultFontSize = 10
    mDefaultTextColor = 0
    mDefaultBackgroundColor = &HC0FFFF
    mDefaultBold = False
    mDefaultItalic = False
    mDefaultUnderline = False
    mDefaultStrikeThru = False
    mDefaultAlignment = "Left"
    mItentSize = 500
    
    DocumentName = mDocumentName
    DefaultFont = mDefaultFont
    DefaultFontSize = mDefaultFontSize
    DefaultTextColor = mDefaultTextColor
    DefaultBackgroundColor = mDefaultBackgroundColor
    DefaultBold = mDefaultBold
    DefaultItalic = mDefaultItalic
    DefaultUnderline = mDefaultUnderline
    DefaultStrikeThru = mDefaultStrikeThru
    DefaultAlignment = mDefaultAlignment
    IndentSize = mItentSize
        
    Saved = False
    Cancelled = False
    
    With tlbar
    
        .Buttons("Cut").Enabled = False
        .Buttons("Copy").Enabled = False
        .Buttons("Print").Enabled = False
        .Buttons("Undo").Enabled = False
        .Buttons("Redo").Enabled = False
        .Buttons("Paste").Enabled = False
        .Buttons("Save").Enabled = False
                
    End With
 
    With msEdit
       
        .SelFontSize = DefaultFontSize
        .SelFontName = DefaultFont
        .SelColor = DefaultTextColor
        .BackColor = DefaultBackgroundColor
        .SelBold = DefaultBold
        .SelItalic = DefaultItalic
        .SelUnderline = DefaultUnderline
        .SelStrikeThru = DefaultStrikeThru
        
    End With

End Sub

Private Sub msBold()
    If mnuBold Then
        msEdit.SelBold = False
        mnuBold = False
        tlbar.Buttons("Bold").Value = tbrUnpressed
    Else
        msEdit.SelBold = True
        mnuBold = True
        tlbar.Buttons("Bold").Value = tbrPressed
    End If
End Sub

Private Sub msCenter()
    
    msEdit.SelAlignment = 2
    
    tlbar.Buttons("Left").Value = tbrUnpressed
    tlbar.Buttons("Center").Value = tbrPressed
    tlbar.Buttons("Right").Value = tbrUnpressed

End Sub


Private Sub msCopy()
    Clipboard.Clear
    Clipboard.SetText msEdit.SelText, 1
    
    tlbar.Buttons("Paste").Enabled = True
    
End Sub

Private Sub msCut()
    
    Clipboard.Clear
    Clipboard.SetText msEdit.SelText, 1
    msEdit.SelText = ""
    
    tlbar.Buttons("Paste").Enabled = True

End Sub

Private Sub msDecreaseIndent()
    msEdit.SelIndent = msEdit.SelIndent - IndentSize
End Sub

Public Sub SaveNow()
    
    If Saved = False Then
        Dim boolFlag As Boolean
    
        On Error GoTo handlelit
        boolFlag = False
        With CommonDialog1
            .Filter = "Editor Documents (*.rtf)|*.rtf"
            .ShowSave
        End With
        
        If Not boolFlag Then
            msEdit.SaveFile CommonDialog1.FileName, 0
            Saved = True
        End If
     End If
     
     Exit Sub
    
handlelit:
    If Err.Number = cdlCancel Then
        Saved = False
        Resume Next
    Else
            MsgBox Err.Description
    End If
    
    
End Sub


Private Sub msFontColor()
    On Error Resume Next
    CommonDialog1.ShowColor
    msEdit.SelColor = CommonDialog1.Color
End Sub
Private Sub msFont()
    On Error Resume Next
    
    CommonDialog1.Flags = cdlCFBoth Or cdlCFTTOnly Or cdlCFEffects
                                                      
    With msEdit
        CommonDialog1.FontName = .SelFontName
        CommonDialog1.FontSize = .SelFontSize
        CommonDialog1.FontBold = .SelBold
        CommonDialog1.FontStrikethru = .SelStrikeThru
        CommonDialog1.FontUnderline = .SelUnderline
        CommonDialog1.FontItalic = .SelItalic
        CommonDialog1.Color = .SelColor
    End With
    
    CommonDialog1.ShowFont
    
    With msEdit
        .SelFontName = CommonDialog1.FontName
        .SelFontSize = CommonDialog1.FontSize
        .SelBold = CommonDialog1.FontBold
        .SelItalic = CommonDialog1.FontItalic
        .SelStrikeThru = CommonDialog1.FontStrikethru
        .SelUnderline = CommonDialog1.FontUnderline
        .SelColor = CommonDialog1.Color
    End With
     
End Sub

Private Sub msIncreaseIndent()
    msEdit.SelIndent = msEdit.SelIndent + IndentSize
End Sub

Private Sub msItalic()
    If mnuItalic Then
        msEdit.SelItalic = False
        mnuItalic = False
        tlbar.Buttons("Italic").Value = tbrUnpressed
    Else
        msEdit.SelItalic = True
        mnuItalic = True
        tlbar.Buttons("Italic").Value = tbrPressed
    End If
End Sub

Private Sub msLeft()
    msEdit.SelAlignment = 0
    
    tlbar.Buttons("Left").Value = tbrPressed
    tlbar.Buttons("Center").Value = tbrUnpressed
    tlbar.Buttons("Right").Value = tbrUnpressed

End Sub

Private Sub msNew()
    
    With tlbar
        .Buttons("Cut").Enabled = True
        .Buttons("Copy").Enabled = True
        .Buttons("Print").Enabled = True
        .Buttons("Undo").Enabled = True
        .Buttons("Redo").Enabled = True
        .Buttons("Paste").Enabled = True
        .Buttons("Save").Enabled = True
    End With
    
    msEdit.Text = ""
    
End Sub

Private Sub msPaste()
    
    msEdit.SelText = Clipboard.GetText(1)
    tlbar.Buttons("Save").Enabled = True
    
End Sub

Private Sub msRedo()
    
    Redo
    tlbar.Buttons("Undo").Enabled = True
    
End Sub

Private Sub msRight()
    msEdit.SelAlignment = 1
    
    tlbar.Buttons("Left").Value = tbrUnpressed
    tlbar.Buttons("Center").Value = tbrUnpressed
    tlbar.Buttons("Right").Value = tbrPressed
    
End Sub

Private Sub msStrikeThrough()
    If mnuStrikeThru Then
        msEdit.SelStrikeThru = False
        mnuStrikeThru = False
        tlbar.Buttons("StrikeThru").Value = tbrUnpressed
    Else
        msEdit.SelStrikeThru = True
        mnuStrikeThru = True
        tlbar.Buttons("StrikeThru").Value = tbrPressed
    End If
End Sub


Private Sub msUnderline()
    If mnuUnderline Then
        msEdit.SelUnderline = False
        mnuUnderline = False
        tlbar.Buttons("Underline").Value = tbrUnpressed
    Else
        msEdit.SelUnderline = True
        mnuUnderline = True
        tlbar.Buttons("Underline").Value = tbrPressed
    End If
End Sub
Sub SetButtons()
    
    If msEdit.SelBold = True Then
        mnuBold = True
        tlbar.Buttons("Bold").Value = tbrPressed
    Else
        mnuBold = False
        tlbar.Buttons("Bold").Value = tbrUnpressed
    End If
    
    If msEdit.SelItalic = True Then
        mnuItalic = True
        tlbar.Buttons("Italic").Value = tbrPressed
    Else
        mnuItalic = False
        tlbar.Buttons("Italic").Value = tbrUnpressed
    End If
    
    If msEdit.SelUnderline = True Then
        mnuUnderline = True
        tlbar.Buttons("Underline").Value = tbrPressed
    Else
        mnuUnderline = False
        tlbar.Buttons("Underline").Value = tbrUnpressed
    End If
    
    If msEdit.SelStrikeThru = True Then
        mnuStrikeThru = True
        tlbar.Buttons("StrikeThru").Value = tbrPressed
    Else
        mnuStrikeThru = False
        tlbar.Buttons("StrikeThru").Value = tbrUnpressed
    End If
    
    If msEdit.SelAlignment = 0 Then
        tlbar.Buttons("Left").Value = tbrPressed
        tlbar.Buttons("Center").Value = tbrUnpressed
        tlbar.Buttons("Right").Value = tbrUnpressed
    Else
        If msEdit.SelAlignment = 1 Then
            tlbar.Buttons("Left").Value = tbrUnpressed
            tlbar.Buttons("Center").Value = tbrUnpressed
            tlbar.Buttons("Right").Value = tbrPressed
        Else
            tlbar.Buttons("Left").Value = tbrUnpressed
            tlbar.Buttons("Center").Value = tbrPressed
            tlbar.Buttons("Right").Value = tbrUnpressed
        End If
        
    End If
    
End Sub
Private Sub msUndo()
    Undo
    tlbar.Buttons("Redo").Enabled = True
End Sub

Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Save"
             SaveNow
        Case "New"
            msNew
        Case "Open"
            msOpen
        Case "Print"
            msPrint
        Case "Cut"
            msCut
        Case "Copy"
            msCopy
        Case "Paste"
            msPaste
        Case "Undo"
            msUndo
        Case "Redo"
            msRedo
        Case "Color"
            msFontColor
        Case "Font"
            msFont
        Case "StrikeThru"
            msStrikeThrough
        Case "Underline"
            msUnderline
        Case "Italic"
            msItalic
        Case "Bold"
            msBold
        Case "Left"
            msLeft
        Case "Center"
            msCenter
        Case "Right"
            msRight
        Case "IncreaseIndent"
            msIncreaseIndent
        Case "DecreaseIndent"
            msDecreaseIndent
        Case "InsertPicture"
              InsertPicture ""
        Case "InsertObject"
              InsertObject
        Case "InsertText"
             InsertFile ""
        Case "Replace"
             FindReplaceDialog
        End Select
End Sub

Private Sub tlbar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
       ' PopupMenu mnuView
    End If
End Sub

Public Sub Undo()
    'if the Index is = to 0, then It shouldn't undo anymore
    If gintIndex = 0 Then Exit Sub
    
    'This is the basic undo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    msEdit.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
    
End Sub
Public Sub Redo()
    'This is the basic redo
    gblnIgnoreChange = True
    gintIndex = gintIndex + 1
    On Error Resume Next
    msEdit.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub msEdit_Change()
    If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
        gstrStack(gintIndex) = msEdit.TextRTF
    End If
    Saved = False
    tlbar.Buttons("Save").Enabled = True
    tlbar.Buttons("Print").Enabled = True
    tlbar.Buttons("Undo").Enabled = True
    ShowStatus
End Sub

Private Sub msEdit_Click()
    SetButtons
    
End Sub

Private Sub msEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If msEdit.SelText <> "" Then
        tlbar.Buttons("Copy").Enabled = True
        tlbar.Buttons("Cut").Enabled = True
    Else
        tlbar.Buttons("Copy").Enabled = False
        tlbar.Buttons("Cut").Enabled = False
    End If
    
    SetButtons
    
End Sub

Private Sub msEdit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If msEdit.SelText <> "" Then
        tlbar.Buttons("Copy").Enabled = True
        tlbar.Buttons("Cut").Enabled = True
    Else
        tlbar.Buttons("Copy").Enabled = False
        tlbar.Buttons("Cut").Enabled = False
    End If
    
    SetButtons
    
End Sub
Private Sub UserControl_Resize()
    On Error Resume Next
    msEdit.Move 0, 0 + tlbar.Height, UserControl.ScaleWidth, UserControl.ScaleHeight - tlbar.Height - StatusBar1.Height
End Sub
Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = msEdit.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    On Error Resume Next
    msEdit.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = m_FontUnderline
End Property
Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    m_FontUnderline = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = m_FontStrikethru
End Property
Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    m_FontStrikethru = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = m_FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    m_FontSize = New_FontSize
    PropertyChanged "FontSize"
End Property
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = m_FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    m_FontName = New_FontName
    PropertyChanged "FontName"
End Property
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = m_FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    m_FontItalic = New_FontItalic
    PropertyChanged "FontItalic"
End Property
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = m_FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    m_FontBold = New_FontBold
    PropertyChanged "FontBold"
End Property
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = msEdit.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set msEdit.Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in an object."
    Text = msEdit.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    msEdit.Text() = New_Text
    PropertyChanged "Text"
End Property
Public Property Get TextRTF() As String
    TextRTF = msEdit.TextRTF
End Property

Public Property Let TextRTF(ByVal New_TextRTF As String)
    msEdit.TextRTF = TextRTF
    PropertyChanged "TextRTF"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
'    m_Enabled = m_def_Enabled
    m_ForeColor = m_def_ForeColor
    m_FontUnderline = m_def_FontUnderline
    m_FontStrikethru = m_def_FontStrikethru
    m_FontSize = m_def_FontSize
    m_FontName = m_def_FontName
    m_FontItalic = m_def_FontItalic
    m_FontBold = m_def_FontBold
    m_TextRTF = m_def_TextRTF
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    msEdit.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    msEdit.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_FontUnderline = PropBag.ReadProperty("FontUnderline", m_def_FontUnderline)
    m_FontStrikethru = PropBag.ReadProperty("FontStrikethru", m_def_FontStrikethru)
    m_FontSize = PropBag.ReadProperty("FontSize", m_def_FontSize)
    m_FontName = PropBag.ReadProperty("FontName", m_def_FontName)
    m_FontItalic = PropBag.ReadProperty("FontItalic", m_def_FontItalic)
    m_FontBold = PropBag.ReadProperty("FontBold", m_def_FontBold)
    Set msEdit.Font = PropBag.ReadProperty("Font", Ambient.Font)
    msEdit.Text = PropBag.ReadProperty("Text", "")
    m_TextRTF = PropBag.ReadProperty("TextRTF", m_def_TextRTF)
    msEdit.MaxLength = PropBag.ReadProperty("MaxLength", 0)
        
    tlbar.Buttons("Open").Visible = PropBag.ReadProperty("Open_Visible", m_def_Open_Visible)
    tlbar.Buttons("Save").Visible = PropBag.ReadProperty("Save_Visible", m_def_Save_Visible)
    tlbar.Buttons("New").Visible = PropBag.ReadProperty("New_Visible", m_def_New_Visible)
    tlbar.Buttons("Print").Visible = PropBag.ReadProperty("Print_Visible", m_def_Print_Visible)
    
    msEdit.BackColor = PropBag.ReadProperty("BackColor", &HC0FFFF)
        
    msEdit.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", msEdit.BackColor, &H80000005)
    Call PropBag.WriteProperty("BorderStyle", msEdit.BorderStyle, 1)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("FontUnderline", m_FontUnderline, m_def_FontUnderline)
    Call PropBag.WriteProperty("FontStrikethru", m_FontStrikethru, m_def_FontStrikethru)
    Call PropBag.WriteProperty("FontSize", m_FontSize, m_def_FontSize)
    Call PropBag.WriteProperty("FontName", m_FontName, m_def_FontName)
    Call PropBag.WriteProperty("FontItalic", m_FontItalic, m_def_FontItalic)
    Call PropBag.WriteProperty("FontBold", m_FontBold, m_def_FontBold)
    Call PropBag.WriteProperty("Font", msEdit.Font, Ambient.Font)
    Call PropBag.WriteProperty("Text", msEdit.Text, "")
    Call PropBag.WriteProperty("TextRTF", m_TextRTF, m_def_TextRTF)
    Call PropBag.WriteProperty("MaxLength", msEdit.MaxLength, 0)
    
    Call PropBag.WriteProperty("Open_Visible", tlbar.Buttons("Open").Visible, m_def_Open_Visible)
    Call PropBag.WriteProperty("Save_Visible", tlbar.Buttons("Save").Visible, m_def_Save_Visible)
    Call PropBag.WriteProperty("New_Visible", tlbar.Buttons("New").Visible, m_def_New_Visible)
    Call PropBag.WriteProperty("Print_Visible", tlbar.Buttons("Print").Visible, m_def_Print_Visible)
    
    Call PropBag.WriteProperty("BackColor", msEdit.BackColor, &HC0FFFF)
        
    Call PropBag.WriteProperty("Enabled", msEdit.Enabled, True)
End Sub
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets a value indicating whether there is a maximum number of characters a RichTextBox control can hold and, if so, specifies the maximum number of characters."
    MaxLength = msEdit.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    msEdit.MaxLength = New_MaxLength
    PropertyChanged "MaxLength"
End Property

Public Sub msOpen()
    
    On Error GoTo errHandling

    With CommonDialog1
        .Flags = cdlOFNFileMustExist
        .Filter = "Text Documents|*.txt;*.rtf"
        .ShowOpen
    End With
    
    msEdit.LoadFile CommonDialog1.FileName
    
    With tlbar
    
        .Buttons("Print").Enabled = True
        .Buttons("Cut").Enabled = False
        .Buttons("Copy").Enabled = False
        .Buttons("Undo").Enabled = False
        .Buttons("Redo").Enabled = False
        .Buttons("Paste").Enabled = False
        .Buttons("Save").Enabled = False
        
    End With
   
    Exit Sub

errHandling:

    If Err.Number <> cdlCancel Then
        MsgBox Err.Description
    End If
End Sub
Public Property Get Open_Visible() As Boolean
    Open_Visible = tlbar.Buttons("Open").Visible
End Property

Public Property Let Open_Visible(ByVal New_Open_Visible As Boolean)
    tlbar.Buttons("Open").Visible = New_Open_Visible
    PropertyChanged "Open_Visible"
End Property
Public Property Get Save_Visible() As Boolean
    Save_Visible = tlbar.Buttons("Save").Visible
End Property

Public Property Let Save_Visible(ByVal New_Save_Visible As Boolean)
    tlbar.Buttons("Save").Visible = New_Save_Visible
    PropertyChanged "Save_Visible"
End Property
Public Property Get New_Visible() As Boolean
    New_Visible = tlbar.Buttons("New").Visible
End Property
Public Property Let New_Visible(ByVal New_New_Visible As Boolean)
    tlbar.Buttons("New").Visible = New_New_Visible
    PropertyChanged "New_Visible"
End Property
Public Property Get Print_Visible() As Boolean
    Print_Visible = tlbar.Buttons("Print").Visible
End Property

Public Property Let Print_Visible(ByVal New_Print_Visible As Boolean)
    tlbar.Buttons("Print").Visible = New_Print_Visible
    PropertyChanged "Print_Visible"
End Property

Public Sub PrintText()
   ' Uncomment the following statements to expose
   ' the Print Range section of the Print dialog box
   
   CommonDialog1.Min = 0
   CommonDialog1.Max = 9999
   
   CommonDialog1.CancelError = True
   
   On Error Resume Next
   
   CommonDialog1.ShowPrinter
   If Err.Number = 32755 Then
      Exit Sub
   End If
   
   UserControl.MousePointer = vbHourglass
   UserControl.Refresh
   
   LeftMargin = 25: RightMargin = 20
   TopMargin = 20: BottomMargin = 30
   
   If CommonDialog1.Orientation = cdlPortrait Then
      Printer.Orientation = cdlPortrait
   Else
      Printer.Orientation = cdlLandscape
   End If
   
   If PRN Is Printer Then MDIForm1.Show
   
   DoEvents
   
'   SetPrinterFont
   
   PrintWidth = Printer.ScaleWidth - LeftMargin - RightMargin
   PrintHeight = Printer.ScaleHeight - TopMargin - BottomMargin
   
   Dim iChar As Integer
   Dim newitem As String: newitem = ""
   Dim moreWords As Boolean: moreWords = True
   Dim nextWord As String
   
   ClearScreen
   
   PRN.CurrentX = LeftMargin
   PRN.CurrentY = TopMargin
   
   Dim txtLines() As String
   
   txtLines = Split(msEdit.Text, vbCrLf)
   Dim iLine As Integer
   Dim prnLine As String
   
   For iLine = 0 To UBound(txtLines)
      moreWords = True
      iChar = 1
      prnLine = txtLines(iLine)
      PRN.CurrentX = LeftMargin
      While moreWords
         If PRN.CurrentY + Printer.TextHeight("A") > Printer.ScaleHeight - BottomMargin Then
            If Not PRN Is Printer Then
               UserControl.MousePointer = vbDefault
               Dim reply As VbMsgBoxResult
               reply = MsgBox("Continue Preview?", vbYesNo)
               If reply = vbYes Then
                  ClearScreen
               Else
                  Exit Sub
               End If
            Else
               ClearScreen
            End If
         End If
         
         nextWord = GetNextWord(prnLine, iChar)
         iChar = iChar + Len(nextWord)
         If PRN.TextWidth(newitem & nextWord) < PrintWidth Then
            newitem = newitem & nextWord
         Else
            PrintAlignedString newitem & LineContSymbol, txtMain.alignment
            newitem = nextWord
            PRN.CurrentX = 1 * LeftMargin
         End If
         If iChar > Len(prnLine) Then
             PrintAlignedString newitem, txtMain.alignment
             moreWords = False
             newitem = ""
         End If
      Wend
    Next
    UserControl.MousePointer = vbDefault
    PRN.Print newitem
    
End Sub


Private Function GetNextWord(ByVal str As String, ByVal pos As Integer) As String
      Dim nextWord As String
      While pos <= Len(str) And Mid(str, pos, 1) <> " "
         nextWord = nextWord & Mid(str, pos, 1)
         pos = pos + 1
      Wend
      While pos <= Len(str) And Mid(str, pos, 1) = " "
         nextWord = nextWord & Mid(str, pos, 1)
         pos = pos + 1
      Wend
      GetNextWord = nextWord
End Function

Private Sub ClearScreen()
   If PRN Is Printer Then
      PRN.NewPage
   Else
      PRN.Cls
      PRN.Line (LeftMargin, TopMargin)- _
         (LeftMargin + PrintWidth, _
         TopMargin + PrintHeight), _
         RGB(255, 255, 255), BF
   End If
   PRN.CurrentX = LeftMargin
   PRN.CurrentY = TopMargin
End Sub


Private Sub PrintAlignedString(ByVal str As String, ByVal alignment As Integer)
      
      Select Case alignment
      Case 0
         PRN.CurrentX = LeftMargin
      Case 1
         PRN.CurrentX = LeftMargin + PrintWidth - PRN.TextWidth(str)
      Case 2
         PRN.CurrentX = LeftMargin + (PrintWidth - PRN.TextWidth(str)) / 2
      End Select
      PRN.Print str
End Sub


Public Sub msPrint()

Dim PrintableWidth As Long
Dim PrintableHeight As Long
Dim x As Single

   ' Uncomment the following statements to expose
   ' the Print Range section of the Print dialog box
   
   ' CommonDialog1.Min = 0
   ' CommonDialog1.Max = 9999
      
   On Error Resume Next
   CommonDialog1.CancelError = True
   
   CommonDialog1.ShowPrinter

   If Err.Number = 32755 Then
      Exit Sub
   End If
   
   If CommonDialog1.Orientation = cdlPortrait Then
      Printer.Orientation = cdlPortrait
   Else
      Printer.Orientation = cdlLandscape
   End If

    ' Initialize the Printer object.
    x = Printer.TwipsPerPixelX

    ' Tell the RichTextBox to base it's display on the printer.
    WYSIWYG_RTF msEdit, QuarterInch, QuarterInch, QuarterInch, QuarterInch, PrintableWidth, PrintableHeight

    ' Set the form width to match the line width
    
    PrintRTF msEdit, AnInch, AnInch, AnInch, AnInch, "This is the Header", "This is the Footer"

End Sub
' Make a RichTextBox control display itself
' using the same parameters as the default printer.
Private Sub WYSIWYG_RTF(RTF As RichTextBox, LeftMarginWidth As Long, RightMarginWidth As Long, TopMarginWidth As Long, BottomMarginWidth As Long, PrintableWidth As Long, PrintableHeight As Long)

Dim LeftOffset As Long
Dim LeftMargin As Long
Dim RightMargin As Long
Dim TopOffset As Long
Dim TopMargin As Long
Dim BottomMargin As Long
Dim PrinterhDC As Long
Dim r As Long

    ' Start a print job to initialize the Printer object.
    Printer.Print " "
    Printer.ScaleMode = vbTwips

    ' Get the left offset to the printable area on the page in twips.
    LeftOffset = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX)
    LeftOffset = Printer.ScaleX(LeftOffset, vbPixels, vbTwips)

    ' Calculate the left and right margins.
    LeftMargin = LeftMarginWidth - LeftOffset
    RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
    
    ' Calculate the printable width.
    PrintableWidth = RightMargin - LeftMargin

    ' Get the top offset to the printable area on the page.
    TopOffset = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY)
    TopOffset = Printer.ScaleX(TopOffset, vbPixels, vbTwips)

    ' Calculate the top and bottom margins.
    TopMargin = TopMarginWidth - TopOffset
    BottomMargin = (Printer.Height - BottomMarginWidth) - TopOffset

    ' Calculate the printable height.
    PrintableHeight = BottomMargin - TopMargin

    ' Create an hDC for the Printer. This DC must
    ' remain for the control to keep up the
    ' WYSIWYG settings.
    PrinterhDC = CreateDC(Printer.DriverName, _
        Printer.DeviceName, 0, 0)

    ' Tell the RichTextBox to base it's display
    ' on the printer at the desired line width.
    r = SendMessage(RTF.hWnd, EM_SETTARGETDEVICE, PrinterhDC, ByVal PrintableWidth)

    ' Cancel the temporary print job used to get
    ' the printer information.
    Printer.KillDoc
    
End Sub
' Print the contents of a RichTextBox control using the indicated margins.
Private Sub PrintRTF(RTF As RichTextBox, LeftMarginWidth As Long, TopMarginHeight, RightMarginWidth, BottomMarginHeight, Optional ByVal header_text As String = "", Optional ByVal footer_text As String = "")

Dim LeftOffset As Long
Dim TopOffset As Long
Dim LeftMargin As Long
Dim TopMargin As Long
Dim RightMargin As Long
Dim BottomMargin As Long
Dim fr As FormatRange
Dim rcDrawTo As Rect
Dim rcPage As Rect
Dim TextLength As Long
Dim NextCharPosition As Long
Dim r As Long
Dim txt As String

    ' Start a print job to get a valid Printer.hDC.
    Printer.Print " "
    Printer.ScaleMode = vbTwips

    ' Get the offset to the page's printable area.
    LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX), vbPixels, vbTwips)
    TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY), vbPixels, vbTwips)

    ' Calculate the left, top, right, and bottom margins.
    LeftMargin = LeftMarginWidth - LeftOffset
    TopMargin = TopMarginHeight - TopOffset
    RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
    BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset

    ' Set the printable area rectangle.
    rcPage.Left = 0
    rcPage.Top = 0
    rcPage.Right = Printer.ScaleWidth
    rcPage.Bottom = Printer.ScaleHeight

    ' Set the rectangle in which to print
    ' (relative to the printable area).
    rcDrawTo.Left = LeftMargin
    rcDrawTo.Top = TopMargin
    rcDrawTo.Right = RightMargin
    rcDrawTo.Bottom = BottomMargin

    ' Set up the print instructions.
    fr.hdc = Printer.hdc        ' Use the same DC for measuring and rendering
    fr.hdcTarget = Printer.hdc  ' Point at printer hDC
    fr.rc = rcDrawTo            ' Indicate the area on page to draw to
    fr.rcPage = rcPage          ' Indicate entire size of page
    fr.chrg.cpMin = 0           ' Indicate start of text through
    fr.chrg.cpMax = -1          ' end of the text

    ' Get length of text in RTF
    TextLength = Len(RTF.Text)

    ' Loop printing each page until done
    Do
        ' Print the page by sending EM_FORMATRANGE message
        NextCharPosition = SendMessage(RTF.hWnd, EM_FORMATRANGE, True, fr)
        If NextCharPosition >= TextLength Then Exit Do  'If done then exit
        fr.chrg.cpMin = NextCharPosition ' Starting position for next page

        ' Page number.
        txt = "- " & Format$(Printer.Page) & " -"
        Printer.CurrentX = RightMargin - Printer.TextWidth(txt)
        Printer.CurrentY = Printer.ScaleTop + 1440 * 0.25
        Printer.Print Format$(txt)

        ' Header.
        If Len(header_text) > 0 Then
            Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(header_text)) / 2
            Printer.CurrentY = Printer.ScaleTop + TopMargin / 2
            Printer.Print Format$(header_text)
        End If

        ' Footer.
        If Len(footer_text) > 0 Then
            Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(footer_text)) / 2
            Printer.CurrentY = Printer.ScaleTop + Printer.ScaleHeight + TopMargin / 2
            Printer.Print Format$(footer_text)
        End If
    
        Printer.NewPage                  ' Move on to next page
        Printer.Print " " ' Re-initialize hDC
        fr.hdc = Printer.hdc
        fr.hdcTarget = Printer.hdc
    Loop

    ' End the print job.
    Printer.EndDoc

    ' Allow the RTF to free up memory.
    r = SendMessage(RTF.hWnd, EM_FORMATRANGE, False, ByVal CLng(0))
    
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color of an object."
    BackColor = msEdit.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    msEdit.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Private Sub InsertPictureInRichTextBox(RTB As RichTextBox, Picture As StdPicture)
    ' copy into the clipboard
    ' Copy the picture into the clipboard.
    Clipboard.Clear
    Clipboard.SetData Picture
    ' paste into the RichTextBox control
    SendMessage RTB.hWnd, WM_PASTE, 0, 0
    Clipboard.Clear
    
End Sub
Public Sub InsertPicture(sPictureName As String)
   
   On Error Resume Next
   
   ' If the user didn't provide a name then ask for one
   If Len(sPictureName) = 0 Then
       With CommonDialog1
          .Filter = "Office Pictures " & _
          "(*.gif,*.jpg,*.bmp)|*.gif;*.jpg;*.bmp"
          .FilterIndex = 1
          .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
          
          .FileName = ""
          .ShowOpen
          sPictureName = .FileName
       End With
   End If
   
 ' If the user didn't cancel, open the file...
  If Len(sPictureName) Then
      InsertPictureInRichTextBox msEdit, LoadPicture(sPictureName)
  End If
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = msEdit.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    msEdit.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Sub InsertObject()
    ' Show the InsertObject dialog
    frmDummy.OLE1.InsertObjDlg
    If Len(frmDummy.OLE1.Class) Then
        ' If the user selected an object, show it in the RichTextBox control
        msEdit.OLEObjects.Add , , , frmDummy.OLE1.Class
        frmDummy.OLE1.Class = ""
    End If
End Sub
Public Sub InsertFile(sFileName As String)
   
   On Error Resume Next
   
   ' If the user didn't provide a name then ask for one
   If Len(sFileName) = 0 Then
       With CommonDialog1
          .Filter = "Text Files " & _
          "(*.rtf)|*.rtf"
          .FilterIndex = 1
          .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
          
          .FileName = ""
          .ShowOpen
          sFileName = .FileName
       End With
   End If
   
 ' If the user didn't cancel, open the file...
  If Len(sFileName) Then
      Richinsert.LoadFile sFileName
      msEdit.SelText = Richinsert.TextRTF
      Richinsert.TextRTF = ""
      tlbar.Buttons("Undo").Enabled = True
  End If
End Sub


Private Sub ShowStatus()
  
  Dim lCurLine As Long
  Dim ilineCount As Long
'  Dim GetFirstLineVisible As Long
  Dim CurrentColumn As Long
      
   ilineCount = SendMessage(msEdit.hWnd, EM_GETLINECOUNT, 0&, 0&)
'   GetFirstLineVisible = SendMessage(msEdit.hWnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)

   ' Current Line
   lCurLine = 1 + msEdit.GetLineFromChar(msEdit.SelStart)
   
   ' Column
   CurrentColumn = SendMessage(msEdit.hWnd, EM_LINEINDEX, ByVal lCurLine - 1, 0&)
   CurrentColumn = (msEdit.SelStart) - CurrentColumn
      
   StatusBar1.Panels(2).Text = lCurLine
   StatusBar1.Panels(4).Text = CurrentColumn + 1
   StatusBar1.Panels(6).Text = ilineCount
    
End Sub
Public Sub FindReplace(strFindString As String, strReplaceString As String)

Dim lFoundPos As Long           'Position of first character
                                'of match
Dim lFindLength As Long         'Length of string to find
Dim lFindRLength As Long        'Length of replacement string
Dim bTryAgain As Boolean
Dim lOriginalSelStart As Long
Dim lOriginalSelLength As Long
Dim iMatchCount As Long

With msEdit
    'Save the insertion points current location and length
    lOriginalSelStart = .SelStart
    lOriginalSelLength = .SelLength

    'Cache the length of the string to find
    lFindLength = Len(strFindString)
    
    lFindRLength = Len(strReplaceString)

    'Attempt to find the first match
    lFoundPos = .Find(strFindString, 0, , rtfNoHighlight)

    'If the First Item in the msEdit is found then allow a second
    'Match to be attempted
    bTryAgain = IIf(lFoundPos = -1, False, True)

    Do While lFoundPos > 0 Or bTryAgain

        'When the last character is found exit
        If lFoundPos = Len(.Text) Then Exit Do

        'Reset the Try Again
        bTryAgain = False

        'Track Matches
        iMatchCount = iMatchCount + 1

        msEdit.SelStart = lFoundPos
        'The SelLength property is set to 0 as soon as you change SelStart
        .SelLength = lFindLength
        .SelText = strReplaceString

        'Attempt to find the next match
        lFoundPos = .Find(strFindString, lFoundPos + lFindRLength, , rtfNoHighlight)
        
    Loop

    'Restore the insertion point to its original location and length
    .SelStart = lOriginalSelStart
    .SelLength = lOriginalSelLength
End With

'Return the number of matches
'FindReplace = iMatchCount

End Sub

Public Sub FindReplaceDialog()

  Set mReplaceCallback = frmReplace
  mReplaceCallback.Show vbModal

End Sub

Private Sub mReplaceCallback_MyCallBack(ByVal fText As String, ByVal rText As String, bCancel As Boolean)
    If Not bCancel Then
       If fText <> "" Then
            FindReplace fText, rText
       End If
    End If
End Sub

