VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RTBMTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This class was written by Robin Mckay on 5th July 2001
' It was designed to add extra functionality to the
' text box originally supplied by Microsoft... thank you
' for using this class.

' Robin Mckay

Public Property Let Height(Height As Integer)
form1.text1.Height = Height
End Property

Public Property Let Width(Width As Integer)
form1.text1.Width = Width
End Property

Public Property Let Backcolor(Backcolor As Integer)
form1.text1.Backcolor = Backcolor
End Property

Public Property Let Font(Font As String)
form1.text1.Font = Font
End Property

Public Property Let FontSize(FontSize As Integer)
form1.text1.Font.Size = FontSize
End Property

Public Property Let Bold(Bold As Boolean)
If Bold = True Then
    form1.text1.Font.Bold = True
ElseIf Bold = False Then
    form1.text1.Font.Bold = False
End If
End Property
Public Property Let Italic(Italic As Boolean)
If Italic = True Then
    form1.text1.Font.Italic = True
ElseIf Italic = False Then
    form1.text1.Font.Italic = False
End If
End Property
Public Property Let Underline(Underline As Boolean)
If Underline = True Then
    form1.text1.Font.Underline = True
ElseIf Underline = False Then
    form1.text1.Font.Underline = False
End If
End Property
Public Property Let Strikethru(Strikethru As Boolean)
If Strikethru = True Then
    form1.text1.Font.Strikethrough = True
ElseIf Strikethru = False Then
    form1.text1.Font.Strikethru = False
End If
End Property
Public Property Let Enabled(Enabled As Boolean)
If Enabled = True Then
    form1.text1.Enabled = True
ElseIf Enabled = False Then
    form1.text1.Enabled = False
End If
End Property

Function SaveFile()
On Error GoTo cancelerr:
form1.commondialog1.cancelerror = True
form1.commondialog1.Filter = "Text Files|*.txt|All Files|*.*"
form1.commondialog1.ShowSave
Open form1.commondialog1.FileName For Output As #1
Print #1, form1.text1.Text
Close #1
cancelerr:
form1.Show
End Function

Function SaveFileAsHTML()
On Error GoTo cancelerr:
form1.commondialog1.cancelerror = True
form1.commondialog1.Filter = "HTML Files|*.html|HTM Files|*.htm|All Files|*.*"
form1.commondialog1.ShowSave
Open form1.commondialog1.FileName For Output As #1
Print #1, form1.text1.Text
Close #1
cancelerr:
form1.Show
End Function

Function ViewInWord()
Dim word
Set word = CreateObject("word.basic")

word.appshow
word.filenew
word.Insert form1.text1.Text
End Function

Function DisplayAbout()
MsgBox "This class was developed by Robin McKay to add more functionality to the text box control. It is to be used in conjunction with the text box control. A text box must be active on the form for this program to work as well as a common dialog control and some of the functions require you have MS Word installed on your computer." + vbCrLf + vbCrLf + "I had some troubles adding some of the text box features to the class, so if you could help me add them, I would be very greatful." + vbCrLf + vbCrLf + "This class is free for everyone to use." + vbCrLf + vbCrLf + "For those of you wondering why I named this class as I did do, simply because RTBM stands for Robin Thomas Benjamin Mckay", vbInformation, "About RTBMTextBox"
End Function

Function ChangeTextForeColor()
On Error Resume Next

    With form1.commondialog1
        .cancelerror = True
        .dialogtitle = "Change Text Fore Color"
        .Flags = cdlCCFullOpen
        .ShowColor
    End With
    form1.text1.ForeColor = form1.commondialog1.Color
End Function

Public Property Let Name(Name As String)
    form1.text1.Name = Name
End Property

Public Property Let Borderstyle(Borderstyle As Integer)
form1.text1.Borderstyle = Borderstyle
End Property

Public Property Let CausesValidation(CausesValidation As Boolean)
If CausesValidation = True Then
    form1.text1.CausesValidation = True
ElseIf CausesValidation = False Then
    form1.text1.CausesValidation = False
End If
End Property

Public Property Let Alignment(Alignment As Integer)
form1.text1.Alignment = Alignment
End Property

Public Property Let Datafield(Datafield As Integer)
form1.text1.Datafield = Datafield
End Property
Public Property Let DraIicon(DragIcon As String)
form1.text1.DragIcon = DragIcon
End Property

Function SpellCheck()
' This uses the MS Word program...
On Error Resume Next
Dim x
Set x = CreateObject("word.application")

x.Visible = False
x.Documents.Add
x.Selection.Text = form1.text1.Text
x.ActiveDocument.CheckSpelling
form1.text1.Text = x.Selection.Text

x.Quit
Set x = Nothing
End Function

Function SendMail()
' This uses the MS Word program...
On Error Resume Next
Dim x
Set x = CreateObject("word.application")

x.Visible = False
x.Documents.Add
x.Selection.Text = form1.text1.Text
x.ActiveDocument.SendMail
End Function

Public Property Let DragMode(DragMode As Integer)
form1.text1.DragMode = DragMode
End Property

Public Property Let HelpContextID(HelpContextID As Integer)
form1.text1.HelpContextID = HelpContextID
End Property
Public Property Let HideSelection(HideSelection As Boolean)
If HideSelection = True Then
    form1.text1.HideSelection = True
ElseIf HideSelection = False Then
    form1.text1.HideSelection = False
End If
End Property

Public Property Let Left(Left As Integer)
form1.text1.Left = Left
End Property

Public Property Let LinkMode(LinkMode As Integer)
form1.text1.LinkMode = LinkMode
End Property
Public Property Let LinkTimeOut(LinkTimeOut As Integer)
form1.text1.LinkTimeOut = LinkTimeOut
End Property
Public Property Let Locked(Locked As Boolean)
If Locked = False Then
    form1.text1.Locked = False
ElseIf Locked = True Then
    form1.text1.Locked = True
End If
End Property
Public Property Let MaxLength(MaxLength As Boolean)
form1.text1.MaxLength = MaxLength
End Property

Public Property Let MousePointer(MousePointer As Integer)
form1.text1.MousePointer = MousePointer
End Property

Public Property Let OLEDragMode(OLEDragMode As Integer)
form1.text1.OLEDragMode = OLEDragMode
End Property

Public Property Let OLEDropMode(OLEDropMode As Integer)
form1.text1.OLEDropMode = OLEDropMode
End Property

Public Property Let RightToLeft(RightToLeft As Boolean)
form1.text1.RightToLeft = RightToLeft
End Property

Function PreviewAsWebPage()
form1.commondialog1.FileName = "c:\windows\test.htm"
Open form1.commondialog1.FileName For Output As #1
Print #1, form1.text1.Text
Close #1
Shell ("start c:\windows\test.htm")
End Function

Public Property Let TabIndex(TabIndex As Integer)
form1.text1.TabIndex = TabIndex
End Property

Public Property Let TabStop(TabStop As Boolean)
If TabStop = False Then
    form1.text1.TabStop = False
ElseIf TabStop = True Then
    form1.text1.TabStop = True
End If
End Property

Public Property Let Text(Text As String)
form1.text1.Text = Text
End Property

Public Property Let ToolTipText(ToolTipText As String)
form1.text1.ToolTipText = ToolTipText
End Property

Public Property Let Top(Top As Integer)
form1.text1.Top = Top
End Property

Public Property Let Visible(Visible As Boolean)
If Visible = True Then
    form1.text1.Visible = True
ElseIf Visible = False Then
    form1.text1.Visible = False
End If
End Property
Public Property Let WhatsThisHelpID(WhatsThisHelpID As Integer)
form1.text1.WhatsThisHelpID = WhatsThisHelpID
End Property

Sub LinkExecute(Command As String)
form1.text1.LinkExecute (Command)
End Sub

Sub Move(Left As Integer, Top As Integer, Height As Integer, Width As Integer)
form1.text1.Left = Left
form1.text1.Top = Top
form1.text1.Height = Height
form1.text1.Width = Width
End Sub

Public Property Let SelText(SelText As String)
form1.text1.SelText = SelText
End Property

Public Property Let SelLength(SelLength As Integer)
form1.text1.SelLength = SelLength
End Property

Public Property Let SelStart(SelStart As Integer)
form1.text1.SelStart = SelStart
End Property

Function SelectAll()
form1.text1.SetFocus
form1.text1.SelStart = 0
form1.text1.SelLength = Len(form1.text1)
End Function

Function Cut()
Clipboard.Clear
form1.text1.SetFocus
Clipboard.SetText form1.text1.SelText
form1.text1.SelText = ""
End Function

Function Copy()
Clipboard.Clear
form1.text1.SetFocus
Clipboard.SetText form1.text1.SelText
End Function

Function Paste()
form1.text1.SelText = Clipboard.GetText
End Function

Function LoadFile()
With form1.commondialog1
    .cancelerror = True
    .Filter = "TXT Files|*.txt|All Files|*.*"
    .showopen
End With
form1.richtextbox1.LoadFile (form1.commondialog1.FileName)
form1.text1.Text = form1.richtextbox1.Text
End Function

Function PrintText()
On Error Resume Next
Printer.Print form1.text1.Text
Printer.EndDoc
End Function

Function TimeDate()
form1.text1.SelText = Now
End Function

Function Delete()
form1.text1.SelText = ""
End Function

Function InsertTextFile()
On Error GoTo cancelerr:
    With form1.commondialog1
        .cancelerror = True
        .dialogtitle = "Insert Text File"
        .Filter = "Text File|*.txt|All Files|*.*"
        .showopen
    End With
    form1.richtextbox1.LoadFile (form1.commondialog1.FileName)
    form1.text1.Text = form1.richtextbox1.Text
cancelerr:
form1.Show
End Function














