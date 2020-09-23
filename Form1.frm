VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Simple Search"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Cbo_Operator 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2640
      List            =   "Form1.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2760
      Width           =   1695
   End
   Begin VB.ComboBox Cbo_Category 
      Height          =   315
      ItemData        =   "Form1.frx":002C
      Left            =   960
      List            =   "Form1.frx":0039
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      DataField       =   "ContactTitle"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      DataField       =   "ContactName"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      DataField       =   "CompanyName"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customers"
      Top             =   3480
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comparison Operator :"
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Category to Search :"
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Contact Title :"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Contact Name :"
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Company Name :"
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'If your newbie to databases like me, then you might
'wanna check out this cool URL
'http://www.vb-world.net/databases/dbtutorial/
'Its written in a very easy to understand syntax.

Private Sub Command1_Click()
On Error GoTo OOPS
Dim My_Query As String

'Remark :
'The database KNOWS if a field is numeric or not
'IF the field you are trying to search is a string
'then you must enclose it with "MySting" or 'Mystring'
'Otherwise if the value is numeric, then the database
'expects a query like WHERE CustomerID=1
'and NOT! WHERE CustomerID='1'
'otherwise it crashes with an error

MyQuery = Cbo_Category & " " & Cbo_Operator & " '" & Text1.Text & "'"

Data1.RecordSource = "SELECT * FROM Customers WHERE " & MyQuery
'This just means:
' If the search is positive, i want you to return me all
' fields "SELECT" "*", and now i want you to
' search the table "Customers" "WHERE" the entry
' "Cbo_Category"'is equal to "Cbo_Operator" what i said
' "MyQuery".
Data1.Refresh

'We need to move to the last entry and back to the first
'or else we get a wrong record count, this is a bug i think.
Data1.Recordset.MoveLast: Data1.Recordset.MoveFirst
MsgBox Data1.Recordset.RecordCount & " matches found."
Exit Sub
OOPS:
MsgBox "No Records Found"
End Sub

Private Sub Form_Load()
MsgBox "If you use the LIKE method, then you can use the *(wildcard) to complete the query." & vbNewLine & "Like do a search for Ma* or M*ria or *m* etc."
'Change this to where your NWIND.MDB is installed
'This DB ships by default with vb5 and 6
Data1.DatabaseName = "C:\Program Files\Microsoft Visual Studio\VB98\NWIND.MDB"
Cbo_Category.ListIndex = 0
Cbo_Operator.ListIndex = 0
End Sub
