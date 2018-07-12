VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command1_Click()
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Show
        
        file = .SelectedItems(1)
    End With
    
    Image0.Picture = file
    
    'Modified from https://msdn.microsoft.com/en-us/library/ms807027.aspx
    
    Dim cnn As ADODB.Connection
    Set cnn = New ADODB.Connection

    cnn.Provider = "MSDASQL"

    'Read that user must have bulkadmin role
    
    cnn.ConnectionString = "driver={SQL Server};" & _
    "server=SERVERNAME;uid=testuser;pwd=password;database=TESTING"
    cnn.Open

    If cnn.State = adStateOpen Then
        MsgBox "Database connection successful!"
    Else
        MsgBox "Database connection error!"
    End If
    
    'Path is on server and not client.
    'If you need it from the client then some additional work will have to be done.
    
    Set rs = cnn.Execute("INSERT INTO Images (ImageAsBinary) SELECT BulkColumn FROM OPENROWSET(BULK N'" & _
                            file & _
                            "', SINGLE_BLOB) image;")
    
    cnn.Close
End Sub
