VERSION 5.00
Begin VB.Form CompareCopyMethods 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create A Distinct Physical Copy Of A Recordset Object"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "CompareCopyMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function Clone(ByVal rstSource As ADODB.Recordset) As ADODB.Recordset
    'Create a copy of a Recordset using ADO's Recordset.Clone method
        
    Dim rstCopy As ADODB.Recordset
    
    
    Set rstCopy = rstSource.Clone
    
    Set Clone = rstCopy

End Function

Function Copy(ByVal rstSource As ADODB.Recordset) As ADODB.Recordset
    'Create a copy of a Recordset using AD0 2.5 Stream Object and XML
    
    Dim rstCopy As ADODB.Recordset
    Dim objStream As ADODB.Stream
    
    
    'Create a New ADO 2.5 Stream object
    Set objStream = New ADODB.Stream
    
    
    'Save the Recordset to the Stream object in XML format
    rstSource.Save objStream, adPersistXML
    
    
    'Create an exact copy of the saved Recordset from the Stream Object
    Set rstCopy = New ADODB.Recordset
    
    
    rstCopy.Open objStream
    
    
    'Close and de-reference the Stream object
    objStream.Close
    Set objStream = Nothing
    
    Set Copy = rstCopy
    
    
End Function


Sub CreateDistinctRecordsetCopy()
    'Compare ADO's Recordset.Clone and ADO 2.5 Stream object's Recordset copying  methods
    
    Dim rstOne As ADODB.Recordset
    Dim rstTwo As ADODB.Recordset
    Dim rstClone As ADODB.Recordset
    Dim rstCopy As ADODB.Recordset
    
    
    'Create two Fabricated Recordsets
    Set rstOne = CreateFabricatedRecordset
    
    Set rstTwo = CreateFabricatedRecordset
        
    
    'Create a cloned copy of the Recordset using ADO's Recordset.Clone method
    Set rstClone = Clone(rstOne)
    
    'Create a copy of the Recordset using the ADO 2.5 Stream object
    Set rstCopy = Copy(rstTwo)
    
    
    'Delete a record from both Recordset copies
   rstClone.Delete
   rstCopy.Delete
   
    
    'If the cloned Recordset and it's original contain the same number of records then ADO Recordset.Clone copies
    'and their original recordsets point to the same data structures and are not completely distinct copies.
    
    If (rstOne.RecordCount = rstClone.RecordCount) Then
    
        MsgBox "Recordset.Clone Copies Are Not Completely Separate Objects From Their Original Recordsets"
        
    Else
    
        MsgBox "Recordset.Clone Copies Are Completely Separate Objects From Their Original Recordsets"
        
    End If


    'If the Recordset copied using ADO 2.5 Stream object and it's original contain differing number of records then
    'ADO 2.5 Stream object Recordset copies are  completely distinct copies of the original recordsets.
    
    If (rstTwo.RecordCount = rstCopy.RecordCount) Then
    
        MsgBox "ADO 2.5 Stream Recordset Copies Are Not Completely Separate Objects From Their Original Recordsets"
        
    Else
    
        MsgBox "ADO 2.5 Stream Recordset Copies Are Completely Separate Objects From Their Original Recordsets"
        
    End If
    
    
End Sub

Private Sub cmdCreate_Click()
    Call Me.CreateDistinctRecordsetCopy
End Sub

Function CreateFabricatedRecordset() As ADODB.Recordset
    'Creates a Fabricated Recordser populated with People's names
    
    Dim rst As ADODB.Recordset
    Dim varField As Variant
    
    
    'Create a Fabricated Recordset
    Set rst = New ADODB.Recordset
        
    With rst.Fields
            .Append "LastName", adVarChar, 20
    End With
        
    'Open the recordset
    rst.Open
    
    
    varField = Array("LastName")
    
    
    'Populate the fabricated Recordset with 3 names
    With rst
        .AddNew varField, Array("John")
        .AddNew varField, Array("Paul")
        .AddNew varField, Array("King")
    End With
    
    
    'Move to the first record
    rst.MoveFirst
    
    
    'Return the created Recordset
    Set CreateFabricatedRecordset = rst
    

End Function
