Attribute VB_Name = "CreateDB"
' ==============================================================
' Module:       CreateDB
' Purpose:      Create Database
' ==============================================================
' qbd DATABASE CODE CREATOR
' ==============================================================
' WHAT TO DO NEXT:
' 1.  Add reference to Microsoft DA0 3.6 Library
' 2.  Check the Database_Create() function for Optional Changes
' 3.  To create a database use:
'     bOkay = Database_Create sFilename
'     Where sFilename is the Path and Name of the Database
'     and bOkay is a boolean return value.  If return is false
'     then the creation routine was unsuccessful.
' ==============================================================

Private dbData As Database

Public Function Database_Create(ByVal sFilename As String) As Boolean

' Code created by the qbd Database Code Creator
' Use Find '#' to check optional settings

On Error GoTo Database_Create_Error

Dim dtTable As TableDef

' Create the Database
' # Add password: insert '& ";pwd=NewPassword" after dbLangGeneral
' # Encrypt: insert '+ dbEncrypt' after dbVersion30
Set dbData = DBEngine.CreateDatabase(sFilename, dbLangGeneral, dbVersion40)

' Create table:'Pages'
Set dtTable = dbData.CreateTableDef("Pages")


' Create Indexes for table: Pages
Index_Create dtTable, "PageID", "PageID", dbLong
Index_Create dtTable, "PrimaryKey", "PageID", dbLong, , , True, True
' Create fields
Field_Create dtTable, "PageID", dbLong, , dbAutoIncrField + dbFixedField
Field_Create dtTable, "PageTitle", dbText, 100, dbVariableField
Field_Create dtTable, "PageFilename", dbText, 50, dbVariableField
Field_Create dtTable, "PageHTML", dbMemo, , dbVariableField
dbData.TableDefs.Append dtTable

' Create table:'Template'
Set dtTable = dbData.CreateTableDef("Template")


' Create fields
Field_Create dtTable, "htmlBGColor", dbText, 7, dbVariableField
Field_Create dtTable, "htmlText", dbText, 7, dbVariableField
Field_Create dtTable, "htmlLink", dbText, 7, dbVariableField
Field_Create dtTable, "htmlAlink", dbText, 7, dbVariableField
Field_Create dtTable, "htmlVlink", dbText, 7, dbVariableField
Field_Create dtTable, "htmlLeftMargin", dbInteger, , dbFixedField, , 0
Field_Create dtTable, "htmlTopMargin", dbInteger, , dbFixedField, , 0
Field_Create dtTable, "htmlMetaTags", dbMemo, , dbVariableField
Field_Create dtTable, "htmlScript", dbMemo, , dbVariableField
Field_Create dtTable, "htmlHeaderEnd", dbMemo, , dbVariableField
Field_Create dtTable, "htmlBodyHeader", dbMemo, , dbVariableField
Field_Create dtTable, "htmlFooter", dbMemo, , dbVariableField
Field_Create dtTable, "Notes", dbMemo, , dbVariableField
dbData.TableDefs.Append dtTable


Set dtTable = Nothing
Set dbData = Nothing

' Creation Successful
Database_Create = True
Exit Function

' Whoops an error occured
Database_Create_Error:
' #Add code to trap for errors
Database_Create = False
End Function

Private Sub Field_Create(dtTable As TableDef, _
                         Name As String, _
                         FieldType As Integer, _
                         Optional Size As Integer = 0, _
                         Optional Attributes As Long = 0, _
                         Optional Required As Boolean = False, _
                         Optional DefaultValue As String = "")
Dim dfField As Field

On Error GoTo Field_Create_Err

' Create Field in Table: dtTable

If FieldType = dbText Then
  Set dfField = dtTable.CreateField(Name, FieldType, Size)
Else
  Set dfField = dtTable.CreateField(Name, FieldType)
End If

dfField.Attributes = Attributes
dfField.Required = Required
dfField.DefaultValue = DefaultValue

dtTable.Fields.Append dfField

Set dfField = Nothing
Exit Sub
Field_Create_Err:
' Whoops an error occured
' #Add code to trap for errors
Set dfField = Nothing
End Sub
Private Sub Index_Create(dtTable As TableDef, _
                         Name As String, _
                         FieldName As String, _
                         FieldType As DataTypeEnum, _
                         Optional Size As Integer = 0, _
                         Optional Sort As Boolean = False, _
                         Optional Primary As Boolean = False, _
                         Optional Unique As Boolean = False)

On Error GoTo Index_Create_Err

Dim diIndex As Index
Dim dfField As Field

Set diIndex = dtTable.CreateIndex(Name)
Set dfField = diIndex.CreateField(FieldName, FieldType)

If FieldType = dbText Then
dfField.Size = Size
End If

If Sort Then
dfField.Attributes = dbDescending
End If

With diIndex
  .Fields.Append dfField
  .Primary = Primary
  .Unique = Unique
End With

dtTable.Indexes.Append diIndex

Set diIndex = Nothing
Set dfField = Nothing
Exit Sub

Index_Create_Err:
' Whoops an error occured
' #Add code to trap for errors
Set diIndex = Nothing
Set dfField = Nothing

End Sub
Private Sub Relation_Create(Name As String, _
                            Table As String, _
                            ForeignTable As String, _
                            Field As String, _
                            ForeignField As String, _
                            Optional Attributes As Long = 0)

On Error GoTo Relation_Create_Err

Dim drRelation As Relation
Dim dfField As Field
Set drRelation = dbData.CreateRelation(Name, Table, ForeignTable, Attributes)
drRelation.Fields.Append drRelation.CreateField(Field)
drRelation.Fields(Field).ForeignName = ForeignField
dbData.Relations.Append drRelation

Set dfField = Nothing
Set drRelation = Nothing

Exit Sub
Relation_Create_Err:
' Whoops an error occured
' #Add code to trap for errors
Set dfField = Nothing
Set drRelation = Nothing

End Sub


