Option Compare Database
Option Explicit

Private Sub BtnExport2MySQL_Click()
Dim curDB As Database
Set curDB = CurrentDb

Dim tblName As String
tblName = "Task_T"

AppendStringToFile ("DROP TABLE IF EXISTS `" & tblName & "`;" & vbCrLf & vbCrLf)
AppendStringToFile ("CREATE TABLE " & tdf.Name & "(" & vbCrLf)

Dim tbl As TableDef
Set tbl = curDB.TableDefs(tblName)

Dim fld As Field

For Each fld In tdf.Fields
    AppendStringToFile (ComposeFieldStirng(fld))
Next fld


End Sub
Sub AppendStringToFile(strLine As String)
    Dim fPath As String
    fPath = "C:\OtherMy\MyCodes\VBA\MDB2SQL\backup.sql"
    Open strFile_Path For Append As #1
    Write #1, strLine
    Close #1
End Sub
Function CreateDictionary() As Object
Dim d As Object
Set d = CreateObject("Scripting.Dictionary")
d.Add dbDate, "DATETIME"
d.Add dbTime, "DATETIME"
d.Add dbTimeStamp, "DATETIME"
d.Add dbMemo, "LONGTEXT"
d.Add dbByte, "TINYINT(3) UNSIGNED"
d.Add dbInteger, "INTEGER"
d.Add dbLong, "INTEGER"
d.Add dbNumeric, "DECIMAL(16,4)"
d.Add dbDecimal, "DECIMAL(16,4)"
d.Add dbSingle, "FLOAT"
d.Add dbFloat, "FLOAT"
d.Add dbSingle, "FLOAT"
d.Add dbDouble, "DOUBLE"
d.Add dbGUID, "GUID"
d.Add dbBoolean, "TINYINT(1)"
d.Add dbCurrency, "DECIMAL(19,4)"
d.Add dbText, "VARCHAR(fld.Size)"
CreateDictionary = d
End Function