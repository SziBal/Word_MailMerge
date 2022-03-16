Attribute VB_Name = "NewMacros1"
Sub Levelezes()
Dim MainDoc As Document, TargetDoc As Document
Dim dbPath As String
Dim recordNumber As Long, totalRecord As Long
Dim strFileName As String
Dim strFileExists As String
Dim nameW As String
Dim nameE As String
Dim SOURCE_FILE_PATH As String
Dim FOLDER_SAVED As String

Set MainDoc = ActiveDocument

nameW = ActiveDocument.FullName
SOURCE_FILE_PATH = Replace(nameW, "docx", "xlsx")
FOLDER_SAVED = Left(nameW, Len(nameW) - 5) & "/"

If Dir(FOLDER_SAVED) = "" Then MkDir FOLDER_SAVED

 

If Dir(SOURCE_FILE_PATH) <> "" Then
        With MainDoc.MailMerge
            
                .OpenDataSource name:=SOURCE_FILE_PATH, sqlstatement:="SELECT * FROM [Sheet1$]"
                
                
                totalRecord = .DataSource.RecordCount
        
                For recordNumber = 1 To totalRecord
                    With .DataSource
                        .ActiveRecord = recordNumber
                        .FirstRecord = recordNumber
                        .LastRecord = recordNumber
                    End With
                    
                    strFileName = FOLDER_SAVED & .DataSource.DataFields("column1").Value & "_" & Replace(.DataSource.DataFields("column2").Value, "/", "-") & ".docx"
                    strFileExists = Dir(strFileName)
        
                    If strFileExists = "" Then
                        .Destination = wdSendToNewDocument
                        .Execute False
                        
                        Set TargetDoc = ActiveDocument
                        TargetDoc.SaveAs2 strFileName, wdFormatDocumentDefault
                        TargetDoc.Close False
                        Set TargetDoc = Nothing
                    End If
                            
                Next recordNumber
        
        End With
    Else
        MsgBox "Nincs meg a megfelelõ excel file" & SOURCE_FILE_PATH
    End If
Set MainDoc = Nothing
End Sub
