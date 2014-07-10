Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_Description = "Macro Macro1\nMacro enregistrée le Lun 07/07/14 par install."
' Macro Macro1
' Macro enregistrée le Lun 07/07/14 par install.
    EditCopy
    SelectTaskField Row:=97, Column:="Nom"
    SetTaskField Field:="Nom", Value:="Cadrage"
    SelectTaskField Row:=1, Column:="Nom"
    SetTaskField Field:="Nom", Value:="Réalisation"
    SelectTaskField Row:=1, Column:="Nom"
    SelectTaskField Row:=-2, Column:="Remarques"
    EditPaste
    SelectTaskField Row:=1, Column:="Remarques"
    EditPaste
End Sub

Sub OrgaSelection()

    
    For iTask = ActiveSelection.Tasks.Count To 1 Step -1
        Dim t As Task
        
        Set t = ActiveSelection.Tasks.Item(iTask)
        
        AddRealString t
        InsConception t
    
    Next

End Sub

Sub AddRealString(t As Task)
' Add " - réalisation" to task

      t.Name = t.Name & " - réalisation"

End Sub

Sub InsConception(t As Task)
' Insert conception task before t

    
    Dim newTask As Task
    Dim newName As String
    
    Dim oldPredecessors As String
    
    
    oldPredecessors = t.Predecessors
    
    newName = replString(t.Name, "- réalisation", "- conception")
    
    Set newTask = ActiveSelection.Tasks.Add(newName, t.ID)
    
    newTask.Predecessors = ""
    t.Predecessors = oldPredecessors

End Sub

Function replString(inString As String, oldString As String, newString As String) As String
    Dim retString As String
    
    If InStr(inString, oldString) > 0 Then
        retString = Left(inString, InStr(inString, oldString) - 1)
    End If
    
    replString = retString & newString
    
End Function



Sub Macro2()
' Macro Macro2
' Macro enregistrée le Lun 07/07/14 par install.
    RowInsert
    SelectTaskField Row:=0, Column:="Nom"
    SetTaskField Field:="Nom", Value:="tretr"
    SelectTaskField Row:=2, Column:="Nom"
End Sub
