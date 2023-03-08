Attribute VB_Name = "Module1"
Sub ConvertSelectedMailtoTask()
    Dim objTask As Outlook.TaskItem
    Dim objMail As Outlook.MailItem
     
    Set objTask = Application.CreateItem(olTaskItem)
   
    For Each objMail In Application.ActiveExplorer.Selection
 
With objTask
    .Subject = objMail.Subject
    .StartDate = objMail.ReceivedTime
    .Body = objMail.Body
    ' en tout cas DateSerial(4501, 1, 1) signifie "pas de date" ...
    .StartDate = DateSerial(4501, 1, 1)
    .DueDate = DateSerial(4501, 1, 1)
    .Save
End With
     
    Next
    Set objTask = Nothing
    Set objMail = Nothing
End Sub

Sub MoveToPom1()

    MoveCurrentTaskToFolder ("Pom 1")
    
End Sub

Sub MoveToPom2()

    MoveCurrentTaskToFolder ("Pom 2")
    
End Sub

Sub MoveToPom4()

    MoveCurrentTaskToFolder ("Pom 4")
    
End Sub

Sub MoveToPom8()

    MoveCurrentTaskToFolder ("Pom 8")
    
End Sub

Sub MoveToPomQ()

    MoveCurrentTaskToFolder ("Pom ?")
    
End Sub

Sub MoveToDeps()

    MoveCurrentTaskToFolder ("Deps")
    
End Sub

Sub MoveToOther()

    MoveCurrentTaskToFolder ("Other")
    
End Sub

Sub MoveTozzArchive()

    MoveCurrentTaskToFolder ("zzArchive")
    
End Sub

Sub MoveToInBox()

    Dim objTask As Outlook.TaskItem
    Dim newFolder As Folder
    
    Set objTask = Application.ActiveExplorer.Selection.Item(1)
    
    Set newFolder = Session.GetDefaultFolder(olFolderTasks)
    objTask.Move newFolder
    
End Sub




Sub MoveCurrentTaskToFolder(newFolderName As String)

    Dim objTask As Outlook.TaskItem
    Dim newFolder As Folder
    
    Set objTask = Application.ActiveExplorer.Selection.Item(1)
    
    Set newFolder = Session.GetDefaultFolder(olFolderTasks).Folders.Item(newFolderName)
    objTask.Move newFolder
    
    
    
End Sub

'Dim WithEvents objPane As NavigationPane
'
'Private Sub Application_Startup()
'    Set objPane = Application.ActiveExplorer.NavigationPane
'
'End Sub
  
  
  
Sub objPane_ModuleSwitch(ByVal CurrentModule As NavigationModule)
  Dim objModule As TasksModule
  Dim objGroup As NavigationGroup
  Dim objNavFolder As NavigationFolder

 If CurrentModule.NavigationModuleType <> olModuleTasks Then
     Set objModule = objPane.Modules.GetNavigationModule(olModuleTasks)
     Set objGroup = objModule.NavigationGroups("My Tasks")

' Change the 2 to start in a different folder
     Set objNavFolder = objGroup.NavigationFolders.Item(2)
     objNavFolder.IsSelected = True
  End If
  Set objNavFolder = Nothing
  Set objGroup = Nothing
  Set objModule = Nothing
 End Sub
                                                
 Sub delsig()
    ' delete imposed signature which is in the last 5 lines of mail
    SendKeys "^{END}+{UP}+{UP}+{UP}+{UP}+{UP}{DEL}", True
End Sub

