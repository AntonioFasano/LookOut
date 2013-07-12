VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LookOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


' (shell-command (substitute-in-file-name "$home/Desktop/outlook/git/source/build.cmd"))

 
'/**
' A library to free from Outlook Contacts tax.
'
' @author Antonio Fasano
' @version 1.0 first documented version
'**/

Option Explicit
Private EmptyContacts

'Structure for duplicate analysis
Private Type ItemDB
    ContactID As String
    FolderId As String
    FolderPath As String
    LastNameAndFirstName As String
    FileAs As String
    Duplicate As Boolean
End Type

'/**
' Parse a file system path string, expandng env variables, and return the expanded path or an empty string.
' \param  sPath is a file system path sting, possibly with environment variables
' \param  PathExist if true  `sPath` must exist.
' @return [As String] The file system object representing the passed path string.
' If the path is detected as not valid or the path, being `PathExist` true, doesn't exist, then return an empty string.
'
' In Tools->References:
'
'  "Microsoft Scripting Runtime", "Windows Script Host Object Model"
'
'  need to be checked
'**/
Private Function GetFSPath(sPath As String, Optional PathExist As Variant) As String

   'In Tools->References check "Microsoft Scripting Runtime"
    Dim fso As New Scripting.FileSystemObject
    
    'In Tools->References check "Windows Script Host Object Model"
    Dim oWsh As New WshShell
    
    
    'Check folder paths
    sPath = oWsh.ExpandEnvironmentStrings(sPath)
    Dim sParent  As String: sParent = fso.GetParentFolderName(sPath)
    Dim bPathExist As Boolean: bPathExist = CBool(PathExist)
    GetFSPath = ""
        
    If Not fso.FolderExists(sParent) Then
        MsgBox (sParent & vbCrLf & "does not exist.")
        GoTo ExitFunction:
    End If
    
    If bPathExist = True And Not (fso.FileExists(sPath) Or fso.FolderExists(sPath)) Then
        MsgBox (sPath & vbCrLf & "does not exist.")
        GoTo ExitFunction:
    End If

    
    If bPathExist = False And Not IsMissing(PathExist) And (fso.FileExists(sPath) Or fso.FolderExists(sPath)) Then
        MsgBox (sPath & vbCrLf & "does exist.")
        GoTo ExitFunction:
    End If
   
    
    sPath = fso.GetAbsolutePathName(sPath)
    If fso.FolderExists(sPath) Then sPath = sPath & "\"
    
    GetFSPath = sPath
    
    
ExitFunction:
    Set fso = Nothing
    Set oWsh = Nothing

End Function

'/**
' Returns the path of the global default contact folder.
' @return [As String] The path of the global contact folder.
'
' As many commands accept as input the global default contact folder,
' you may want to learn  its path with respect to Outlook (tree) navigation pane.
' Note: The name of the store contained in Outlook paths is orignally like  its related  PST file,
' but the user can change it. With RClick on the store node,
' check `Data File Properties->Advanced`to learn about the PST name.
'**/
Public Property Get GlobalDefaultPath() As String
    GlobalDefaultPath = GetFolder(DefaultFolder:=True).FolderPath
End Property

'/**
' Returns the path of default contact folder for a given store.
' \param  StoreName is a string containing the name of store preceeded by `\\`, possibly like the realted PST file.
' @return [As String] The path of default contact folder for a given store.
'
' As many commands accept as input the store default contact folder,
' you may want to learn  its path with respect to Outlook (tree) navigation pane.
' Note: The name of the store contained in Outlook paths is orignally like  its related  PST file,
' but the user can change it. With RClick on the store node,
' check `Data File Properties->Advanced`to learn about the PST name.
'**/
Function StoreDefaultPath(StoreName As String) As String

    Dim s As String
    s = Replace(StoreName, "\\", "")
    If UBound(Split(s, "\")) > 0 Then
        MsgBox ("Use only the store name.")
        Exit Function
    End If

    StoreDefaultPath = GetFolder(StoreName).FolderPath
    
End Function

'/**
' Parse Outlook folder paths and returns the related Outlook folder object
' \param  FolderPath is the path to parse formatted like `\\root\folder\subbolder\etc`; where `root` is the main node in the navigation pane and corrisponds to a single Outlook data file (.pst).
' Note: PST file name can be different from displayed node name
' If the path is simply `\\root`, the default contact folder for the given root will be used
' \param  DefaultFolder If True and with FolderPath omitted cause to return the global default folder object.
' @return [As Outlook.Folder] The Outlook folder object representing a valid string path, `Nothing` otherwise.
'**/
Private Function GetFolder(Optional ByVal FolderPath As String, Optional DefaultFolder As Boolean) As Outlook.Folder

    Dim ppath() As String
    Dim isStore As Boolean
    Dim oStore As Outlook.Store, oFolder As Outlook.Folder, colFolders As Outlook.Folders


    Set GetFolder = Session.GetDefaultFolder(olFolderContacts)
    
    If DefaultFolder And FolderPath <> "" Then
        MsgBox ("You asked both for the default contact folder:" & vbCrLf & _
                GetFolder.FolderPath & vbCrLf & "and for folder: " & vbCrLf & FolderPath _
                & vbCrLf & "Please, choose one alternative.")
        GoTo ErrHandler
    End If
    
    If DefaultFolder Then Exit Function
    
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Mid(FolderPath, 3, Len(FolderPath))
    Else
        MsgBox ("Contact folders paths have to start with '\\'")
        GoTo ErrHandler
    End If
    
    If Right(FolderPath, 1) = "\" Then
        FolderPath = Mid(FolderPath, 1, Len(FolderPath) - 1)
    End If
    
    
        
    If FolderPath = "" Then
        MsgBox ("Wrong path")
        GoTo ErrHandler
    End If


    ppath = Split(FolderPath, "\")

    isStore = False
    For Each oStore In Application.Session.Stores
    
        If LCase(oStore.DisplayName) = LCase(ppath(0)) Then
            isStore = True
            Exit For
        End If
          
    Next
    If Not isStore Then
        MsgBox (ppath(0) & vbCrLf & "does not exist.")
        GoTo ErrHandler
    End If


    If UBound(ppath) = 0 Then
        Set GetFolder = oStore.GetDefaultFolder(olFolderContacts)
        Exit Function
    End If

    
    Dim i As Integer
    On Error Resume Next
    
    Set oFolder = oStore.GetRootFolder
    For i = 1 To UBound(ppath)
        Set colFolders = oFolder.Folders
        Set oFolder = Nothing
        Set oFolder = colFolders.Item(ppath(i))
        If oFolder Is Nothing Then
            MsgBox ("\\" & FolderPath & vbCrLf & "does not exist")
            GoTo ErrHandler:
        End If
    Next i
    
    Set GetFolder = oFolder


Exit Function
ErrHandler:
    Set GetFolder = Nothing

End Function

'/**
' Add categories to all contacts in a given folder (and its subfolders),
'  named after the name of the folder and subfolders
'  Previous category names are preserved unless they duplicate new ones. Comparison is not case sensitive.
' \param ParentFolder is the folder to be processed (subfolders included).
' \param DefaultFolder If True, ParentFolder (to be omitted) will be the global default contact folder.
' \param FolderNames is automatically created by recursion and contains subfolders.
' \param GoogleStyle If True, nested folder-categeories are concatenated with a dot, e.g. `cat.subcat`.
'**/
Sub Folder2Cat(Optional ParentFolder As String, Optional DefaultFolder As Boolean, _
               Optional ByVal FolderNames As String, Optional GoogleStyle As Boolean _
               )

    Dim oParentFolder As Outlook.Folder, oFolder As Outlook.Folder
    Dim oContact As Outlook.ContactItem
    Dim folderCategs()  As String, folderCateg As Variant, _
        currentCategs() As String, currentCateg As Variant


    'Check folder paths
    Set oParentFolder = GetFolder(ParentFolder, DefaultFolder)
    If oParentFolder Is Nothing Then Exit Sub
                
      
    If FolderNames = "" Then
        FolderNames = oParentFolder.Name
    Else
        FolderNames = FolderNames & ", " & oParentFolder.Name
    End If
     
    Debug.Print vbCrLf & ">" & FolderNames
    
    
    'Loop through contact items
    Dim sCat As String, x(0) As String 'Category string and dummy array differences
    For Each oContact In oParentFolder.Items

        Debug.Print oContact.FileAs
        
       'Each folder the contact belongs to will become a category of its
       ' Outlook "categories" field is a string like such: "cat1, cat2, etc."
       ' Leading/triling commas ", cat1, cat2," are trimmed out by Outlook
       '----------------------------------------------------------------
        folderCategs = Split(FolderNames, ", ")           ' Array of folder cats to apply to current item
        currentCategs = Split(oContact.Categories, ", ")  ' Array of current cats
        
        'Subtract folderCategs from currentCategs (case insensitive)
        currentCategs = ArrayDiff(currentCategs, folderCategs)
        
        'Subtract default "Contacts" folder from folderCategs (not useful as a category)
        x(0) = oParentFolder.Store.GetDefaultFolder(olFolderContacts)
        folderCategs = ArrayDiff(folderCategs, x)
        
        'Rebuild "Categories" field
        sCat = Join(currentCategs, ", ") & ", " & Join(folderCategs, ", ")
        If GoogleStyle Then sCat = Join(currentCategs, ", ") & ", " & Join(folderCategs, ".")
        oContact.Categories = sCat
        
        oContact.Save
        Debug.Print "-> " & oContact.Categories
        
    Next 'contact
    Set oContact = Nothing


    'Get subfolders by recursion
    If (oParentFolder.Folders.Count > 0) Then
        For Each oFolder In oParentFolder.Folders
            Folder2Cat oFolder.FolderPath, , FolderNames, GoogleStyle
        Next
    End If
    
    Set oParentFolder = Nothing
    Set oFolder = Nothing
    
    
End Sub


'/**
' Add a category to all contacts in a given folder (but not to its subfolders)
'  Previous categories names are preserved unless they duplicate new one.
' \param FolderPath is the path folder whose contacts are to be categorised.
' \param NewCat is the category to add.
'**/
Sub AddCat(FolderPath As String, NewCat As String)

    Dim oFolderPath As Outlook.Folder, oContact As Outlook.ContactItem
    
        
    'Check paths of folders
    Set oFolderPath = GetFolder(FolderPath)
    If oFolderPath Is Nothing Then GoTo ExitSub:

    Dim sCat As String
    For Each oContact In oFolderPath.Items
    
        Debug.Print oContact.FileAs, oContact.Categories
        sCat = Replace(oContact.Categories, NewCat, "", , , vbTextCompare)
        oContact.Categories = sCat & ", " & NewCat
        oContact.Save
        Debug.Print oContact.FileAs, oContact.Categories

    Next
    
ExitSub:
    Set oFolderPath = Nothing
    Set oContact = Nothing
    
End Sub


'/**
' Remove a category to all contacts in a given folder (but not to its subfolders)
' \param FolderPath  is the folder whose contacts will be subject to category removal.
' \param Cat2Del is the category to remove.
'**/
Sub DelCat(FolderPath As String, Cat2Del As String)

    Dim oFolderPath As Outlook.Folder, oContact As Outlook.ContactItem
    
        
    'Check paths of folders
    Set oFolderPath = GetFolder(FolderPath)
    If oFolderPath Is Nothing Then GoTo ExitSub:

    Dim sCat As String
    For Each oContact In oFolderPath.Items
    
        Debug.Print oContact.FileAs, oContact.Categories
        sCat = Replace(oContact.Categories, Cat2Del, "", , , vbTextCompare)
        oContact.Categories = sCat
        oContact.Save
        Debug.Print oContact.FileAs, oContact.Categories

    Next
    
ExitSub:
    Set oFolderPath = Nothing
    Set oContact = Nothing
    
End Sub



'/**
' Remove all categories from a given folder and its subfolders

' \param ParentFolder  is the folder to be processed (subfolders included).
' \param DefaultFolder If True, ParentFolder (to be omitted) will be the global default contact folder.
'**/
Sub DelAllCats(Optional ParentFolder As String, Optional DefaultFolder As Boolean _
               )
    Dim oParentFolder As Outlook.Folder, oFolder As Outlook.Folder
    Dim oContact As Outlook.ContactItem

    'Check folder paths
    Set oParentFolder = GetFolder(ParentFolder, DefaultFolder)
    If oParentFolder Is Nothing Then Exit Sub
      
    'Take note of current folder for the user
    Debug.Print ">" & oParentFolder.FolderPath
          
    
    'Loop through contact items
    For Each oContact In oParentFolder.Items
        
        Debug.Print oContact.FileAs & "  [" & oContact.Categories & "]"
        oContact.Categories = ""
        oContact.Save

    Next
    Set oContact = Nothing


    'Get subfolders by recursion
    If (oParentFolder.Folders.Count > 0) Then
        For Each oFolder In oParentFolder.Folders
            DelAllCats oFolder.FolderPath
        Next
    End If
    
    Set oParentFolder = Nothing
    Set oFolder = Nothing
    
End Sub



'/**
' Fills empty contact names (of a given folder and subfolders) with file-as field.
' \param ParentFolder is the folder to be processed (subfolders included).
' \param DefaultFolder If True, ParentFolder (to be omitted) will be the global default contact folder.
'**/
Sub FillEmpty(Optional ParentFolder As String, Optional DefaultFolder As Boolean _
               )

    Dim oParentFolder As Outlook.Folder, oFolder As Outlook.Folder
    Dim oContact As Outlook.ContactItem
        
    'Check folder paths
    Set oParentFolder = GetFolder(ParentFolder, DefaultFolder)
    If oParentFolder Is Nothing Then Exit Sub
          
    'Take note of current folder for the user
    Debug.Print ">" & oParentFolder.FolderPath
     
    
    'Loop through contact items
    For Each oContact In oParentFolder.Items
    
        If oContact.FirstName = "" And oContact.LastName = "" Then
    
            Debug.Print " " & oContact.FileAs
            oContact.LastName = oContact.FileAs
            oContact.Save
                
        End If
               
    Next
    Set oContact = Nothing


    'Get subfolders by recursion
    If (oParentFolder.Folders.Count > 0) Then
        For Each oFolder In oParentFolder.Folders
            FillEmpty oFolder.FolderPath
        Next
    End If
         
    Set oParentFolder = Nothing
    Set oFolder = Nothing

End Sub

'/**
' Slave for ShowEmpty. See ShowEmpty synopsis for more.
' \param ParentFolder is the folder to be processed (subfolders included).
' \param DefaultFolder If True, ParentFolder (to be omitted) will be the global default contact folder.
'**/
Private Sub ShowEmpty_slave(Optional ParentFolder As String, Optional DefaultFolder As Boolean)

    Dim oParentFolder As Outlook.Folder, oFolder As Outlook.Folder
    Dim oContact As Outlook.ContactItem
    
        
    'Check folder paths
    Set oParentFolder = GetFolder(ParentFolder, DefaultFolder)
    If oParentFolder Is Nothing Then Exit Sub
          
    'Take note of current folder for the user
    Debug.Print ".";
     
    
    'Loop through contact items
    For Each oContact In oParentFolder.Items
    
        If oContact.FirstName = "" And oContact.LastName = "" Then
            EmptyContacts = True
            Debug.Print " "
            Debug.Print " " & oContact.FileAs & " in " & oParentFolder.FolderPath
        End If
               
    Next
    Set oContact = Nothing

    'Get subfolders by recursion
    If (oParentFolder.Folders.Count > 0) Then
        For Each oFolder In oParentFolder.Folders
            ShowEmpty_slave oFolder.FolderPath
        Next
    End If
         
    Set oParentFolder = Nothing
    Set oFolder = Nothing
 

End Sub

'/**
' Show empty contact names possibly present in a given folder and its subfolders.
' \param ParentFolder is the folder to be processed (subfolders included).
' \param DefaultFolder If True, ParentFolder (to be omitted) will be the global default contact folder.
'**/
Sub ShowEmpty(Optional ParentFolder As String, Optional DefaultFolder As Boolean)
    
EmptyContacts = False
If Not DefaultFolder Then
    ShowEmpty_slave ParentFolder
Else
    ShowEmpty_slave DefaultFolder:=True
End If
    
If Not EmptyContacts Then Debug.Print Chr(13) & "No empty name contact found"


End Sub
'/**
' Copy all contacts in TreeParent and its subfoders to the single folder DestFolder,
' without keeping the orignal folder stucutre.
' \param TreeParent Outlook folder to flatten, ie to copy without folder structure.
' \param DestFolder destination folder for flattened contacts. Should exist.
' \param Recursed is used internally to distinguish the first run in recursion cycles.
'**/
Sub FlattenFolderCopy(TreeParent As String, DestFolder As String, _
                        Optional Recursed As Boolean _
               )
                   
    Dim oTreeParent  As Outlook.Folder, oDestFolder As Outlook.Folder, _
        oFolder As Outlook.Folder, collFolder As Outlook.Items, _
        oContact2 As Outlook.ContactItem, oContact As Outlook.ContactItem


    Set oTreeParent = GetFolder(TreeParent)
    Set oDestFolder = GetFolder(DestFolder)
    If oTreeParent Is Nothing Or oDestFolder Is Nothing Then GoTo ErrHandler:
       
       
    'If DestFolder is already dirty on first run, exit
    If Not Recursed And oDestFolder.Items.Count > 0 Then
           ' MsgBox ("Destination folder:" & vbCrLf & oDestFolder.FolderPath & vbCrLf & "not empty")
            'GoTo ErrHandler:
    End If
       
        
    'Take note of current folder for the user
    Debug.Print ">" & oTreeParent.FolderPath
    

    'Loop through contact items and move to destination folder
    Dim i As Integer, n As Integer
    n = oTreeParent.Items.Count
    Set collFolder = oTreeParent.Items
    'For Each oContact In oTreeParent.Items
    For i = 1 To n
        Set oContact = collFolder(i)
        Set oContact2 = oContact.Copy
        'Debug.Print i & " to copy: " & oContact.FileAs & " " & oContact2.FileAs
        If oContact2.Move(oDestFolder) Is Nothing Then
            MsgBox ("Copy error with: " & oContact2.FileAs & vbCrLf & "in folder:" & vbCrLf & oTreeParent)
            GoTo ErrHandler:
        End If
    Next i
    Debug.Print i - 1 & " contacts copied"
    Set oContact = Nothing
    Set oContact2 = Nothing
    Set collFolder = Nothing

            
    'Get subfolders by recursion
    If (oTreeParent.Folders.Count > 0) Then
        For Each oFolder In oTreeParent.Folders
            FlattenFolderCopy oFolder.FolderPath, DestFolder, Recursed:=True
        Next
    End If
    
ErrHandler:
    Set oFolder = Nothing
    Set oDestFolder = Nothing
    Set oTreeParent = Nothing
    
End Sub

'/**
' Copy all contacts in TreeParent and its subfoders to the single folder DestFolder,
' without keeping the orignal folder stucutre.
' \param TreeParent Outlook folder to flatten, ie to move from without folder structure.
' \param DestFolder destination folder for flattened contacts. Should exist.
' \param Recursed is used internally to distinguish the first run in recursion cycles.
'**/
Sub FlattenFolderMove(SourceFolder As String, DestFolder As String, _
                        Optional Recursed As Boolean _
               )
              
       
    Dim oSourceFolder  As Outlook.Folder, oDestFolder As Outlook.Folder, _
        oFolder As Outlook.Folder
               
    'Check folder paths
    If DestFolder = "" Then
        MsgBox ("Destination folder not given")
        GoTo ErrHandler:
    End If
    Set oSourceFolder = GetFolder(SourceFolder)
    Set oDestFolder = GetFolder(DestFolder)
    If oSourceFolder Is Nothing Or oDestFolder Is Nothing Then GoTo ErrHandler:
       
       
    'If DestFolder is already dirty on first run, exit
    If Not Recursed And oDestFolder.Items.Count > 0 Then
            MsgBox ("Destination folder:" & vbCrLf & oDestFolder.FolderPath & vbCrLf & "not empty")
            GoTo ErrHandler:
    End If
       
       
              
    'Take note of current folder for the user
    Debug.Print ">" & oSourceFolder.FolderPath
                
    
    'Loop through contact items and move them to destination folder
            
    'As we move contact their item number changes
    'With 2 items, when you ask to move the 2nd, it is actually the first (and you run out bound)
    'The reverse strategy can prevent this
    Dim i As Integer, n As Integer
    n = oSourceFolder.Items.Count
    For i = 0 To n - 1
    
        Debug.Print oSourceFolder.Items(n - i).FileAs
        oSourceFolder.Items(n - i).Move oDestFolder
        
    Next i



    'Get subfolders by recursion
    If (oSourceFolder.Folders.Count > 0) Then
        For Each oFolder In oSourceFolder.Folders
            FlattenFolderMove oFolder.FolderPath, DestFolder, Recursed:=True
        Next
    End If
    
    
    
ErrHandler:
    Set oFolder = Nothing
    Set oDestFolder = Nothing
    Set oSourceFolder = Nothing
End Sub


'/**
' Find "weak" duplcates in a whole folder tree based on Last name-First name and Fileas.
' \param ParentFolder is the folder to be processed (subfolders included).
' \param DefaultFolder If True, ParentFolder (to be omitted) will be the global default contact folder.
'
' "Weak duplicates" share only name fields, but can have different values for other fields
'**/
Sub FindWDup(Optional ParentFolder As String, Optional DefaultFolder As Boolean)

    Dim oParentFolder As Outlook.Folder, oFolder As Outlook.Folder, _
        oContact As Outlook.ContactItem
   
   
    Dim oContDB() As ItemDB
    Dim nContact As Integer
    Dim nDupes As Integer: nDupes = 0
    Dim DupeIndex() As String


    'Check paths of folders
    Set oParentFolder = GetFolder(ParentFolder, DefaultFolder)
    If oParentFolder Is Nothing Then GoTo ExitSub:


    'Fill contact DB
    FindWDup_Slave oParentFolder, oContDB

    'Find duplicates in oContDB
    Debug.Print "Looking for duplicates... "

    Dim i As Integer, j As Integer: nContact = UBound(oContDB)
    For i = 1 To nContact
        For j = i + 1 To nContact
            If oContDB(i).LastNameAndFirstName = oContDB(j).LastNameAndFirstName _
                Or oContDB(i).FileAs = oContDB(j).FileAs Then
            
                If Not oContDB(i).Duplicate Then
                    nDupes = nDupes + 1
                    ReDim Preserve DupeIndex(1 To nDupes)
                    DupeIndex(nDupes) = oContDB(i).ContactID
                End If
                If Not oContDB(j).Duplicate Then
                    nDupes = nDupes + 1
                    ReDim Preserve DupeIndex(1 To nDupes)
                    DupeIndex(nDupes) = oContDB(j).ContactID
                End If
                
            End If
        Next j
    Next i
    
    
    'Print duplicates in DupeIndex, if any
    If nDupes = 0 Then
        Debug.Print "No duplicates found"
    Else
    
        For i = 1 To UBound(DupeIndex)
            Set oContact = Session.GetItemFromID(DupeIndex(i))
            Set oFolder = oContact.Parent
            Debug.Print oFolder.FolderPath, oContact.FileAs
        Next i
        
    End If
        
ExitSub:
    Set oParentFolder = Nothing
    Set oFolder = Nothing
    Set oContact = Nothing

End Sub


'/**
' Slave for FindWDup. See FindWDup for more.
' \param ParentFolder is the folder to be processed (subfolders included).
' \param oContDB is a structure created by recursion with contacts details.
'
' This service function fills oContDB with contacts in oParentFolder and its subfolders,
' for subsequent analysis by the main function FindWDup.
'**/
Private Sub FindWDup_Slave(ByVal oParentFolder As Outlook.Folder, _
               ByRef oContDB() As ItemDB _
               )


    Dim oFolder As Outlook.Folder
    Dim oContact As Outlook.ContactItem
    

    Dim nContact As Integer
    


    'Check if GetFolder() returned nothing
    If oParentFolder Is Nothing Then Exit Sub
    
    Debug.Print "Scanning " & oParentFolder.FolderPath
    
    'Init nContact to 0 if oContDB uninitialized
    If (Not Not oContDB) = 0 Then
        nContact = 0
    Else
        nContact = UBound(oContDB)
    End If
    
    'Loop through contact items and fill the oContDB DB with contacts
    If oParentFolder.Items.Count > 0 Then ReDim Preserve oContDB(1 To nContact + oParentFolder.Items.Count)
    For Each oContact In oParentFolder.Items
        
        nContact = nContact + 1
        With oContDB(nContact)
            .ContactID = oContact.EntryID
            .FileAs = oContact.FileAs
            .LastNameAndFirstName = oContact.LastNameAndFirstName
            .Duplicate = False
            '.FolderId
        End With
    
    Next
    Set oContact = Nothing


    'Get subfolders by recursion
    If (oParentFolder.Folders.Count > 0) Then
        For Each oFolder In oParentFolder.Folders
            FindWDup_Slave oFolder, oContDB
        Next
    End If
    
    
    Set oFolder = Nothing

    
End Sub


'/**
' Find duplicates in a single folder, based on Last name-First name, Fileas and Modified,
' and copy non-duplicate contacts to a new folder.
' \param SourceFolder folder to be scanned for duplicates (subfolders excluded).
' \param DestFolder destination folder for copying non-duplicate contacts.
'
' Given the same names and modification time, contacts are considered equal,
' so only one instance is copied do the destination folder.
'**/
Sub CopyUnique(SourceFolder As String, DestFolder As String)

    Dim oSourceFolder As Outlook.Folder, oDestFolder  As Outlook.Folder, _
        oContact As Outlook.ContactItem, oContact2Check As Outlook.ContactItem, oCopiedContact As ContactItem, _
        unique As Boolean


    'Check paths of folders
    Set oSourceFolder = GetFolder(SourceFolder, DefaultFolder:=False)
    Set oDestFolder = GetFolder(DestFolder, DefaultFolder:=False)
    If oSourceFolder Is Nothing Or oDestFolder Is Nothing Then GoTo ExitSub:


    'Exit if source folder empty
    Dim nContact As Integer: nContact = oSourceFolder.Items.Count
    If nContact = 0 Then GoTo ExitSub:

    'Scan for unique contacts and copy them to destination folder
    Dim i As Integer, j As Integer
    For i = 1 To nContact
        
        unique = True
        Set oContact = oSourceFolder.Items(i)
        
        'Check if for j<i there are some duplicates
        For j = 1 To i - 1
            Set oContact2Check = oSourceFolder.Items(j)
            If oContact2Check.FileAs = oContact.FileAs And _
               oContact2Check.LastNameAndFirstName = oContact.LastNameAndFirstName And _
               oContact2Check.LastModificationTime = oContact.LastModificationTime Then
               unique = False
               Exit For
            End If
        Next j
        If unique Then
            Set oCopiedContact = oContact.Copy
            oCopiedContact.Move oDestFolder
            Debug.Print i & " Copied unique: " & oContact.FileAs
        Else
         Debug.Print i & " Not copied dup: " & oContact.FileAs
        End If
        
     Next i
 

ExitSub:
    Set oSourceFolder = Nothing
    Set oDestFolder = Nothing
    Set oContact = Nothing
    Set oContact2Check = Nothing
    Set oCopiedContact = Nothing

End Sub


'/**
' Export all contact in a folder tree as vCards (.vcf) including Outlook categories.
' \param SavePath is the full path of export dir. The tree structure is not kept.
'  It must exist and can have env variables, e.g. `%%USERPROFILE%\Desktop\Export`.
' \param SourceFolder is the Outlook folder to export (subfolders included).
' \param DefaultFolder If True, SourceFolder (to be omitted) will be the global default contact folder.
'
' The vCard filenames are based on the contact `FileAs` fields.
' If a duplicate contact is found, the procedure stops exporting further contacts
'  and skips to the next folder in the queue, if any.
'
' In Tools->References:
'
'  "Microsoft Scripting Runtime", "Windows Script Host Object Model"
'
'  need to be checked
'
'**/
Sub ExportVcf(SavePath As String, _
                Optional SourceFolder As String, Optional DefaultFolder As Boolean _
           )

    Dim oSourceFolder As Outlook.Folder, oFolder As Outlook.Folder
    Dim oContact As Outlook.ContactItem
        
    
    'In Tools->References check "Microsoft Scripting Runtime"
    Dim fso As New Scripting.FileSystemObject
    Dim fContact  As TextStream
    
    'In Tools->References check "Windows Script Host Object Model"
    Dim oWsh As New WshShell
    
    Dim sCategories As String
    Dim sContact As String, i As Integer
    
    'Check folder paths
    Set oSourceFolder = GetFolder(SourceFolder, DefaultFolder)
    If oSourceFolder Is Nothing Then Exit Sub
        
    SavePath = oWsh.ExpandEnvironmentStrings(SavePath)
    If Not fso.FolderExists(SavePath) Then
        MsgBox (SavePath & vbCrLf & "does not exist.")
        Exit Sub
    End If
    SavePath = fso.GetAbsolutePathName(SavePath)
    SavePath = SavePath & "\"
    
    
    'Take note of current folder for the user
    Debug.Print ">" & oSourceFolder.FolderPath
        
    

    'Loop through contact items and export
    For Each oContact In oSourceFolder.Items

        'Format categories for vcf export
        sCategories = "CATEGORIES;CHARSET=UTF-8:" & Replace(oContact.Categories, ",", "\,")
        Debug.Print oContact.FileAs
        Debug.Print sCategories
                    
                    
        'Export contact no categories and break in case of duplicates
        If fso.FileExists(SavePath & oContact.FileAs & ".vcf") Then
            MsgBox ("Contact: " & vbCrLf & oContact.FileAs & vbCrLf & "already exists." & vbCrLf _
            & "Stop exporting folder:" & vbCrLf & oSourceFolder.FolderPath)
            Exit Sub
        End If
        oContact.SaveAs SavePath & oContact.FileAs & ".vcf", olVCard
                    
        'Parse and add categories to vcf file
        Set fContact = fso.OpenTextFile(SavePath & oContact.FileAs & ".vcf", ForReading, False)
        sContact = ""
        'Read first 2 lines
        sContact = sContact & vbCrLf & fContact.ReadLine _
                            & vbCrLf & fContact.ReadLine
        'Add category line as third line
        sContact = sContact & vbCrLf & sCategories
        'Read remaining
        sContact = sContact & vbCrLf & fContact.ReadAll
        fContact.Close
        'Write back the new vcf string to cvf file
        Set fContact = fso.OpenTextFile(SavePath & oContact.FileAs & ".vcf", ForWriting, False)
        fContact.Write sContact
        fContact.Close
                        
        
    Next
    Set oContact = Nothing
    Set fContact = Nothing
    Set fso = Nothing
    Set oWsh = Nothing

    
    'Get subfolders by recursion
    If (oSourceFolder.Folders.Count > 0) Then
        For Each oFolder In oSourceFolder.Folders
            ExportVcf SavePath, oFolder.FolderPath
        Next
    End If
    
    Set oSourceFolder = Nothing
    Set oFolder = Nothing


End Sub

'/**
' Combine single-contact vCards, in a given directory, into a single multi-contact vCard.
' \param SourceFolder is the file system folder containing vCards to export (subfolders excluded).
'  It must exist and can have env variables, e.g. `%USERPROFILE%\Desktop\SourceVcf`
' \param MultivCard is the full path of the multi-contact vCard.
'  It can have env variables, e.g. `%%USERPROFILE%\Desktop\Multi.vcf`
'
' Before combining, source vCards are checked to be single-contact vCards.
' Non vCard files in `SourceFolder` are ignored.
'
' In Tools->References:
'
'  "Microsoft Scripting Runtime", "Windows Script Host Object Model"
'
'  need to be checked
'
'**/
Sub MultiVcf(SourceFolder As String, MultivCard As String)


    Dim oContact As Outlook.ContactItem
            
    'In Tools->References check "Microsoft Scripting Runtime"
    Dim fso As New Scripting.FileSystemObject
    Dim oSourceFolder As Scripting.Folder, oSourceFile As Scripting.File
    Dim fContact As Scripting.TextStream, oMultivCard As Scripting.TextStream
        
        
    'In Tools->References check "Windows Script Host Object Model"
    Dim oWsh As New WshShell
        
    'Check paths
    SourceFolder = GetFSPath(SourceFolder, PathExist:=True)
    MultivCard = GetFSPath(MultivCard, PathExist:=False)
    If SourceFolder = "" Or MultivCard = "" Then GoTo ExitSub:
    
    If fso.FileExists(SourceFolder) Then
        MsgBox (SourceFolder & vbCrLf & "is a file, not a folder.")
        GoTo ExitSub:
    End If
    
    Set oSourceFolder = fso.GetFolder(SourceFolder)
  
                        
    'Parse vcf files for multi-contact
    Debug.Print vbCrLf & "Check if any vCard is already a multi-contact."
    Dim line As String, n As Integer, finePrint As Integer: finePrint = 0
    For Each oSourceFile In oSourceFolder.Files
        finePrint = finePrint + 1
        If finePrint / 50 = Int(finePrint / 50) Then Debug.Print
        Debug.Print ".";
        If UCase(fso.GetExtensionName(oSourceFile)) = UCase("vcf") Then
            Set fContact = fso.OpenTextFile(oSourceFile.Path, ForReading, False)
            n = 0
            Do Until fContact.AtEndOfStream Or n = 2
                line = UCase(fContact.ReadLine)
                n = n - CInt(Left(line, 11) = "BEGIN:VCARD") 'When BEGIN:VCARD found, -cint= 1
            Loop
            fContact.Close
         End If
        
        If n > 1 Then
            MsgBox (oSourceFile.Name & vbCrLf & "already multi")
            GoTo ExitSub:
        End If
    Next
    


    'Combine vCards
    Debug.Print vbCrLf & "Combining files"
    Set oMultivCard = fso.CreateTextFile(MultivCard)
    finePrint = 0
    For Each oSourceFile In oSourceFolder.Files
        finePrint = finePrint + 1
        If finePrint / 50 = Int(finePrint / 50) Then Debug.Print
        Debug.Print ".";
        If UCase(fso.GetExtensionName(oSourceFile)) = UCase("vcf") Then
            Set fContact = fso.OpenTextFile(oSourceFile.Path, ForReading, False)
            oMultivCard.WriteLine fContact.ReadAll
            fContact.Close
        End If
        
        If n > 1 Then
            MsgBox (oSourceFile.Name & vbCrLf & "already multi")
            Exit For
        End If
    Next
    oMultivCard.Close
    
    Debug.Print vbCrLf & "Combined " & finePrint & " files."
    
    
    
        
'Set ol = New LookOut
'ol.MultiVcf "%USERPROFILE%\desktop\exportVcf", "%USERPROFILE%\desktop\ciaoooo.vcf"
    
    
ExitSub:
    Set oContact = Nothing
    Set fContact = Nothing
    Set oSourceFolder = Nothing
    Set oSourceFile = Nothing
    Set oMultivCard = Nothing
    Set fso = Nothing
    Set oWsh = Nothing



End Sub
'/**
' Subtract elements of an array B from an array A.
' \param A array of source elements to which subtract elements.
' \param B array of elements to subtract to A
'**/
Private Function ArrayDiff(A() As String, B() As String) As String()

    Dim aItem As Variant
    Dim C() As String: C = Split("")
    Dim i As Integer: i = -1
        
    'If an element of A is not in B, keep it
    'Kept items are set in a new arrat C
    For Each aItem In A
        If UBound(Filter(B, aItem)) < 0 Then
            i = i + 1
            ReDim Preserve C(i)
            C(i) = aItem
        End If
    Next aItem
    
    ArrayDiff = C

End Function


