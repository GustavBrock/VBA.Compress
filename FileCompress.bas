Attribute VB_Name = "FileCompress"
Option Compare Text
Option Explicit

' Compression and decompression methods v1.3.0
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Compress
'
' Set of functions to zip, unzip, compress, decompress, archive, and dearchive
' zip, cab (cabinet), and tar and tgz (archive) files and folders.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)


' Select Early Binding (True) or Late Binding (False).
#Const EarlyBinding = True
        
                                               
' General constants.
'
' Wait forever.
Private Const Infinite              As Long = &HFFFF

' Process Security and Access Rights.
'
' The right to use the object for synchronization.
' This enables a thread to wait until the object is in the signaled state.
Private Const Synchronize           As Long = &H100000

' Constants for WaitForSingleObject.
'
' The specified object is a mutex object that was not released by the thread
' that owned the mutex object before the owning thread terminated.
' Ownership of the mutex object is granted to the calling thread and the
' mutex state is set to nonsignaled.
Private Const StatusAbandonedWait0  As Long = &H80
Private Const WaitAbandoned         As Long = StatusAbandonedWait0 + 0
' The state of the specified object is signaled.
Private Const StatusWait0           As Long = &H0
Private Const WaitObject0           As Long = StatusWait0 + 0
' The time-out interval elapsed, and the object's state is nonsignaled.
Private Const WaitTimeout           As Long = &H102
' The function has failed. To get extended error information, call GetLastError.
Private Const WaitFailed            As Long = &HFFFFFFFF


' Missing enum when using late binding.
'
#If EarlyBinding = False Then

    Public Enum IOMode
        ForAppending = 8
        ForReading = 1
        ForWriting = 2
    End Enum
    
    Public Enum DriveTypeConst
        Unknown = 0
        Removable = 1
        Fixed = 2
        Network = 3
        CDRom = 4
        RamDisk = 5
    End Enum
    
    Public Enum SpecialFolder
        WindowsFolder = 0       ' The Windows folder contains files installed by the Windows operating system.
        SystemFolder = 1        ' The System folder contains libraries, fonts, and device drivers.
        TemporaryFolder = 2     ' The Temp folder is used to store temporary files.
    End Enum
    
    Public Enum FileAttribute
        Normal = 0              ' Normal file. No attributes are set.
        ReadOnly = 1            ' Read-only file. Attribute is read/write.
        Hidden = 2              ' Hidden file. Attribute is read/write.
        System = 4              ' System file. Attribute is read/write.
        Volume = 8              ' Disk drive volume label. Attribute is read-only.
        Directory = 16          ' Folder or directory. Attribute is read-only.
        Archive = 32            ' File has changed since last backup. Attribute is read/write.
        Alias = 1024            ' Link or shortcut. Attribute is read-only.
        Compressed = 2048       ' Compressed file. Attribute is read-only.
    End Enum
    
#End If


' API declarations.

' Opens an existing local process object.
' If the function succeeds, the return value is an open handle
' to the specified process.
' If the function fails, the return value is NULL (0).
' To get extended error information, call GetLastError.
'
#If VBA7 Then
    Private Declare PtrSafe Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) _
        As LongPtr
#Else
    Private Declare Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) _
        As Long
#End If

' The WaitForSingleObject function returns when one of the following occurs:
' - the specified object is in the signaled state.
' - the time-out interval elapses.
'
' The dwMilliseconds parameter specifies the time-out interval, in milliseconds.
' The function returns if the interval elapses, even if the object's state is
' nonsignaled.
' If dwMilliseconds is zero, the function tests the object's state and returns
' immediately.
' If dwMilliseconds is Infinite, the function's time-out interval never elapses.
'
#If VBA7 Then
    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" ( _
        ByVal hHandle As LongPtr, _
        ByVal dwMilliseconds As Long) _
        As Long
#Else
    Private Declare Function WaitForSingleObject Lib "kernel32" ( _
        ByVal hHandle As Long, _
        ByVal dwMilliseconds As Long) _
        As Long
#End If

' Closes an open object handle.
' If the function succeeds, the return value is nonzero.
' If the function fails, the return value is zero.
' To get extended error information, call GetLastError.
'
#If VBA7 Then
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As LongPtr) _
        As Long
#Else
    Private Declare Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) _
        As Long
#End If

' Suspends the execution of the current thread until the time-out interval elapses.
'
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" ( _
        ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" ( _
        ByVal dwMilliseconds As Long)
#End If

' Compress a file or a folder to a cabinet file/folder.
'
' A single file will be compressed to a file with a ".*_" extension by default,
' optionally with a ".cab" extension.
' A folder will be compressed to a file with a ".cab" extension by default,
' optionally with a custom extension.
'
' Parameters:
'   Path:
'       Valid (UNC) path to the file or folder to compress.
'   Destination:
'       (Optional) Valid (UNC) path to a folder or to a file with a
'       cabinet extension or other extension.
'   Overwrite:
'       (Optional) Overwrite (default) or leave an existing cabinet file.
'       If False, the created cabinet file will be versioned:
'           Example.cab, Example (2).cab, etc.
'       If True, an existing cabinet file will first be deleted, then recreated.
'   SingleFileExtension:
'       (Optional) ".*_" style or (default) ".cab" file extension.
'       If False, the created cabinet file extension will be "cab".
'       If True and source file's extension's last character is not an underscore,
'       the created cabinet file extension will be named as the source file,
'       but with an underscore as the last character of the extension.
'       In both cases, a specified Destination filename will override this setting.
'   HighCompression:
'       (Optional) Use standard compression or high compression.
'       If False, use standard MSZIP compression. Faster, but larger file size.
'       If True, use LZX compression. Slower, but smaller file size.
'
'   Path and Destination can be relative paths. If so, the current path is used.
'
'   If success, 0 is returned, and Destination holds the full path of the created cabinet file.
'   If error, error code is returned, and Destination will be zero length string.
'
' Early binding requires references to:
'
'   Shell:
'       Microsoft Shell Controls And Automation
'
'   Scripting.FileSystemObject:
'       Microsoft Scripting Runtime
'
' 2022-04-20. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function Cab( _
    ByVal Path As String, _
    Optional ByRef Destination As String, _
    Optional ByVal OverWrite As Boolean = True, _
    Optional ByVal SingleFileExtension As Boolean, _
    Optional ByVal HighCompression As Boolean = True) _
    As Long
    
#If EarlyBinding Then
    ' Microsoft Scripting Runtime.
    Dim FileSystemObject    As Scripting.FileSystemObject
    ' Microsoft Shell Controls And Automation.
    Dim ShellApplication    As Shell
    
    Set FileSystemObject = New Scripting.FileSystemObject
    Set ShellApplication = New Shell
#Else
    Dim FileSystemObject    As Object
    Dim ShellApplication    As Object

    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
    Set ShellApplication = CreateObject("Shell.Application")
#End If
    
    ' Extension of a cabinet file holding one or more files.
    Const CabExtensionName  As String = "cab"
    Const CabExtension      As String = "." & CabExtensionName
    ' Extension of a cabinet file holding one file only.
    Const CabExtensionName1 As String = "*_"
    ' Extension for a Disk Directive File for MakeCab.exe.
    Const DdfExtensionName  As String = "ddf"
    Const DdfExtension      As String = "." & DdfExtensionName
    ' Custom error values.
    Const ErrorPathFile     As Long = 75
    Const ErrorOther        As Long = -1
    Const ErrorNone         As Long = 0
    ' Maximum (arbitrary) allowed count of created cabinet versions.
    Const MaxCabVersion     As Integer = 1000
    
    ' MakeCab directive constants.
    Const CompressionHigh   As String = "LZX"
    Const CompressionLow    As String = "MSZIP"
    
    Dim FileNames           As Variant
    
    Dim CabPath             As String
    Dim CabName             As String
    Dim CabFile             As String
    Dim CabBase             As String
    Dim CabTemp             As String
    Dim CabMono             As Boolean
    Dim Extension           As String
    Dim ExtensionName       As String
    Dim Version             As Integer
    Dim Item                As Long
    Dim PathName            As String
    Dim CurrentDirectory    As String
    Dim TempDirectory       As String
    Dim Result              As Long
    
    If FileSystemObject.FileExists(Path) Then
        ' The source is an existing file.
        CabMono = True
        CabName = FileSystemObject.GetFileName(Path)
        If SingleFileExtension = True Then
            ExtensionName = FileSystemObject.GetExtensionName(Path)
            ' Check if the file already has a cabinet-style extension.
            If Right(ExtensionName, 1) = Right(CabExtensionName1, 1) Then
                ' Remove extension.
                ExtensionName = ""
                ' Add cabinet extension later.
            Else
                ' Apply cabinet-style extension.
                Mid(CabName, Len(CabName)) = Right(CabExtensionName1, 1)
                ExtensionName = FileSystemObject.GetExtensionName(CabName)
            End If
        End If
        If ExtensionName = "" Then
            CabName = FileSystemObject.GetBaseName(Path) & CabExtension
            ExtensionName = FileSystemObject.GetExtensionName(CabName)
        End If
        CabPath = FileSystemObject.GetFile(Path).ParentFolder
        Extension = "." & ExtensionName
    ElseIf FileSystemObject.FolderExists(Path) Then
        ' The source is an existing folder or fileshare.
        CabBase = FileSystemObject.GetBaseName(Path)
        If CabBase <> "" Then
            ' Path is a folder.
            CabName = CabBase & CabExtension
            CabPath = FileSystemObject.GetFolder(Path).ParentFolder
            Extension = CabExtension
        Else
            ' Path is a fileshare, thus has no parent folder.
        End If
    Else
        ' The source does not exist.
    End If
       
    If CabName = "" Then
        ' Nothing to compress. Exit.
        Destination = ""
    Else
        If Destination <> "" Then
            If FileSystemObject.GetExtensionName(Destination) = "" Then
                ' Destination is a folder.
                If FileSystemObject.FolderExists(Destination) Then
                    CabPath = Destination
                Else
                    ' No folder for the cabinet file. Exit.
                    Destination = ""
                End If
            Else
                ' Destination is a single compressed file.
                CabName = FileSystemObject.GetFileName(Destination)
                If CabName = Destination Then
                    ' No path given. Use CabPath as is.
                Else
                    ' Use path of Destination.
                    CabPath = FileSystemObject.GetParentFolderName(Destination)
                End If
            End If
        Else
            ' Use the already found folder of the source.
            Destination = CabPath
        End If
    End If
    
    If Destination <> "" Then
        CabFile = FileSystemObject.BuildPath(CabPath, CabName)
        
        If FileSystemObject.FileExists(CabFile) Then
            If OverWrite = True Then
                ' Delete an existing file.
                FileSystemObject.DeleteFile CabFile, True
                ' At this point either the file is deleted or an error is raised.
            Else
                CabBase = FileSystemObject.GetBaseName(CabFile)
                ExtensionName = FileSystemObject.GetExtensionName(CabFile)
                If ExtensionName <> CabExtensionName Then
                    Extension = "." & ExtensionName
                End If
                ' Modify name of the cabinet file to be created to preserve an existing file:
                '   "Example.cab" -> "Example (2).cab", etc.
                Version = Version + 1
                Do
                    Version = Version + 1
                    CabFile = FileSystemObject.BuildPath(CabPath, CabBase & Format(Version, " \(0\)") & Extension)
                Loop Until FileSystemObject.FileExists(CabFile) = False Or Version > MaxCabVersion
                If Version > MaxCabVersion Then
                    ' Give up.
                    Err.Raise ErrorPathFile, "Cab Create", "File could not be created."
                End If
                CabName = FileSystemObject.GetFileName(CabFile)
            End If
        End If
        ' Set returned file name.
        Destination = CabFile
        
        ' Get list of files to compress.
        FileNames = FolderFileNames(Path)
        
        ' Prepare a temporary ddf file to control makecab.exe.
        CabTemp = FileSystemObject.BuildPath(CabPath, FileSystemObject.GetBaseName(FileSystemObject.GetTempName()) & DdfExtension)
        ' Resolve relative paths.
        CabTemp = FileSystemObject.GetAbsolutePathName(CabTemp)
        Path = FileSystemObject.GetAbsolutePathName(Path)
        
        ' Build the directive file.
        With FileSystemObject.OpenTextFile(CabTemp, ForWriting, True)
            .Write ".Set CabinetName1=""" & CabName & """" & vbCrLf
            .Write ".Set CompressionMemory=21" & vbCrLf
            .Write ".Set CompressionType=" & IIf(HighCompression, CompressionHigh, CompressionLow) & vbCrLf
            .Write ".Set DiskDirectoryTemplate=""" & CabPath & """" & vbCrLf
            .Write ".Set MaxDiskSize=0" & vbCrLf
            .Write ".Set InfFileName=NUL" & vbCrLf
            .Write ".Set RptFileName=NUL" & vbCrLf
            .Write ".Set UniqueFiles=OFF" & vbCrLf
            .Write ".Set SourceDir=""" & IIf(CabMono, FileSystemObject.GetParentFolderName(Path), Path) & """" & vbCrLf
            ' Append list of files to compress.
            For Item = LBound(FileNames) To UBound(FileNames)
                .Write """" & FileNames(Item) & """" & vbCrLf
            Next
            .Close
        End With
        
        ' Record the current directory.
        CurrentDirectory = CurDir
        ' Change current directory to temp folder.
        TempDirectory = Environ("temp")
        ChDrive TempDirectory
        ChDir TempDirectory

        ' Create the cabinet file.
        PathName = "makecab.exe /v1 /f """ & CabTemp & """"
        ' ShellWait returns True for no errors.
        Result = ShellWait("cmd /c " & PathName & "", vbMinimizedNoFocus)
        
        ' Restore the current directory.
        ChDrive CurrentDirectory
        ChDir CurrentDirectory
        
        ' Delete the directive file.
        FileSystemObject.DeleteFile CabTemp, True
    End If
    
    Set ShellApplication = Nothing
    Set FileSystemObject = Nothing
    
    If Err.Number <> ErrorNone Then
        Destination = ""
        Result = Err.Number
    ElseIf Destination = "" Then
        Result = ErrorOther
    End If
    
    Cab = Result

End Function

' Extract files from a cabinet file to a folder using Windows Explorer.
'
' Parameters:
'   Path:
'       Valid (UNC) path to a valid cabinet file. Extension can be another than "cab".
'   Destination:
'       (Optional) Valid (UNC) path to the destination folder.
'   Overwrite:
'       (Optional) Overwrite (default) or leave an existing folder.
'       If False, an existing folder will keep other files than those in the extracted cabinet file.
'       If True, an existing folder will first be deleted, then recreated.
'
'   Path and Destination can be relative paths. If so, the current path is used.
'
'   If success, 0 is returned, and Destination holds the full path of the created folder.
'   If error, error code is returned, and Destination will be zero length string.
'
' Early binding requires references to:
'
'   Shell:
'       Microsoft Shell Controls And Automation
'
'   Scripting.FileSystemObject:
'       Microsoft Scripting Runtime
'
' 2023-10-28. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function DeCab( _
    ByVal Path As String, _
    Optional ByRef Destination As String, _
    Optional ByVal OverWrite As Boolean = True) _
    As Long
    
#If EarlyBinding Then
    ' Microsoft Scripting Runtime.
    Dim FileSystemObject    As Scripting.FileSystemObject
    ' Microsoft Shell Controls And Automation.
    Dim ShellApplication    As Shell
    
    Set FileSystemObject = New Scripting.FileSystemObject
    Set ShellApplication = New Shell
#Else
    Dim FileSystemObject    As Object
    Dim ShellApplication    As Object

    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
    Set ShellApplication = CreateObject("Shell.Application")
#End If
               
    ' Extension of a cabinet file holding one or more files.
    Const CabExtensionName  As String = "cab"
    ' Extension of a cabinet file holding one file only.
    Const CabExtensionName1 As String = "??_"
    Const CabExtension      As String = "." & CabExtensionName
    ' Extension of an archive file holding one or more files.
    Const TarExtensionName  As String = "tar"
    ' Extension of a compressed archive file holding one or more files.
    Const TgzExtensionName  As String = "tgz"
    ' Mandatory extension of zip file.
    Const ZipExtensionName  As String = "zip"
    ' Custom error values.
    Const ErrorNone         As Long = 0
    Const ErrorOther        As Long = -1
    
    Dim CabName             As String
    Dim CabPath             As String
    Dim CabTemp             As String
    Dim CabMono             As Boolean
    Dim Result              As Long
    
    If FileSystemObject.FileExists(Path) Then
        ' The source is an existing file.
        CabName = FileSystemObject.GetBaseName(Path)
        CabPath = FileSystemObject.GetFile(Path).ParentFolder
        ' Check if the extension matches that of a cabinet file holding one file only.
        CabMono = FileSystemObject.GetExtensionName(Path) Like CabExtensionName1
    End If
    
    If CabName = "" Then
        ' Nothing to extract. Exit.
        Destination = ""
    Else
        ' Select or create destination folder.
        If Destination <> "" Then
            ' Extract to a custom folder.
            If _
                FileSystemObject.GetExtensionName(Destination) = CabExtensionName Or _
                FileSystemObject.GetExtensionName(Destination) = TarExtensionName Or _
                FileSystemObject.GetExtensionName(Destination) = TgzExtensionName Or _
                FileSystemObject.GetExtensionName(Destination) = ZipExtensionName Then
                ' Do not extract to a folder named *.cab, *.tar, or *.zip.
                ' Strip extension.
                Destination = FileSystemObject.BuildPath( _
                    FileSystemObject.GetParentFolderName(Destination), _
                    FileSystemObject.GetBaseName(Destination))
            End If
        Else
            If CabMono Then
                ' Single-file cabinet.
                ' Extract to the folder of the cabinet file.
                Destination = CabPath
            Else
                ' Multiple-files cabinet.
                ' Extract to a subfolder of the folder of the cabinet file.
                Destination = FileSystemObject.BuildPath(CabPath, CabName)
            End If
        End If
            
        If FileSystemObject.FolderExists(Destination) Then
            If OverWrite = True Then
                ' Delete the existing folder.
                FileSystemObject.DeleteFolder Destination, True
            ElseIf FileSystemObject.GetFolder(Destination).Files.Count > 0 Then
                ' Files exists and should not be overwritten.
                ' Exit.
                Destination = ""
            End If
        End If
        If Destination <> "" Then
            If Not FileSystemObject.FolderExists(Destination) Then
                ' Create the destination folder.
                FileSystemObject.CreateFolder Destination
            End If
        End If
        
        If Not FileSystemObject.FolderExists(Destination) Then
            ' For some reason the destination folder does not exist and cannot be created.
            ' Exit.
            Destination = ""
        ElseIf Destination <> "" Then
            ' Destination folder existed or has been created successfully.
            ' Resolve relative paths.
            Destination = FileSystemObject.GetAbsolutePathName(Destination)
            Path = FileSystemObject.GetAbsolutePathName(Path)
            ' Check file extension.
            If FileSystemObject.GetExtensionName(Path) = CabExtensionName Then
                ' File extension is OK.
                CabTemp = Path
            Else
                ' Rename the cabinet file by adding a cabinet extension.
                CabTemp = Path & CabExtension
                FileSystemObject.MoveFile Path, CabTemp
            End If
            ' Extract files and folders from the cabinet file to the destination folder.
            ' Note, that when copying from a cab file, overwrite flag is ignored.
            ShellApplication.Namespace(CVar(Destination)).CopyHere ShellApplication.Namespace(CVar(CabTemp)).Items
            If CabTemp <> Path Then
                ' Remove the cabinet extension to restore the original file name.
                FileSystemObject.MoveFile CabTemp, Path
            End If
        End If
    End If
    
    Set ShellApplication = Nothing
    Set FileSystemObject = Nothing
    
    If Err.Number <> ErrorNone Then
        Destination = ""
        Result = Err.Number
    ElseIf Destination = "" Then
        Result = ErrorOther
    End If
    
    DeCab = Result
     
End Function

' Returns an array of file names of the specified Path
' and its subfolders including subfolder name but without
' the root path (drive letter and parent folder).
' Names of subfolders themselves are excluded.
'
' The array holds one file name, if Path is a file.
'
' Will fail if permission to a subfolder is denied.
'
' Example:
'   FileNameArray = FolderFileNames("C:\Windows\bootstat.dat")
'   will hold:
'       bootstat.dat
'
'   FileNameArray = FolderFileNames("C:\Windows")
'   will hold:
'       bfsvc.exe
'       bootstat.dat
'       ...
'       addins\FXSEXT.ecf
'       appcompat\appraiser\APPRAISER_FileInventory.xml
'       ...
'
' Format is similar to the DOS command with no root path:
'   Dir "C:\Windows" /A:-D /B /S
'   that will output:
'       C:\Windows\bfsvc.exe
'       C:\Windows\bootstat.dat
'       ...
'       C:\Windows\addins\FXSEXT.ecf
'       C:\Windows\appcompat\appraiser\APPRAISER_FileInventory.xml
'       ...
'
' Parameter ParentFolderName is for internal use only and
' must not be specified initially.
'
' 2018-07-25. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function FolderFileNames( _
    ByVal Path As String, _
    Optional ByVal ParentFolderName As String) _
    As Variant

#If EarlyBinding Then
    ' Microsoft Scripting Runtime.
    Dim FileSystemObject    As Scripting.FileSystemObject
    Dim Folder              As Scripting.Folder
    Dim SubFolder           As Scripting.Folder
    Dim Files               As Scripting.Files
    Dim File                As Scripting.File
    
    Set FileSystemObject = New Scripting.FileSystemObject
#Else
    Dim FileSystemObject    As Object
    Dim Folder              As Object
    Dim SubFolder           As Object
    Dim Files               As Object
    Dim File                As Object

    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
#End If

    Dim FileList            As Variant
    Dim FileListSub         As Variant

    If FileSystemObject.FolderExists(Path) Then
        Set Folder = FileSystemObject.GetFolder(Path)
        Set Files = Folder.Files
    
        For Each File In Files
            If IsEmpty(FileList) Then
                FileList = Array(FileSystemObject.BuildPath(ParentFolderName, File.Name))
            Else
                FileList = Split(Join(FileList, ":") & ":" & FileSystemObject.BuildPath(ParentFolderName, File.Name), ":")
            End If
        Next
        For Each SubFolder In Folder.SubFolders
            FileListSub = FolderFileNames(SubFolder.Path, FileSystemObject.BuildPath(ParentFolderName, FileSystemObject.GetBaseName(SubFolder)))
            If Not IsEmpty(FileListSub) Then
                If IsEmpty(FileList) Then
                    FileList = FileListSub
                Else
                    FileList = Split(Join(FileList, ":") & ":" & Join(FileListSub, ":"), ":")
                End If
            End If
        Next
    ElseIf FileSystemObject.FileExists(Path) Then
        FileList = Array(FileSystemObject.GetFile(Path).Name)
    Else
        ' Nothing to return.
        ' Return Empty.
    End If
    
    FolderFileNames = FileList
    
End Function

' Check if (a file of) a folder is aliased, meaning located
' on a drive linked to a folder on another drive.
' Returns True if the folder is linked, False if not.
'
' 2022-03-26. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function IsFolderAlias( _
    ByVal Path As String) _
    As Boolean
    
#If EarlyBinding Then
    ' Microsoft Scripting Runtime.
    Dim FileSystemObject    As Scripting.FileSystemObject
    Dim Folder              As Folder
    
    Set FileSystemObject = New Scripting.FileSystemObject
#Else
    Dim FileSystemObject    As Object
    Dim Folder              As Object

    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
#End If

    Dim IsAlias             As Boolean
    
    If FileSystemObject.FileExists(Path) Then
        Path = FileSystemObject.GetParentFolderName(Path)
    End If
    If FileSystemObject.FolderExists(Path) Then
        Set Folder = FileSystemObject.GetFolder(Path)
        
        Do While Not Folder.IsRootFolder
            If (Folder.Attributes And Alias) = Alias Then
                IsAlias = True
                Exit Do
            End If
            Set Folder = Folder.ParentFolder
        Loop
        
        Set Folder = Nothing
    End If
    
    Set FileSystemObject = Nothing
    
    IsFolderAlias = IsAlias

End Function

' Lists the files of a folder and its subfolders
' including the subfolder name but without the
' root path (drive letter and parent folder).
'
' Returns the count of files.
'
' Will fail if permission to a subfolder is denied.
'
' Example:
'   FileCount = ListFolderFiles("C:\Windows")
'   will list:
'       bfsvc.exe
'       bootstat.dat
'       ...
'       addins\FXSEXT.ecf
'       appcompat\appraiser\APPRAISER_FileInventory.xml
'       ...
'
' 2017-10-22. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function ListFolderFiles( _
    ByVal Path As String) _
    As Long
    
    Dim FileNames   As Variant
    Dim Item        As Long
    
    FileNames = FolderFileNames(Path)
    If Not IsEmpty(FileNames) Then
        For Item = LBound(FileNames) To UBound(FileNames)
            Debug.Print FileNames(Item)
        Next
    End If
    
    ListFolderFiles = Item

End Function

' Shells out to an external process and waits until the process ends.
' Returns 0 (zero) for no errors, or an error code.
'
' The call will wait for an infinite amount of time for the process to end.
' The process will seem frozen until the shelled process terminates. Thus,
' if the shelled process hangs, so will this.
'
' A better approach could be to wait a specific amount of time and, when the
' time-out interval expires, test the return value. If it is WaitTimeout, the
' process is still not signaled. Then either wait again or continue with the
' processing.
'
' Waiting for a DOS application is different, as the DOS window doesn't close
' when the application is done.
' To avoid this, prefix the application command called (shelled to) with:
' "command.com /c " or "cmd.exe /c ".
'
' For example:
'   Command = "cmd.exe /c " & Command
'   Result = ShellWait(Command)
'
' 2018-04-06. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function ShellWait( _
    ByVal Command As String, _
    Optional ByVal WindowStyle As VbAppWinStyle = vbNormalNoFocus) _
    As Long

    Const InheritHandle As Long = &H0
    Const NoProcess     As Long = 0
    Const NoHandle      As Long = 0
    
#If VBA7 Then
    Dim ProcessHandle   As LongPtr
#Else
    Dim ProcessHandle   As Long
#End If
    Dim DesiredAccess   As Long
    Dim ProcessId       As Long
    Dim WaitTime        As Long
    Dim Closed          As Boolean
    Dim Result          As Long
  
    If Len(Trim(Command)) = 0 Then
        ' Nothing to do. Exit.
    Else
        ProcessId = Shell(Command, WindowStyle)
        If ProcessId = NoProcess Then
            ' Process could not be started.
        Else
            ' Get a handle to the shelled process.
            DesiredAccess = Synchronize
            ProcessHandle = OpenProcess(DesiredAccess, InheritHandle, ProcessId)
            ' Wait "forever".
            WaitTime = Infinite
            ' If successful, wait for the application to end and close the handle.
            If ProcessHandle = NoHandle Then
                ' Should not happen.
            Else
                ' Process is running.
                Result = WaitForSingleObject(ProcessHandle, WaitTime)
                ' Process ended.
                Select Case Result
                    Case WaitObject0
                        ' Success.
                    Case WaitAbandoned, WaitTimeout, WaitFailed
                        ' Know error.
                    Case Else
                        ' Other error.
                End Select
                ' Close process.
                Closed = CBool(CloseHandle(ProcessHandle))
                If Result = WaitObject0 Then
                    ' Return error if not closed.
                    Result = Not Closed
                End If
            End If
        End If
    End If
  
    ShellWait = Result

End Function

' Archive files or folders using Tar to an archive folder.
'
' The files or folders will be compressed to a file with a ".tgz" extension by default,
' optionally with no compression to a file with a ".tar" extension, optionally with a custom extension.
'
' Notes:
'   Neither bzip2 nor lzma are implemented as these either are slower or don't create smaller archives.
'
'   The format of the tar file is the default ustar.
'   The alternative formats, pax, cpio, and shar, are not implemented.
'
' Parameters:
'   Path:
'       Valid (UNC) path to the files or folders to archive.
'   Destination:
'       (Optional) Valid (UNC) path to a folder or to a file with a
'       cabinet extension or other extension.
'   Overwrite:
'       (Optional) Overwrite (default) or leave an existing archive file.
'       If False, the created archive file will be versioned:
'           Example.tar, Example (2).tar, etc.
'       If True, an existing archive file will first be deleted, then recreated.
'   HighCompression:
'       (Optional) Use standard compression or high compression and use the tgz extension by default.
'       If False, use standard gzip compression. Faster, but larger file size.
'       If True, use xz compression. Slower, but smaller file size.
'       Ignored if parameter NoCompression is True.
'   NoCompression:
'       (Optional) Use compression or no compression and, if True, use the tar extension by default.
'       If False, use compression.
'       If True, use no compression. Much faster speed, similar to file copying.
'
'   Path and Destination can be relative paths. If so, the current path is used.
'
'   If success, 0 is returned, and Destination holds the full path of the created archive file.
'   If error, error code is returned, and Destination will be zero length string.
'
' Early binding requires references to:
'
'   Shell:
'       Microsoft Shell Controls And Automation
'
'   Scripting.FileSystemObject:
'       Microsoft Scripting Runtime
'
' 2023-10-28. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function Tar( _
    ByVal Path As String, _
    Optional ByRef Destination As String, _
    Optional ByVal OverWrite As Boolean = True, _
    Optional ByVal HighCompression As Boolean = False, _
    Optional ByVal NoCompression As Boolean = False) _
    As Long
    
#If EarlyBinding Then
    ' Microsoft Scripting Runtime.
    Dim FileSystemObject    As Scripting.FileSystemObject
    ' Microsoft Shell Controls And Automation.
    Dim ShellApplication    As Shell
    
    Set FileSystemObject = New Scripting.FileSystemObject
    Set ShellApplication = New Shell
#Else
    Dim FileSystemObject    As Object
    Dim ShellApplication    As Object

    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
    Set ShellApplication = CreateObject("Shell.Application")
#End If
    
    ' Extension of an archive file with or without compression holding one or more files.
    Const GzExtensionName   As String = "tgz"
    Const NoExtensionName   As String = "tar"
    
    ' Custom error values.
    Const ErrorPathFile     As Long = 75
    Const ErrorOther        As Long = -1
    Const ErrorNone         As Long = 0
    ' Maximum (arbitrary) allowed count of created archive versions.
    Const MaxTarVersion     As Integer = 1000
    
    ' Tar directive constants (note: case sensitive).
    Const CompressionHigh   As String = "J"
    Const CompressionLow    As String = "z"
    Const CompressionNone   As String = ""
    
    Dim TarExtensionName    As String
    Dim TarExtension        As String
       
    Dim TarPath             As String
    Dim TarName             As String
    Dim TarFile             As String
    Dim TarBase             As String
    Dim Extension           As String
    Dim ExtensionName       As String
    Dim Version             As Integer
    Dim PathName            As String
    Dim Result              As Long
    
    ' Extension of an archive file holding one or more files.
    If NoCompression = False Then
        TarExtensionName = GzExtensionName
    Else
        TarExtensionName = NoExtensionName
    End If
    TarExtension = "." & TarExtensionName
    
    If FileSystemObject.FileExists(Path) Then
        ' The source is an existing file.
        TarName = FileSystemObject.GetFileName(Path)
        If ExtensionName = "" Then
            TarName = FileSystemObject.GetBaseName(Path) & TarExtension
            ExtensionName = FileSystemObject.GetExtensionName(TarName)
        End If
        TarPath = FileSystemObject.GetFile(Path).ParentFolder
        Extension = "." & ExtensionName
    ElseIf FileSystemObject.FolderExists(Path) Then
        ' The source is an existing folder or fileshare.
        TarBase = FileSystemObject.GetBaseName(Path)
        If TarBase <> "" Then
            ' Path is a folder.
            TarName = TarBase & TarExtension
            TarPath = FileSystemObject.GetFolder(Path).ParentFolder
            Extension = TarExtension
        Else
            ' Path is a fileshare, thus has no parent folder.
        End If
    Else
        ' The source does not exist.
    End If
       
    If TarName = "" Then
        ' Nothing to compress. Exit.
        Destination = ""
    Else
        If Destination <> "" Then
            If FileSystemObject.GetExtensionName(Destination) = "" Then
                ' Destination is a folder.
                If FileSystemObject.FolderExists(Destination) Then
                    TarPath = Destination
                Else
                    ' No folder for the cabinet file. Exit.
                    Destination = ""
                End If
            Else
                ' Destination is a single compressed file.
                TarName = FileSystemObject.GetFileName(Destination)
                If TarName = Destination Then
                    ' No path given. Use TarPath as is.
                Else
                    ' Use path of Destination.
                    TarPath = FileSystemObject.GetParentFolderName(Destination)
                End If
            End If
        Else
            ' Use the already found folder of the source.
            Destination = TarPath
        End If
    End If
    
    If Destination <> "" Then
        TarFile = FileSystemObject.BuildPath(TarPath, TarName)
        
        If FileSystemObject.FileExists(TarFile) Then
            If OverWrite = True Then
                ' Delete an existing file.
                FileSystemObject.DeleteFile TarFile, True
                ' At this point either the file is deleted or an error is raised.
            Else
                TarBase = FileSystemObject.GetBaseName(TarFile)
                ExtensionName = FileSystemObject.GetExtensionName(TarFile)
                If ExtensionName <> TarExtensionName Then
                    Extension = "." & ExtensionName
                End If
                ' Modify name of the archive file to be created to preserve an existing file:
                '   "Example.tar" -> "Example (2).tar", etc.
                Version = Version + 1
                Do
                    Version = Version + 1
                    TarFile = FileSystemObject.BuildPath(TarPath, TarBase & Format(Version, " \(0\)") & Extension)
                Loop Until FileSystemObject.FileExists(TarFile) = False Or Version > MaxTarVersion
                If Version > MaxTarVersion Then
                    ' Give up.
                    Err.Raise ErrorPathFile, "Tar Create", "File could not be created."
                End If
                TarName = FileSystemObject.GetFileName(TarFile)
            End If
        End If
        ' Set returned file name.
        Destination = TarFile
        ' Create the archive file.
        PathName = "tar.exe -c{2}f ""{1}"" ""{0}"""
        PathName = Replace(PathName, "{0}", Path)
        PathName = Replace(PathName, "{1}", Destination)
        If NoCompression = False Then
            PathName = Replace(PathName, "{2}", IIf(HighCompression, CompressionHigh, CompressionLow))
        Else
            PathName = Replace(PathName, "{2}", CompressionNone)
        End If
        ' ShellWait returns True for no errors.
        Result = ShellWait("cmd /c " & PathName & "", vbMinimizedNoFocus)
    End If
    
    Set ShellApplication = Nothing
    Set FileSystemObject = Nothing
    
    If Err.Number <> ErrorNone Then
        Destination = ""
        Result = Err.Number
    ElseIf Destination = "" Then
        Result = ErrorOther
    End If
    
    Tar = Result

End Function

' Dearchive ("untar") files from a tar file to a folder using Windows Explorer.
' Default behaviour is similar to right-clicking a file/folder and selecting:
'   Unpack all ...
'
' Parameters:
'   Path:
'       Valid (UNC) path to a valid archive file. Extension can be another than "tar".
'   Destination:
'       (Optional) Valid (UNC) path to the destination folder.
'   Overwrite:
'       (Optional) Leave (default) or overwrite an existing folder.
'       If False, an existing folder will keep other files than those in the dearchived tar file.
'       If True, an existing folder will first be deleted, then recreated.
'
'   Path and Destination can be relative paths. If so, the current path is used.
'
'   If success, 0 is returned, and Destination holds the full path of the created folder.
'   If error, error code is returned, and Destination will be zero length string.
'
' Early binding requires references to:
'
'   Shell:
'       Microsoft Shell Controls And Automation
'
'   Scripting.FileSystemObject:
'       Microsoft Scripting Runtime
'
' 2023-10-28. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function UnTar( _
    ByVal Path As String, _
    Optional ByRef Destination As String, _
    Optional ByVal OverWrite As Boolean) _
    As Long
    
#If EarlyBinding Then
    ' Microsoft Scripting Runtime.
    Dim FileSystemObject    As Scripting.FileSystemObject
    ' Microsoft Shell Controls And Automation.
    Dim ShellApplication    As Shell
    
    Set FileSystemObject = New Scripting.FileSystemObject
    Set ShellApplication = New Shell
#Else
    Dim FileSystemObject    As Object
    Dim ShellApplication    As Object

    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
    Set ShellApplication = CreateObject("Shell.Application")
#End If
               
    ' Extension of a cabinet file holding one or more files.
    Const CabExtensionName  As String = "cab"
    ' Extension of an archive file holding one or more files.
    Const TarExtensionName  As String = "tar"
    ' Extension of a compressed archive file holding one or more files.
    Const TgzExtensionName  As String = "tgz"
    ' Mandatory extension of zip file.
    Const ZipExtensionName  As String = "zip"
    Const TarExtension      As String = "." & TarExtensionName
    
    ' Constants for Shell.Application.
    Const DoOverwrite       As Long = &H0&
    Const NoOverwrite       As Long = &H8&
    Const YesToAll          As Long = &H10&
    ' Custom error values.
    Const ErrorNone         As Long = 0
    Const ErrorOther        As Long = -1
    
    Dim TarName             As String
    Dim TarPath             As String
    Dim TarTemp             As String
    Dim Options             As Variant
    Dim Result              As Long
    
    If FileSystemObject.FileExists(Path) Then
        ' The source is an existing file.
        TarName = FileSystemObject.GetBaseName(Path)
        TarPath = FileSystemObject.GetFile(Path).ParentFolder
    End If
    
    If TarName = "" Then
        ' Nothing to untar. Exit.
        Destination = ""
    Else
        ' Select or create destination folder.
        If Destination <> "" Then
            ' Untar to a custom folder.
            If _
                FileSystemObject.GetExtensionName(Destination) = CabExtensionName Or _
                FileSystemObject.GetExtensionName(Destination) = TarExtensionName Or _
                FileSystemObject.GetExtensionName(Destination) = TgzExtensionName Or _
                FileSystemObject.GetExtensionName(Destination) = ZipExtensionName Then
                ' Do not untar to a folder named *.cab, *.tar, tgz, or *.zip.
                ' Strip extension.
                Destination = FileSystemObject.BuildPath( _
                    FileSystemObject.GetParentFolderName(Destination), _
                    FileSystemObject.GetBaseName(Destination))
            End If
        Else
            ' Untar to a subfolder of the folder of the tarfile.
            Destination = FileSystemObject.BuildPath(TarPath, TarName)
        End If
            
        If FileSystemObject.FolderExists(Destination) And OverWrite = True Then
            ' Delete the existing folder.
            FileSystemObject.DeleteFolder Destination, True
        End If
        If Not FileSystemObject.FolderExists(Destination) Then
            ' Create the destination folder.
            FileSystemObject.CreateFolder Destination
        End If
        
        If Not FileSystemObject.FolderExists(Destination) Then
            ' For some reason the destination folder does not exist and cannot be created.
            ' Exit.
            Destination = ""
        Else
            ' Destination folder existed or has been created successfully.
            ' Resolve relative paths.
            Destination = FileSystemObject.GetAbsolutePathName(Destination)
            Path = FileSystemObject.GetAbsolutePathName(Path)
            ' Check file extension.
            If FileSystemObject.GetExtensionName(Path) = TarExtensionName Then
                ' File extension is OK.
                TarTemp = Path
            Else
                ' Rename the tar file by adding a tar extension.
                TarTemp = Path & TarExtension
                FileSystemObject.MoveFile Path, TarTemp
            End If
            ' Untar files and folders from the tar file to the destination folder.
            If OverWrite Then
                Options = DoOverwrite Or YesToAll
            Else
                Options = NoOverwrite Or YesToAll
            End If
            ShellApplication.Namespace(CVar(Destination)).CopyHere ShellApplication.Namespace(CVar(TarTemp)).Items, Options
            If TarTemp <> Path Then
                ' Remove the tar extension to restore the original file name.
                FileSystemObject.MoveFile TarTemp, Path
            End If
        End If
    End If
    
    Set ShellApplication = Nothing
    Set FileSystemObject = Nothing
    
    If Err.Number <> ErrorNone Then
        Destination = ""
        Result = Err.Number
    ElseIf Destination = "" Then
        Result = ErrorOther
    End If
    
    UnTar = Result
     
End Function

' Unzip files from a zip file to a folder using Windows Explorer.
' Default behaviour is similar to right-clicking a file/folder and selecting:
'   Unpack all ...
'
' Parameters:
'   Path:
'       Valid (UNC) path to a valid zip file. Extension can be another than "zip".
'   Destination:
'       (Optional) Valid (UNC) path to the destination folder.
'   Overwrite:
'       (Optional) Leave (default) or overwrite an existing folder.
'       If False, an existing folder will keep other files than those in the extracted zip file.
'       If True, an existing folder will first be deleted, then recreated.
'
'   Path and Destination can be relative paths. If so, the current path is used.
'
'   If success, 0 is returned, and Destination holds the full path of the created folder.
'   If error, error code is returned, and Destination will be zero length string.
'
' Early binding requires references to:
'
'   Shell:
'       Microsoft Shell Controls And Automation
'
'   Scripting.FileSystemObject:
'       Microsoft Scripting Runtime
'
' 2023-10-28. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function UnZip( _
    ByVal Path As String, _
    Optional ByRef Destination As String, _
    Optional ByVal OverWrite As Boolean) _
    As Long
    
#If EarlyBinding Then
    ' Microsoft Scripting Runtime.
    Dim FileSystemObject    As Scripting.FileSystemObject
    ' Microsoft Shell Controls And Automation.
    Dim ShellApplication    As Shell
    
    Set FileSystemObject = New Scripting.FileSystemObject
    Set ShellApplication = New Shell
#Else
    Dim FileSystemObject    As Object
    Dim ShellApplication    As Object

    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
    Set ShellApplication = CreateObject("Shell.Application")
#End If
               
    ' Extension of a cabinet file holding one or more files.
    Const CabExtensionName  As String = "cab"
    ' Extension of an archive file holding one or more files.
    Const TarExtensionName  As String = "tar"
    ' Extension of a compressed archive file holding one or more files.
    Const TgzExtensionName  As String = "tgz"
    ' Mandatory extension of zip file.
    Const ZipExtensionName  As String = "zip"
    Const ZipExtension      As String = "." & ZipExtensionName
    
    ' Constants for Shell.Application.
    Const DoOverwrite       As Long = &H0&
    Const NoOverwrite       As Long = &H8&
    Const YesToAll          As Long = &H10&
    ' Custom error values.
    Const ErrorNone         As Long = 0
    Const ErrorOther        As Long = -1
    
    Dim ZipName             As String
    Dim ZipPath             As String
    Dim ZipTemp             As String
    Dim Options             As Variant
    Dim Result              As Long
    
    If FileSystemObject.FileExists(Path) Then
        ' The source is an existing file.
        ZipName = FileSystemObject.GetBaseName(Path)
        ZipPath = FileSystemObject.GetFile(Path).ParentFolder
    End If
    
    If ZipName = "" Then
        ' Nothing to unzip. Exit.
        Destination = ""
    Else
        ' Select or create destination folder.
        If Destination <> "" Then
            ' Unzip to a custom folder.
            If _
                FileSystemObject.GetExtensionName(Destination) = CabExtensionName Or _
                FileSystemObject.GetExtensionName(Destination) = TarExtensionName Or _
                FileSystemObject.GetExtensionName(Destination) = TgzExtensionName Or _
                FileSystemObject.GetExtensionName(Destination) = ZipExtensionName Then
                ' Do not unzip to a folder named *.cab, *.tar, or *.zip.
                ' Strip extension.
                Destination = FileSystemObject.BuildPath( _
                    FileSystemObject.GetParentFolderName(Destination), _
                    FileSystemObject.GetBaseName(Destination))
            End If
        Else
            ' Unzip to a subfolder of the folder of the zipfile.
            Destination = FileSystemObject.BuildPath(ZipPath, ZipName)
        End If
            
        If FileSystemObject.FolderExists(Destination) And OverWrite = True Then
            ' Delete the existing folder.
            FileSystemObject.DeleteFolder Destination, True
        End If
        If Not FileSystemObject.FolderExists(Destination) Then
            ' Create the destination folder.
            FileSystemObject.CreateFolder Destination
        End If
        
        If Not FileSystemObject.FolderExists(Destination) Then
            ' For some reason the destination folder does not exist and cannot be created.
            ' Exit.
            Destination = ""
        Else
            ' Destination folder existed or has been created successfully.
            ' Resolve relative paths.
            Destination = FileSystemObject.GetAbsolutePathName(Destination)
            Path = FileSystemObject.GetAbsolutePathName(Path)
            ' Check file extension.
            If FileSystemObject.GetExtensionName(Path) = ZipExtensionName Then
                ' File extension is OK.
                ZipTemp = Path
            Else
                ' Rename the zip file by adding a zip extension.
                ZipTemp = Path & ZipExtension
                FileSystemObject.MoveFile Path, ZipTemp
            End If
            ' Unzip files and folders from the zip file to the destination folder.
            If OverWrite Then
                Options = DoOverwrite Or YesToAll
            Else
                Options = NoOverwrite Or YesToAll
            End If
            ShellApplication.Namespace(CVar(Destination)).CopyHere ShellApplication.Namespace(CVar(ZipTemp)).Items, Options
            If ZipTemp <> Path Then
                ' Remove the zip extension to restore the original file name.
                FileSystemObject.MoveFile ZipTemp, Path
            End If
        End If
    End If
    
    Set ShellApplication = Nothing
    Set FileSystemObject = Nothing
    
    If Err.Number <> ErrorNone Then
        Destination = ""
        Result = Err.Number
    ElseIf Destination = "" Then
        Result = ErrorOther
    End If
    
    UnZip = Result
     
End Function

' Zip a file or a folder to a zip file/folder using Windows Explorer.
' Default behaviour is similar to right-clicking a file/folder and selecting:
'   Send to zip file.
'
' Parameters:
'   Path:
'       Valid (UNC) path to the file or folder to zip.
'   Destination:
'       (Optional) Valid (UNC) path to file with zip extension or other extension.
'   Overwrite:
'       (Optional) Leave (default) or overwrite an existing zip file.
'       If False, the created zip file will be versioned: Example.zip, Example (2).zip, etc.
'       If True, an existing zip file will first be deleted, then recreated.
'
'   Path and Destination can be relative paths. If so, the current path is used.
'
'   If success, 0 is returned, and Destination holds the full path of the created zip file.
'   If error, error code is returned, and Destination will be zero length string.
'
' Early binding requires references to:
'
'   Shell:
'       Microsoft Shell Controls And Automation
'
'   Scripting.FileSystemObject:
'       Microsoft Scripting Runtime
'
' 2022-04-20. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function Zip( _
    ByVal Path As String, _
    Optional ByRef Destination As String, _
    Optional ByVal OverWrite As Boolean) _
    As Long
    
#If EarlyBinding Then
    ' Microsoft Scripting Runtime.
    Dim FileSystemObject    As Scripting.FileSystemObject
    ' Microsoft Shell Controls And Automation.
    Dim ShellApplication    As Shell
    
    Set FileSystemObject = New Scripting.FileSystemObject
    Set ShellApplication = New Shell
#Else
    Dim FileSystemObject    As Object
    Dim ShellApplication    As Object

    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
    Set ShellApplication = CreateObject("Shell.Application")
#End If
    
    ' Mandatory extension of zip file.
    Const ZipExtensionName  As String = "zip"
    Const ZipExtension      As String = "." & ZipExtensionName
    ' Native error numbers
    Const ErrorFileNotFound As Long = 53
    Const ErrorFileExists   As Long = 58
    Const ErrorNoPermission As Long = 70
    Const ErrorPathFile     As Long = 75
    ' Custom error numbers.
    Const ErrorOther        As Long = -1
    Const ErrorNone         As Long = 0
    ' Maximum (arbitrary) allowed count of created zip versions.
    Const MaxZipVersion     As Integer = 1000
    
    Dim IsRemovableDrive    As Boolean
    Dim Counter             As Long
    Dim Extension           As String
    Dim ExtensionName       As String
    Dim ZipHeader           As String
    Dim ZipPath             As String
    Dim ZipName             As String
    Dim ZipFile             As String
    Dim ZipBase             As String
    Dim ZipTemp             As String
    Dim Version             As Integer
    Dim Result              As Long
    
    If FileSystemObject.FileExists(Path) Then
        ' The source is an existing file.
        ZipName = FileSystemObject.GetBaseName(Path) & ZipExtension
        ZipPath = FileSystemObject.GetFile(Path).ParentFolder
    ElseIf FileSystemObject.FolderExists(Path) Then
        ' The source is an existing folder or fileshare.
        ZipBase = FileSystemObject.GetBaseName(Path)
        If ZipBase <> "" Then
            ' Path is a folder.
            ZipName = ZipBase & ZipExtension
            ZipPath = FileSystemObject.GetFolder(Path).ParentFolder
        Else
            ' Path is a fileshare, thus has no parent folder.
        End If
    Else
        ' The source does not exist.
    End If
       
    If ZipName = "" Then
        ' Nothing to zip. Exit.
        Destination = ""
    Else
        If Destination <> "" Then
            If FileSystemObject.GetExtensionName(Destination) = "" Then
                If FileSystemObject.FolderExists(Destination) Then
                    ZipPath = Destination
                Else
                    ' No folder for the zip file. Exit.
                    Destination = ""
                End If
            Else
                ' Destination is a file.
                ZipName = FileSystemObject.GetFileName(Destination)
                ZipPath = FileSystemObject.GetParentFolderName(Destination)
                If ZipPath = "" Then
                    ' No target folder specified. Use the folder of the source.
                    ZipPath = FileSystemObject.GetParentFolderName(Path)
                End If
            End If
        Else
            ' Use the already found folder of the source.
            Destination = ZipPath
        End If
    End If
        
    If Destination <> "" Then
        ZipFile = FileSystemObject.BuildPath(ZipPath, ZipName)

        If FileSystemObject.FileExists(ZipFile) Then
            If OverWrite = True Then
                ' Delete an existing file.
                FileSystemObject.DeleteFile ZipFile, True
                ' At this point either the file is deleted or an error is raised.
            Else
                ZipBase = FileSystemObject.GetBaseName(ZipFile)
                ExtensionName = FileSystemObject.GetExtensionName(ZipFile)
                Extension = "." & ExtensionName
                
                ' Modify name of the zip file to be created to preserve an existing file:
                '   "Example.zip" -> "Example (2).zip", etc.
                Version = Version + 1
                Do
                    Version = Version + 1
                    ZipFile = FileSystemObject.BuildPath(ZipPath, ZipBase & Format(Version, " \(0\)") & Extension)
                Loop Until FileSystemObject.FileExists(ZipFile) = False Or Version > MaxZipVersion
                If Version > MaxZipVersion Then
                    ' Give up.
                    Err.Raise ErrorPathFile, "Zip Create", "File could not be created."
                End If
            End If
        End If
        ' Set returned file name.
        Destination = ZipFile
    
        IsRemovableDrive = (FileSystemObject.GetDrive(FileSystemObject.GetDriveName(ZipPath)).DriveType = Removable)
        If Not IsRemovableDrive Then
            ' Check that the file/folder doesn't live on a linked drive.
            IsRemovableDrive = IsFolderAlias(ZipPath)
        End If
        
        ' Create a temporary zip name to allow for a final destination file with another extension than zip.
        If IsRemovableDrive Then
            ZipTemp = FileSystemObject.BuildPath(FileSystemObject.GetSpecialFolder(TemporaryFolder), FileSystemObject.GetBaseName(FileSystemObject.GetTempName()) & ZipExtension)
        Else
            ZipTemp = FileSystemObject.BuildPath(ZipPath, FileSystemObject.GetBaseName(FileSystemObject.GetTempName()) & ZipExtension)
        End If
        ' Create empty zip folder.
        ' Header string provided by Stuart McLachlan <stuart@lexacorp.com.pg>.
        ZipHeader = Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, vbNullChar)
        With FileSystemObject.OpenTextFile(ZipTemp, ForWriting, True)
            .Write ZipHeader
            .Close
        End With
        
        ' Resolve relative paths.
        ZipTemp = FileSystemObject.GetAbsolutePathName(ZipTemp)
        Path = FileSystemObject.GetAbsolutePathName(Path)
        ' Copy the source file/folder into the zip file.
        With ShellApplication
            Debug.Print Timer, "Zipping started . ";
            DoEvents
            .Namespace(CVar(ZipTemp)).CopyHere CVar(Path)
            DoEvents
            ' Ignore error while looking up the zipped file before is has been added.
            On Error Resume Next
            
            ' Wait for the file to created.
            Do Until .Namespace(CVar(ZipTemp)).Items.Count = 1 Or Counter = 10
                ' Wait a little ...
                Sleep 50
                Debug.Print ".";
                DoEvents
                Counter = Counter + 1
            Loop
            Debug.Print Counter
            
            ' Resume normal error handling.
            On Error GoTo 0
            Debug.Print Timer, "Zipping finished."
        End With

        ' Copy (Rename) the temporary zip to its final name.
        On Error Resume Next
        Do
            DoEvents
            FileSystemObject.MoveFile ZipTemp, ZipFile
            Debug.Print Str(Err.Number);
            Sleep 50
            Select Case Err.Number
                Case ErrorFileExists, ErrorNoPermission
                    ' Continue.
                Case Else
                    ' Unexpected error.
                    Exit Do
            End Select
            ' Expected error; file has been moved.
        Loop Until Err.Number = ErrorFileNotFound
        On Error GoTo 0
        Debug.Print
        Debug.Print Timer, "Moving finished."
    End If
    
    Set ShellApplication = Nothing
    Set FileSystemObject = Nothing
    
    If Err.Number <> ErrorNone Then
        Destination = ""
        Result = Err.Number
    ElseIf Destination = "" Then
        Result = ErrorOther
    End If
    
    Zip = Result

End Function

