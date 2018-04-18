Attribute VB_Name = "modFHand"


    'Type declaration for ChunkSize variable
    Type ChunkSize
        S12000 As String * 12000
        S6000 As String * 6000
        S3000 As String * 3000
        S1500 As String * 1500
        S500 As String * 500
        S100 As String * 100
        S25 As String * 25
        S5 As String * 5
        S1 As String * 1
    End Type
    
    'Declare the variable Bytes as of ChunkSize type
    Dim Bytes As ChunkSize
    Dim ExtMask As String 'This cotains the 0-s for the file extension formating
    Public CancelJob As Boolean
    Public CancelAndExit As Boolean
    Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Function SplitFile(FileName As String, FragmentSize As Long, DeleteFile As Boolean, Optional NumOfFragments As Integer) As Integer

    Dim SourceBytes As Long
    Dim SourceFile As String
    Dim DestinationFile As String
    Dim FragmentNumber As Integer
    Dim BytesDone As Long
    Dim FPath As String
    Dim FName As String
    Dim FNameNoExt As String
    Dim ErrorCode As Integer
    Dim dFileNoMask As String
    On Error GoTo ErrorHandler
    
    'Make sure the file exists
    If FileName = "" Or Dir(FileName) = "" Then
        ErrorCode = 1
        GoTo ErrorHandler
    End If
    
    'Ensure that the Fragment size is valid
    If FragmentSize = 0 Then
        ErrorCode = 2
        GoTo ErrorHandler
    End If
    
    'Retrieve the path name where file exists
    Do
        i = i + 1
        'Find the first occurance of the "\" in the FileName string from the right
        j = InStr(Len(FileName) - i, FileName, "\", vbTextCompare)
    Loop Until j > 0
    
    'Extract the file name
    FName = Right$(FileName, Len(FileName) - j)
    
    'Extract the path name
    FPath = Left$(FileName, j)
    
    'Find the file name without extension
    'Find the first occurance of the '.' in the FileName string
    j = InStr(1, FName, ".", vbTextCompare)
    If j = 0 Then   'File name does not contain a '.' character
        FNameNoExt = FName
    Else    'File name does contain the '.' character
        FNameNoExt = Left$(FName, j - 1)
    End If
    
    'Get total number or bytes in the source file
    SourceBytes = FileLen(FileName)
    'getting extension mask format
    Dim m
    m = Fix(SourceBytes / FragmentSize)
    If m < SourceBytes / FragmentSize Then m = m + 1
    ExtMask = String$(Len(Str$(m)) - 1, vbKey0)
    dFileNoMask = FPath & FName 'File name for deleting fragment files
    Debug.Print dFileNoMask
    'Open the source file for binary read
    Open FileName For Binary Access Read As #1 Len = 1
    
    Do
    'Clean up everything if canceled
    If CancelJob Then DeleteFragmentFiles ExtMask, dFileNoMask: frmMain.spProg (0): SplitFile = 2: Exit Do
        'Increase the number of Fragments counter by 1
        FragmentNumber = FragmentNumber + 1
        
        'Compose the file name of the new file to be created (file Fragment)
        DestinationFile = FPath & FName & "." & CStr(Format(FragmentNumber, ExtMask))
                    
        'Create the new file Fragment and open it for binary write
        Open DestinationFile For Binary Access Write As #2 Len = 1
        
        'Check whether the remaining bytes to process in the source file are
        'less than Fragment bytes
        If SourceBytes - BytesDone < FragmentSize Then
            RemainingBytes = SourceBytes - BytesDone
        Else
            RemainingBytes = FragmentSize
        End If
       
       'Read bytes from the source file and write them to the destination file (the current Fragment file)
       'Depending on the remaining bytes to read and write, the routine below will read the largest possible
       'chunk of data
       Do
            
            Select Case RemainingBytes
                Case Is >= 12000
                    'Read 12000 bytes of data from the source file
                    Get #1, , Bytes.S12000
                    'Write the bytes to the destination file
                    Put #2, , Bytes.S12000
                    'Decrease the number of remaining bytes by 12000
                    RemainingBytes = RemainingBytes - 12000
                    'Update the bytes done counter
                    BytesDone = BytesDone + 12000
                    'Yield to windows and other processes to do their jobs
                    'Also, this helps fulshing the disk buffers to the file
                    DoEvents
                Case 6000 To 11999
                    Get #1, , Bytes.S6000
                    Put #2, , Bytes.S6000
                    RemainingBytes = RemainingBytes - 6000
                    BytesDone = BytesDone + 6000
                    DoEvents
                Case 3000 To 5999
                    Get #1, , Bytes.S3000
                    Put #2, , Bytes.S3000
                    RemainingBytes = RemainingBytes - 3000
                    BytesDone = BytesDone + 3000
                    DoEvents
                Case 1500 To 2999
                    Get #1, , Bytes.S1500
                    Put #2, , Bytes.S1500
                    RemainingBytes = RemainingBytes - 1500
                    BytesDone = BytesDone + 1500
                    DoEvents
                Case 500 To 1499
                    Get #1, , Bytes.S500
                    Put #2, , Bytes.S500
                    RemainingBytes = RemainingBytes - 500
                    BytesDone = BytesDone + 500
                    DoEvents
                Case 100 To 499
                    Get #1, , Bytes.S100
                    Put #2, , Bytes.S100
                    RemainingBytes = RemainingBytes - 100
                    BytesDone = BytesDone + 100
                    DoEvents
                Case 25 To 99
                    Get #1, , Bytes.S25
                    Put #2, , Bytes.S25
                    RemainingBytes = RemainingBytes - 25
                    BytesDone = BytesDone + 25
                    DoEvents
                Case 5 To 24
                    Get #1, , Bytes.S5
                    Put #2, , Bytes.S5
                    RemainingBytes = RemainingBytes - 5
                    BytesDone = BytesDone + 5
                    DoEvents
                Case 1 To 4
                    Get #1, , Bytes.S1
                    Put #2, , Bytes.S1
                    RemainingBytes = RemainingBytes - 1
                    BytesDone = BytesDone + 1
                    DoEvents
                Case Is = 0
                    'When the loop enters here, the Fragment bytes are completed.
                    'Close the Fragment file and exit the loop
                    Close 2
                    DoEvents
                    Exit Do
            End Select
            
            'Update the percent control on the form
            frmMain.spProg (Int((BytesDone / SourceBytes) * 100))
            'Refresh the form and yield to windows
            DoEvents
        Loop
        
    Loop Until BytesDone = SourceBytes
    'Close the source file
    Close 1
    
    'Delete the source file if necessary
    If DeleteFile = True And CancelJob = False Then
    Kill FileName
    End If
    
    NumOfFragments = FragmentNumber
    If CancelJob = False Then SplitFile = 0
    frmMain.spProg (0)
    Exit Function
    
ErrorHandler:
SplitFile = 1
Exit Function
End Function

Function MergeFiles(SourceFile As String, DeleteFile As Boolean, Optional NumOfFragments As Integer) As Integer

    Dim TotalBytes As Long
    Dim DestinationFile As String
    Dim FragmentFile As String
    Dim FragmentNumber As Integer
    Dim Fragments As Integer
    Dim BytesDone As Long
    Dim FPath As String
    Dim FName As String
    Dim FNameNoExt As String
    Dim ErrorCode As Integer
    Dim dFileNoMask As String
    On Error GoTo ErrorHandler
    
    'Make sure the source file name is given and is valid (exists)
    If SourceFile = "" Or Dir(SourceFile) = "" Then
        ErrorCode = 1
        GoTo ErrorHandler
    End If
    
    'Find the number of Fragments of the split file
    'Retrieve the path name where files exist
    Do
        i = i + 1
        'Find the first occurance of the "\" in the SourceFile string from the right
        j = InStr(Len(SourceFile) - i, SourceFile, "\", vbTextCompare)
    Loop Until j > 0
    
    'Extract the file name
    FName = Right$(SourceFile, Len(SourceFile) - j)
    
    'Extract the path name
    FPath = Left$(SourceFile, j)
    
    'Find the file name without extension
    'Find the first occurance of the '.' in the SourceFile string
    j = InStr(1, StrReverse(FName), ".", vbTextCompare)
    If j = 0 Then   'File name does not contain a '.' character
        FNameNoExt = FName
    Else    'File name does contain the '.' character
        FNameNoExt = Left$(FName, Len(FName) - j)
    End If
    
    
    'getting extension mask format
    ExtMask = String$(InStr(1, StrReverse(SourceFile), ".") - 1, vbKey0)
     
    'Now find the number of Fragments of the split file that reside in
    'the same directory where the source file is
    'Also count the total number of bytes in the Fragments (this will be
    'used for the calculation of the percent done value
    dFileNoMask = FPath & FNameNoExt
    Do
        'Increase the number of Fragments counter by 1
        Fragments = Fragments + 1
        
        'Compose the Fragment file name and check
        FragmentFile = FPath & FNameNoExt & "." & CStr(Format(Fragments, ExtMask))
        If Dir(FragmentFile) = "" Then Exit Do
        TotalBytes = TotalBytes + FileLen(FragmentFile)
    Loop
    
    Fragments = Fragments - 1 'This is the number of Fragments found
    Debug.Print Fragments
    'Check the detected number of Fragments. If is =0, then the given
    'file name is not a Fragment file
    If Fragments = 0 Then
        ErrorCode = 2
        GoTo ErrorHandler
    End If
    
    'Check if the destination file to be created does exist in the same dir
    'If yes, return error in the function return value
    DestinationFile = FPath & FNameNoExt
    If Dir(DestinationFile) <> "" Then
        ErrorCode = 3
        GoTo ErrorHandler
    End If

    'Open the destination file for binary write
    Open DestinationFile For Binary Access Write As #1 Len = 1
    
    Do
        'Clean up after canceling
        If CancelJob Then Close 1: Kill DestinationFile: frmMain.spProg (0): MergeFiles = 2:  Exit Do
        
        'Increase the number of Fragments counter by 1
        FragmentNumber = FragmentNumber + 1
        
        'Compose the file name of the new Fragment file to be opened and read
        SourceFile = FPath & FNameNoExt & "." & CStr(Format(FragmentNumber, ExtMask))
        
        'Open the source file Fragment for binary read
        Open SourceFile For Binary Access Read As #2 Len = 1
        'Get the total number of bytes in the current Fragment file
        RemainingBytes = FileLen(SourceFile)
       
       'Read bytes from the source file (the current Fragment file) and write them to the destination file
       'Depending on the remaining bytes to read and write, the routine below will read the largest possible
       'chunk of data
       Do
            
            Select Case RemainingBytes
                Case Is >= 12000
                    'Read 12000 bytes of data from the source file
                    Get #2, , Bytes.S12000
                    'Write the bytes to the destination file
                    Put #1, , Bytes.S12000
                    'Decrease the number of remaining bytes by 12000
                    RemainingBytes = RemainingBytes - 12000
                    'Update the bytes done counter
                    BytesDone = BytesDone + 12000
                    'Yield to windows and other processes to do their jobs
                    'Also, this helps fulshing the disk buffers to the file
                    DoEvents
                Case 6000 To 11999
                    Get #2, , Bytes.S6000
                    Put #1, , Bytes.S6000
                    RemainingBytes = RemainingBytes - 6000
                    BytesDone = BytesDone + 6000
                    DoEvents
                Case 3000 To 5999
                    Get #2, , Bytes.S3000
                    Put #1, , Bytes.S3000
                    RemainingBytes = RemainingBytes - 3000
                    BytesDone = BytesDone + 3000
                    DoEvents
                Case 1500 To 2999
                    Get #2, , Bytes.S1500
                    Put #1, , Bytes.S1500
                    RemainingBytes = RemainingBytes - 1500
                    BytesDone = BytesDone + 1500
                    DoEvents
                Case 500 To 1499
                    Get #2, , Bytes.S500
                    Put #1, , Bytes.S500
                    RemainingBytes = RemainingBytes - 500
                    BytesDone = BytesDone + 500
                    DoEvents
                Case 100 To 499
                    Get #2, , Bytes.S100
                    Put #1, , Bytes.S100
                    RemainingBytes = RemainingBytes - 100
                    BytesDone = BytesDone + 100
                    DoEvents
                Case 25 To 99
                    Get #2, , Bytes.S25
                    Put #1, , Bytes.S25
                    RemainingBytes = RemainingBytes - 25
                    BytesDone = BytesDone + 25
                    DoEvents
                Case 5 To 24
                    Get #2, , Bytes.S5
                    Put #1, , Bytes.S5
                    RemainingBytes = RemainingBytes - 5
                    BytesDone = BytesDone + 5
                    DoEvents
                Case 1 To 4
                    Get #2, , Bytes.S1
                    Put #1, , Bytes.S1
                    RemainingBytes = RemainingBytes - 1
                    BytesDone = BytesDone + 1
                    DoEvents
                Case Is = 0
                    'When the loop enters here, the Fragment bytes are completed.
                    'Close the Fragment file and exit the loop
                    Close 2
                    DoEvents
                    Exit Do
            End Select
            
            'Update the percent control on the form
            frmMain.spProg (Int((BytesDone / TotalBytes) * 100))
            'Refresh the form and yield to windows
            DoEvents
        Loop
        
    Loop Until FragmentNumber = Fragments
    'Close the destination file
    Close 1
    If CancelAndExit Then End
        'Delete the fragment files if necessary
        If DeleteFile = True And CancelJob = False Then
        DeleteFragmentFiles ExtMask, dFileNoMask
        End If
    
    NumOfFragments = Fragments
    If CancelJob = False Then MergeFiles = 0
    frmMain.spProg (0)
    Exit Function
    
ErrorHandler:
MergeFiles = 1
Exit Function
End Function
    
Sub DeleteFragmentFiles(eMask As String, sFileName As String)

Dim fFile As String
Dim fNum As Integer
Do
    fNum = fNum + 1
    fFile = sFileName & "." & CStr(Format(fNum, eMask))
    If Dir(fFile) = "" Then Exit Do
    Kill fFile
    Debug.Print fFile
Loop
If CancelAndExit Then End 'make sure the program closes properly after cancel
End Sub
