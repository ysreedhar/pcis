Attribute VB_Name = "MDialogHook"
  Option Explicit
 
 
  Public Const FCIDM_SHVIEW_LARGEICON As Long = &H7029&   ' 28713
  Public Const FCIDM_SHVIEW_SMALLICON As Long = &H702A&   ' 28714
  Public Const FCIDM_SHVIEW_LIST As Long = &H702B&        ' 28715
  Public Const FCIDM_SHVIEW_REPORT As Long = &H702C&      ' 28716
  Public Const FCIDM_SHVIEW_THUMBNAIL As Long = &H702D&   ' 28717
  Public Const FCIDM_SHVIEW_TILE As Long = &H702E&        ' 28718
  '***********************************************************************

  Public Const WM_COMMAND As Long = &H111&
  
  
  Public Enum OPENFILENAME_FLAGS
    OFN_ALLOWMULTISELECT = &H200
    OFN_CREATEPROMPT = &H2000
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_EXTENSIONDIFFERENT = &H400&
    OFN_FILEMUSTEXIST = &H1000
    OFN_HIDEREADONLY = &H4&
    OFN_NOCHANGEDIR = &H8&
    OFN_NOLONGNAMES = &H40000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NOVALIDATE = &H100
    OFN_OVERWRITEPROMPT = &H2&
    OFN_PATHMUSTEXIST = &H800
    OFN_READONLY = &H1
    OFN_SHAREAWARE = &H4000
    OFN_SHAREFALLTHROUGH = 2
    OFN_SHAREWARN = 0
    OFN_SHARENOWARN = 1
    OFN_SHOWHELP = &H10
    OFS_MAXPATHNAME = 128
    
    ' #if /* WINVER >= 0x0400 */
    OFN_EXPLORER = &H80000                     '// new look commdlg
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000                   '// force long names for 3.x modules
    OFN_ENABLEINCLUDENOTIFY = &H400000         '// send include message to callback
    OFN_ENABLESIZING = &H800000                '// enables the sizing for the dialog
    ' #endif /* WINVER >= 0x0400 */

    '#if (_WIN32_WINNT >= 0x0500)
    'OFN_USESHELLITEM = &H1000000               '// disabling support for IShellItem for now (see comdlg32\commdlg.h)
    OFN_DONTADDTORECENT = &H2000000
    OFN_FORCESHOWHIDDEN = &H10000000           '// Show All files including System and hidden files
    '#endif // (_WIN32_WINNT >= 0x0500)
    
    '//FlagsEx Values
    '#if (_WIN32_WINNT >= 0x0500)
    OFN_EX_NOPLACESBAR = &H1
    '#endif // (_WIN32_WINNT >= 0x0500)

  End Enum
  
  Public Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As OPENFILENAME_FLAGS
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String

    ' new members of this struct added in version 5 of the shell
    ' we can still use this struct with older versions of the shell
    ' because we pass the size of the struct expected by the function
    pvReserved As Long
    dwReserved As Long
    FlagsEx As Long
  End Type
  
  Public Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (lpOpenfilename As OPENFILENAME) As Long
  Public Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

  Public Declare Function CommDlgExtendedError Lib "comdlg32" () As Long


  Public Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
  End Type


  ' messages used in the hook proc
  Public Const WM_NOTIFY As Long = &H4E&

  Public Const CDN_FIRST As Long = (0& - 601&)
  Public Const CDN_LAST As Long = (0& - 699&)

  ' Notifications when Open or Save dialog status changes
  Public Const CDN_INITDONE As Long = (CDN_FIRST - &H0&)
  Public Const CDN_SELCHANGE As Long = (CDN_FIRST - &H1&)
  Public Const CDN_FOLDERCHANGE As Long = (CDN_FIRST - &H2&)
  Public Const CDN_SHAREVIOLATION As Long = (CDN_FIRST - &H3&)
  Public Const CDN_HELP As Long = (CDN_FIRST - &H4&)
  Public Const CDN_FILEOK As Long = (CDN_FIRST - &H5&)
  Public Const CDN_TYPECHANGE As Long = (CDN_FIRST - &H6&)
  Public Const CDN_INCLUDEITEM As Long = (CDN_FIRST - &H7&)

  Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent&, ByVal hWndChildAfter&, ByVal lpClassName$, ByVal lpWindowName$) As Long

  Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc&, ByVal hWnd&, ByVal Msg&, ByVal wParam&, ByVal lParam&) As Long


  Public Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename$, lpdwHandle&) As Long
  Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename$, ByVal dwHandle&, ByVal dwLen&, lpData As Any) As Long
  Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long


  Public Const WM_SETREDRAW As Long = &HB&
  Public Const WM_GETMINMAXINFO As Long = &H24&
  Public Const WM_WINDOWPOSCHANGING As Long = &H46&


  Public Const WM_RBUTTONUP As Long = &H205&

  Public Const MK_RBUTTON As Long = &H2&

  Public Const GWL_WNDPROC As Long = (-4&)

  Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)

  Public Declare Function GetDesktopWindow Lib "user32" () As Long

  Public Declare Function GetParent Lib "user32" (ByVal hWnd&) As Long

  Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, wParam As Any, lParam As Any) As Long
  Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd&, ByVal wMsg&, wParam As Any, lParam As Any) As Long
  Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd&, ByVal nIndex&, ByVal dwNewLong&) As Long

Public Function DialogHookProc(ByVal hDlg&, ByVal nMsg&, ByVal wParam&, ByVal lParam&) As Long
  ' this is the dialog hook proc.  it is called by the dialog to inform us when certian
  ' actions occur.
  
  ' the hDlg param in a hook proc is the handle of a sub dialog created to contain any
  ' controls that we might want to add to the parent dialog.  to get the actual handle
  ' to the dialog itself, we need to use the GetParent function.
  
  Dim hLV&, lpNMHDR As NMHDR
      
  Select Case nMsg
    ' the WM_NOTIFY message with a code of CDN_FOLDERCHANGE is sent when the
    ' folder view is changing before the dialog is displayed.
    Case WM_NOTIFY
        CopyMemory lpNMHDR, ByVal lParam, Len(lpNMHDR)
                    
        Select Case lpNMHDR.code
          '*******************************************************************************
          ' code that calles the undocumemted messages for changing the listview's view
          ' Thanks go to Brad Martinez for discovering these messages
          Case CDN_FOLDERCHANGE
            hLV = FindWindowEx(GetParent(hDlg), 0, "SHELLDLL_DefView", vbNullString)
            
            If hLV Then
              Call SendMessage(hLV, WM_COMMAND, ByVal FCIDM_SHVIEW_REPORT, ByVal 0&)
            End If
          '*******************************************************************************
            
        End Select
      
  End Select
    
End Function

Public Function ReturnProcAddress(ByVal lpProc&) As Long
  ' helper function to return the address of the hook proc
  ReturnProcAddress = lpProc
End Function

Public Function Is2KShell() As Boolean
  ' this function returns the version of the Comdlg32.dll on the system
  ' this info is used to determine which version of the OPENFILENAME struct
  ' should be passed to the dialog functions
  
  Dim nBuffSize&, nDiscard&, lpBuffer&, nVerMajor&, abytBuffer() As Byte
  
  Const FILE_NAME As String = "Comdlg32.dll"
  
  nBuffSize = GetFileVersionInfoSize(FILE_NAME, nDiscard)
  
  If nBuffSize > 0 Then
    ReDim abytBuffer(nBuffSize - 1) As Byte
    
    Call GetFileVersionInfo(FILE_NAME, 0&, nBuffSize, abytBuffer(0))
    
    If VerQueryValue(abytBuffer(0), "\", lpBuffer, nDiscard) Then
      CopyMemory nVerMajor, ByVal lpBuffer + 10, 2&
        
      If nVerMajor >= 5 Then Is2KShell = True
    End If
  End If
  
End Function
