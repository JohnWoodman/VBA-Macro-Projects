Private Declare PtrSafe Function CreateFile Lib "Kernel32" _
 Alias "CreateFileA" (ByVal lpFileName As String, _
 ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
 ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, _
 ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As LongPtr

Private Declare PtrSafe Function CloseHandle Lib "Kernel32" (ByVal hObject As LongPtr) As Long

Const MAX_PATH = 260
Public Enum MINIDUMP_TYPE
    MiniDumpNormal = 0
    MiniDumpWithDataSegs = 1
    MiniDumpWithFullMemory = 2
    MiniDumpWithHandleData = 4
    MiniDumpFilterMemory = 8
    MiniDumpScanMemory = 10
    MiniDumpWithUnloadedModules = 20
    MiniDumpWithIndirectlyReferencedMemory = 40
    MiniDumpFilterModulePaths = 80
    MiniDumpWithProcessThreadData = 100
    MiniDumpWithPrivateReadWriteMemory = 200
    MiniDumpWithoutOptionalData = 400
    MiniDumpWithFullMemoryInfo = 800
    MiniDumpWithThreadInfo = 1000
    MiniDumpWithCodeSegs = 2000
End Enum

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32DefaultHeapIDB As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    pcPriClassBaseB As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Private Declare PtrSafe Function MiniDumpWriteDump Lib "dbghelp" ( _
 ByVal hProcess As LongPtr, _
 ByVal ProcessId As Long, _
 ByVal hFile As LongPtr, _
 ByVal DumpType As MINIDUMP_TYPE, _
 ByVal ExceptionParam As LongPtr, _
 ByVal UserStreamParam As LongPtr, _
 ByVal CallackParam As LongPtr) As Boolean
 
Private Declare PtrSafe Function GetCurrProcess Lib "Kernel32" _
 Alias "GetCurrentProcess" () As LongPtr
 
Private Declare PtrSafe Function GetCurrProcessID Lib "Kernel32" _
 Alias "GetCurrentProcessId" () As Long
 
Private Declare PtrSafe Function Create32Snapshot Lib "Kernel32" _
 Alias "CreateToolhelp32Snapshot" ( _
 ByVal dwFlags As Long, _
 ByVal th32ProcessID As Long) As Long
 
Private Declare PtrSafe Function Process32First Lib "Kernel32" (ByVal hSnapShot As LongPtr, uProcess As PROCESSENTRY32) As Long
Private Declare PtrSafe Function Process32Next Lib "Kernel32" (ByVal hSnapShot As LongPtr, uProcess As PROCESSENTRY32) As Long
Private Declare PtrSafe Function OpenProcess Lib "kernel32.dll" ( _
        ByVal dwAccess As Long, _
        ByVal fInherit As Integer, _
        ByVal hObject As Long _
    ) As LongPtr
Private Declare PtrSafe Function IsUserAnAdmin Lib "shell32" () As Long
Private Declare PtrSafe Function GetLastError Lib "kernel32.dll" () As Long

 
 

Private Sub Document_Open()
    
    Const GENERIC_WRITE = &H40000000
    Const GENERIC_READ = &H80000000
    Const FILE_ATTRIBUTE_NORMAL = &H80
    Const CREATE_ALWAYS = 2
    Const OPEN_ALWAYS = 4
    Const INVALID_HANDLE_VALUE = -1
    
    procPID = 0
    Dim procHandle As LongPtr
    Dim outFile As LongPtr
    Dim fileToDump As String
    
    'Location of file to dump to
    fileToDump = "c:\Users\john\Desktop\process.dmp"
    outFile = CreateFile(fileToDump, GENERIC_WRITE Or GENERIC_READ, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    
    If outFile = INVALID_HANDLE_VALUE Then
        MsgBox ("Error Opening File")
        CloseHandle (outFile)
    Else
        'MsgBox ("Opened File!")
    End If

    Const TH32CS_SNAPPROCESS = &H2
    Dim snapshot As Long
    snapshot = Create32Snapshot(TH32CS_SNAPPROCESS, ByVal 0&)
    
    Dim uProcess As PROCESSENTRY32
    uProcess.dwSize = Len(uProcess)
    
    Dim ProcessFound As Boolean
    ProcessFound = Process32First(snapshot, uProcess)
    
    Dim ProcID As Long
    
    Dim ProcName As String
    'Process name to dump
    ProcName = "notepad.exe"
    Do
        If Left$(uProcess.szExeFile, Len(ProcName)) = LCase$(ProcName) Then
            ProcID = uProcess.th32ProcessID
            'MsgBox ("Process Found: " & ProcID)
            ProcessFound = False
        Else
            ProcessFound = Process32Next(snapshot, uProcess)
        End If
    Loop While ProcessFound
    
    Dim hParent As LongPtr
    Const PROCESS_ALL_ACCESS = &H1F0FFF
    hParent = OpenProcess(PROCESS_ALL_ACCESS, 0, ProcID)
    CloseHandle (snapshot)
    
    Dim dumped As Boolean
    dumped = MiniDumpWriteDump(hParent, _
                      ProcID, _
                      outFile, _
                      MINIDUMP_TYPE.MiniDumpWithFullMemory, _
                      0, _
                      0, _
                      0)
                      
    If dumped Then
        'MsgBox ("True baby")
    Else
        MsgBox ("Nah Chief")
        MsgBox (Hex(Err.LastDllError))
    End If
    
    CloseHandle (outFile)
    CloseHandle (proc)
End Sub
