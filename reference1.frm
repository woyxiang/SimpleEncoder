VERSION 5.00
Begin VB.Form reference1 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "reference1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 定义两个Windows API中的结构体，用于存储进程启动信息和进程信息
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

' 声明Windows API函数CreateProcess，用于创建一个新的进程。
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Any, ByVal lpThreadAttributes As Any, ByVal bInheritHandles As Boolean, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Boolean

' 声明Windows API函数ReadFile，用于读取文件或者进程输出。
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesReaded As Long, ByVal lpOverlapped As Any) As Boolean

' 声明Windows API函数CloseHandle，用于关闭一个打开的句柄。
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Boolean

Private Sub Command1_Click()
   ' 定义两个变量，分别用于存储进程启动信息和进程信息。
   Dim siStartInfo  As STARTUPINFO
   Dim piProcInfo As PROCESS_INFORMATION
   
   ' 设置了进程启动信息的一些参数。
   siStartInfo.cb = Len(siStartInfo)
   siStartInfo.dwFlags = &H100 Or &H1
   
   ' 调用CreateProcess函数创建一个新的进程。这里的"ffmpeg -i input.mp4 output.avi"是你要执行的命令，你需要根据实际情况进行修改。
   CreateProcess vbNullString, "ffmpeg -i input.mp4 output.avi", ByVal 0&, ByVal 0&, True, 0&, ByVal 0&, vbNullString, siStartInfo, piProcInfo

   ' 分配了一段内存用于存储输出，然后调用ReadFile函数读取进程的输出。
   Dim strOutput As String
   strOutput = Space$(1024)
   Dim lngBytesReaded As Long
   ReadFile siStartInfo.hStdOutput, ByVal strOutput, Len(strOutput), lngBytesReaded, ByVal 0&

   ' 显示读取到的输出。你可以将这一行修改为你自己的代码，将输出显示在你的软件界面上。
   MsgBox Left$(strOutput, lngBytesReaded)

   ' 关闭进程和线程的句柄。
   CloseHandle piProcInfo.hProcess
   CloseHandle piProcInfo.hThread
End Sub


