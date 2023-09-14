VERSION 5.00
Begin VB.Form reference2 
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
Attribute VB_Name = "reference2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 定义Windows API函数和结构体
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Any, ByVal lpThreadAttributes As Any, ByVal bInheritHandles As Boolean, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Boolean
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesReaded As Long, ByVal lpOverlapped As Any) As Boolean
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Boolean

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

' 定义全局变量用于存储进程信息和输出句柄。
Dim piProcInfo As PROCESS_INFORMATION
Dim hOutput As Long

Private Sub Form_Load()
   ' 初始化STARTUPINFO结构体。
   Dim siStartInfo  As STARTUPINFO
   siStartInfo.cb = Len(siStartInfo)
   siStartInfo.dwFlags = &H100 Or &H1
   
   ' 创建新的进程并获取其输出句柄。
   CreateProcess vbNullString, "ffmpeg -i input.mp4 output.avi", ByVal 0&, ByVal 0&, True, 0&, ByVal 0&, vbNullString, siStartInfo, piProcInfo

   ' 存储输出句柄。
   hOutput = siStartInfo.hStdOutput
   
   ' 启动定时器，每秒读取一次输出。
   Timer1.Interval = 1000
   Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
   ' 定期读取输出，并显示在TextBox1上。
   Dim strOutput As String
   Dim lngBytesReaded As Long
   
   strOutput = Space$(1024)
   ReadFile hOutput, ByVal strOutput, Len(strOutput), lngBytesReaded, ByVal 0&
   
   TextBox1.Text = TextBox1.Text & Left$(strOutput, lngBytesReaded)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' 关闭句柄。
   CloseHandle piProcInfo.hProcess
   CloseHandle piProcInfo.hThread
End Sub

