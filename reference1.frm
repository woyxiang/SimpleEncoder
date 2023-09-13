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
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "reference1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ��������Windows API�еĽṹ�壬���ڴ洢����������Ϣ�ͽ�����Ϣ
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

' ����Windows API����CreateProcess�����ڴ���һ���µĽ��̡�
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Any, ByVal lpThreadAttributes As Any, ByVal bInheritHandles As Boolean, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Boolean

' ����Windows API����ReadFile�����ڶ�ȡ�ļ����߽��������
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesReaded As Long, ByVal lpOverlapped As Any) As Boolean

' ����Windows API����CloseHandle�����ڹر�һ���򿪵ľ����
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Boolean

Private Sub Command1_Click()
   ' ���������������ֱ����ڴ洢����������Ϣ�ͽ�����Ϣ��
   Dim siStartInfo  As STARTUPINFO
   Dim piProcInfo As PROCESS_INFORMATION
   
   ' �����˽���������Ϣ��һЩ������
   siStartInfo.cb = Len(siStartInfo)
   siStartInfo.dwFlags = &H100 Or &H1
   
   ' ����CreateProcess��������һ���µĽ��̡������"ffmpeg -i input.mp4 output.avi"����Ҫִ�е��������Ҫ����ʵ����������޸ġ�
   CreateProcess vbNullString, "ffmpeg -i input.mp4 output.avi", ByVal 0&, ByVal 0&, True, 0&, ByVal 0&, vbNullString, siStartInfo, piProcInfo

   ' ������һ���ڴ����ڴ洢�����Ȼ�����ReadFile������ȡ���̵������
   Dim strOutput As String
   strOutput = Space$(1024)
   Dim lngBytesReaded As Long
   ReadFile siStartInfo.hStdOutput, ByVal strOutput, Len(strOutput), lngBytesReaded, ByVal 0&

   ' ��ʾ��ȡ�������������Խ���һ���޸�Ϊ���Լ��Ĵ��룬�������ʾ�������������ϡ�
   MsgBox Left$(strOutput, lngBytesReaded)

   ' �رս��̺��̵߳ľ����
   CloseHandle piProcInfo.hProcess
   CloseHandle piProcInfo.hThread
End Sub


