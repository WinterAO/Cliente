Attribute VB_Name = "modMonitor"
Option Explicit

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetProcessMemoryInfo Lib "psapi" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long

Private Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

Private prevSystemTime As Currency
Private prevUserTime As Currency
Private prevKernelTime As Currency

Private Declare Sub GlobalMemoryStatusEx Lib "kernel32" (lpBuffer As MEMORYSTATUSEX)
Private Type MEMORYSTATUSEX
    dwLength As Long
    dwMemoryLoad As Long
    ullTotalPhys As Currency
    ullAvailPhys As Currency
    ullTotalPageFile As Currency
    ullAvailPageFile As Currency
    ullTotalVirtual As Currency
    ullAvailVirtual As Currency
    ullAvailExtendedVirtual As Currency
End Type

Private Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Sub InitializeCPUUsage()
'*********************************
'Autor: Lorwik
'Fecha: 09/07/2024
'*********************************

    Dim sysTime As Currency
    Dim userTime As Currency
    Dim kernelTime As Currency
    
    Call GetSystemTimes(sysTime, userTime, kernelTime)
    prevSystemTime = sysTime
    prevUserTime = userTime
    prevKernelTime = kernelTime
End Sub

Private Sub GetSystemTimes(ByRef sysTime As Currency, ByRef userTime As Currency, ByRef kernelTime As Currency)
'*********************************
'Autor: Lorwik
'Fecha: 09/07/2024
'*********************************

    Dim createTime As FILETIME, exitTime As FILETIME, kTime As FILETIME, uTime As FILETIME
    Call GetProcessTimes(GetCurrentProcess(), createTime, exitTime, kTime, uTime)
    
    sysTime = CDbl(kTime.dwLowDateTime) + 4294967296# * kTime.dwHighDateTime
    userTime = CDbl(uTime.dwLowDateTime) + 4294967296# * uTime.dwHighDateTime
    kernelTime = sysTime
End Sub

Public Function CalculateCPUUsage() As Double
'*********************************
'Autor: Lorwik
'Fecha: 09/07/2024
'*********************************

    Dim sysTime As Currency
    Dim userTime As Currency
    Dim kernelTime As Currency
    Dim cpuUsage As Double
    
    Call GetSystemTimes(sysTime, userTime, kernelTime)
    
    Dim sysTimeDiff As Currency
    Dim userTimeDiff As Currency
    Dim kernelTimeDiff As Currency
    
    sysTimeDiff = sysTime - prevSystemTime
    userTimeDiff = userTime - prevUserTime
    kernelTimeDiff = kernelTime - prevKernelTime
    
    If sysTimeDiff <> 0 Then
        cpuUsage = ((userTimeDiff + kernelTimeDiff) / sysTimeDiff) * 100
        ' Asegurarse de que el uso de CPU no supere el 100%
        If cpuUsage > 100 Then cpuUsage = 100
    End If
    
    prevSystemTime = sysTime
    prevUserTime = userTime
    prevKernelTime = kernelTime
    
    CalculateCPUUsage = cpuUsage
End Function

Private Function BytesToMB(ByVal bytes As Currency) As Double
'*********************************
'Autor: Lorwik
'Fecha: 09/07/2024
'*********************************

    BytesToMB = bytes / 1048576 ' 1 MB = 1024 * 1024 bytes
End Function

Private Function BytesToGB(ByVal bytes As Currency) As Double
'*********************************
'Autor: Lorwik
'Fecha: 09/07/2024
'*********************************

    BytesToGB = bytes / 1073741824 ' 1 GB = 1024 * 1024 * 1024 bytes
End Function

Private Function FormatMemory(ByVal bytes As Currency) As String
'*********************************
'Autor: Lorwik
'Fecha: 09/07/2024
'*********************************

    FormatMemory = Format$(BytesToMB(bytes), "#,##0") & " MB, " & Format$(BytesToGB(bytes), "0.0") & " GB"
End Function

Public Sub GetProcessMemoryUsage()
'*********************************
'Autor: Lorwik
'Fecha: 09/07/2024
'*********************************

    Dim hProcess As Long
    Dim pmc As PROCESS_MEMORY_COUNTERS
    
    hProcess = GetCurrentProcess()
    pmc.cb = Len(pmc)
    
    If GetProcessMemoryInfo(hProcess, pmc, pmc.cb) Then
        frmMonitor.lblWorkingSet.Caption = "Tamaño del Conjunto de Trabajo: " & pmc.WorkingSetSize & " bytes (" & FormatMemory(pmc.WorkingSetSize) & ")"
        frmMonitor.lblPagefileUsage.Caption = "Uso del Archivo de Paginación: " & pmc.PagefileUsage & " bytes (" & FormatMemory(pmc.PagefileUsage) & ")"
    Else
        frmMonitor.lblWorkingSet.Caption = "Error obteniendo la información de memoria del proceso."
    End If
End Sub

Public Sub GetMemoryStatus()
'*********************************
'Autor: Lorwik
'Fecha: 09/07/2024
'*********************************

    Dim memStatus As MEMORYSTATUSEX
    memStatus.dwLength = Len(memStatus)
    GlobalMemoryStatusEx memStatus
    frmMonitor.lblMemLoad.Caption = "Carga de Memoria: " & memStatus.dwMemoryLoad & "%"
    frmMonitor.lblTotalPhysMem.Caption = "Memoria Física Total: " & memStatus.ullTotalPhys & " bytes (" & FormatMemory(memStatus.ullTotalPhys) & ")"
    frmMonitor.lblAvailPhysMem.Caption = "Memoria Física Disponible: " & memStatus.ullAvailPhys & " bytes (" & FormatMemory(memStatus.ullAvailPhys) & ")"
End Sub

Public Sub GetCPUUsage()
'*********************************
'Autor: Lorwik
'Fecha: 09/07/2024
'*********************************

    frmMonitor.lblCPUUsage.Caption = "Uso de CPU: " & Format$(CalculateCPUUsage(), "0.0") & "%"
End Sub
