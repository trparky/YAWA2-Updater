Imports System.Runtime.InteropServices

Namespace NativeMethod
    Friend Class NativeMethods
        <DllImport("User32.dll", EntryPoint:="SetForegroundWindow")>
        Public Shared Function SetForegroundWindow(ByVal hWnd As Integer) As Int32
        End Function

        <DllImport("kernel32.dll")>
        Friend Shared Function OpenProcess(dwDesiredAccess As ProcessAccessFlags, bInheritHandle As Boolean, dwProcessId As Integer) As IntPtr
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)>
        Friend Shared Function CloseHandle(hHandle As IntPtr) As Boolean
        End Function

        <DllImport("kernel32.dll", CharSet:=CharSet.Unicode)>
        Friend Shared Function QueryFullProcessImageName(hprocess As IntPtr, dwFlags As Integer, lpExeName As Text.StringBuilder, ByRef size As Integer) As Boolean
        End Function
    End Class

    Module APIs
        <Flags>
        Public Enum ProcessAccessFlags As UInteger
            PROCESS_QUERY_LIMITED_INFORMATION = &H1000
            All = &H1F0FFF
            Terminate = &H1
            CreateThread = &H2
            VirtualMemoryOperation = &H8
            VirtualMemoryRead = &H10
            VirtualMemoryWrite = &H20
            DuplicateHandle = &H40
            CreateProcess = &H80
            SetQuota = &H100
            SetInformation = &H200
            QueryInformation = &H400
            QueryLimitedInformation = &H1000
            Synchronize = &H100000
        End Enum
    End Module
End Namespace