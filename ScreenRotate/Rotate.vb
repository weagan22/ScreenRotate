
Imports System
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

#Region "DEVMODE_STRUCT ................................................................"
<StructLayout(LayoutKind.Sequential)> _
Public Structure POINTL
    Public x As Integer
    Public y As Integer
End Structure

<StructLayout(LayoutKind.Sequential)> _
Public Structure DEVMODE

    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Public dmDeviceName As String
    Public dmSpecVersion As Short
    Public dmDriverVersion As Short
    Public dmSize As Short
    Public dmDriverExtra As Short
    Public dmFields As Integer

    ' display only fields
    Public dmPosition As POINTL
    Public dmDisplayOrientation As Integer
    Public DisplayFixedOutput As Integer

    Public dmColor As Short
    Public dmDuplex As Short
    Public dmYResolution As Short
    Public dmTTOption As Short
    Public dmCollate As Short
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Public dmFormName As String
    Public dmLogPixels As Short
    Public dmBitsPerPel As Integer
    Public dmPelsWidth As Integer
    Public dmPelsHeight As Integer

    Public dmDisplayFlags As Integer
    Public dmDisplayFrequency As Integer

    Public dmICMMethod As Integer
    Public dmICMIntent As Integer
    Public dmMediaType As Integer
    Public dmDitherType As Integer
    Public dmReserved1 As Integer
    Public dmReserved2 As Integer

    Public dmPanningWidth As Integer
    Public dmPanningHeight As Integer

End Structure

#End Region

#Region "PINVOKEDEF ....................................................................."
Class User32
    <DllImport("user32.dll")> _
    Public Shared Function EnumDisplaySettings(ByVal deviceName As String, ByVal modeNum As Integer, ByRef devMode As DEVMODE) As Integer
    End Function

    <DllImport("user32.dll")> _
    Public Shared Function ChangeDisplaySettings(ByRef devMode As DEVMODE, ByVal flags As Integer) As Integer
    End Function

    Public Const ENUM_CURRENT_SETTINGS As Integer = -1
    Public Const CDS_UPDATEREGISTRY As Integer = &H1
    Public Const CDS_TEST As Integer = &H2
    Public Const DISP_CHANGE_SUCCESSFUL As Integer = 0
    Public Const DISP_CHANGE_RESTART As Integer = 1
    Public Const DISP_CHANGE_FAILED As Integer = -1

    Public Const DM_DISPLAYORIENTATION As Integer = &H80
    Public Const DM_PELSWIDTH As Integer = &H80000
    Public Const DM_PELSHEIGHT As Integer = &H100000
End Class
#End Region

Public Class Rotate

    Public Enum RotateDegress
        Degrees_0 = 0
        Degrees_90 = 1
        Degrees_180 = 2
        Degrees_270 = 3
    End Enum

    Public Enum RotateResult
        RotateFailed = User32.DISP_CHANGE_FAILED
        RotateSuccessful = User32.DISP_CHANGE_SUCCESSFUL
        RotateRequiresRestart = User32.DISP_CHANGE_RESTART
    End Enum

    Public Function RotateScreen(ByVal pRotateDegrees As RotateDegress) As RotateResult

        Dim dm As DEVMODE : dm = New DEVMODE()
        Dim rotRet As RotateResult : rotRet = RotateResult.RotateSuccessful

        dm.dmDeviceName = New String(Chr(0), 32)
        dm.dmFormName = New String(Chr(0), 32)
        dm.dmSize = CType(Marshal.SizeOf(dm), Short)

        If (User32.EnumDisplaySettings(Nothing, User32.ENUM_CURRENT_SETTINGS, dm) <> 0) Then

            ' determine unrotated pixel sizes
            Select Case CType(dm.dmDisplayOrientation, RotateDegress)
                Case RotateDegress.Degrees_90, RotateDegress.Degrees_270
                    SwapExtents(dm.dmPelsHeight, dm.dmPelsWidth)
            End Select

            ' set rotation parameters
            dm.dmDisplayOrientation = CType(pRotateDegrees, Short)
            Select Case pRotateDegrees
                Case RotateDegress.Degrees_90, RotateDegress.Degrees_270
                    SwapExtents(dm.dmPelsHeight, dm.dmPelsWidth)
            End Select
            dm.dmFields = User32.DM_DISPLAYORIENTATION Or User32.DM_PELSHEIGHT Or User32.DM_PELSWIDTH

            ' query change
            rotRet = User32.ChangeDisplaySettings(dm, User32.CDS_TEST)

            If rotRet = RotateResult.RotateFailed Then
                Return RotateResult.RotateFailed
            Else
                'enact change
                rotRet = User32.ChangeDisplaySettings(dm, User32.CDS_UPDATEREGISTRY)
                Select Case rotRet
                    Case RotateResult.RotateSuccessful, RotateResult.RotateRequiresRestart
                        Return rotRet
                    Case Else
                        Return RotateResult.RotateFailed
                End Select
            End If

        Else
            Return RotateResult.RotateFailed
        End If

    End Function

    Private Sub SwapExtents(ByRef pelHeight As Integer, ByRef pelWidth As Integer)
        pelHeight = pelHeight Xor pelWidth
        pelWidth = pelWidth Xor pelHeight
        pelHeight = pelHeight Xor pelWidth
    End Sub

End Class
