Attribute VB_Name = "Resmod"
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
    Private Const CCDEVICENAME = 32
    Private Const CCFORMNAME = 32
    Private Const DM_BITSPERPEL = &H60000
    Private Const DM_PELSWIDTH = &H80000
    Private Const DM_PELSHEIGHT = &H100000
Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Dim RetValue As Integer

Function getres() As Integer
    Dim hmpfff As Single
    Dim DevM As DEVMODE, I As Integer, ReturnVal As Boolean, _
    RetValue, OldWidth As Single, OldHeight As Single, _
    OldBPP As Integer
    Call EnumDisplaySettings(0&, -1, DevM)
    OldWidth = DevM.dmPelsWidth
    OldHeight = DevM.dmPelsHeight
    OldBPP = DevM.dmBitsPerPel
     DevM.dmPelsWidth = OldWidth
    DevM.dmPelsHeight = OldHeight
    DevM.dmBitsPerPel = OldBPP
    hmpfff = MsgBox("Current Resolution is (" & OldWidth & " x " & OldHeight & ", " & OldBPP & " Bit) If this is 1024*768 then click ok, else click cancel, change your resolution to 1024*768 then restart....", vbInformation + vbOKCancel, "Resolution Confirm:")
    If RetValue = vbCancel Then
       End
    End If
End Function

