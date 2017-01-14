::'@cscript //nologo //e:vbscript "%~f0" %* & @goto :eof

' This file is a batch and vbscript hybrid.
' To function properly it must contain SUB character (0x1A)
' right after the ::' characters in the first line.


' Loads LISP file to an executed AutoCAD/BricsCAD/IntelliCAD application.
' Usage: autolisp [opions] filename

Const ha_AutoCAD = 1
Const ha_BricsCAD = 2
Const ha_IntelliCAD = 3

Dim happ
happ = ha_AutoCAD


Class AutoLisp

    Private App, VLApp
    Private HostApp

    Public Default Function Init(happ)
        HostApp = happ
        Dim progID
        Select Case HostApp
            Case ha_AutoCAD
                progID = "AutoCAD.Application"
            Case ha_BricsCAD
                progID = "BricscadApp.AcadApplication"
            Case ha_IntelliCAD
                progID = "ICAD.Application"
            Case Else
                Die "Init: Unsupported CAD application."
        End Select
        On Error Resume Next
        ' Set App = WScript.GetObject(, progID, "AcadApp_") ' Should allow event handling
        Set App = GetObject(, progID)
        If App Is Nothing Then
            Die "Cannot connect to CAD application."
        End If
        If HostApp = ha_AutoCAD Then
            Set VLApp = App.GetInterfaceObject("VL.Application.16")
            If VLApp Is Nothing Then
                Die "Cannot connect to VisualLISP module."
            End If
        End If
        On Error Goto 0
        Set Init = Me
    End Function

    ' Private Sub Class_Initialize
    ' End Sub

    Private Sub Class_Terminate
        Set VLApp = Nothing
        Set App = Nothing
    End Sub

    Private Function ActiveDocument()
        Dim doc
        On Error Resume Next
        Set doc = App.ActiveDocument
        If doc Is Nothing Then
            Die "Cannot get active CAD document."
        End If
        On Error Goto 0
        Set ActiveDocument = doc
    End Function

    Public Sub Load(filename)
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FileExists(filename) Then
            Die "Cannot find the file specified."
        End If

        filename = Replace(fso.GetAbsolutePathName(filename), "\", "/")

        Dim doc
        Set doc = ActiveDocument ' Test if there is an ActiveDocument
        Select Case HostApp
            Case ha_AutoCAD
                VLApp.ActiveDocument.Functions.Item("load").funcall(CStr(filename))
            Case ha_BricsCAD
                ActiveDocument.EvaluateLisp _
                    "(eval (read """ & _
                    "(load \""" & filename & "\"")" & _
                    """))"
            Case ha_IntelliCAD
                App.LoadLISP(filename)
            Case Else
                Die "Load: Unsupported CAD application."
        End Select
        Set doc = Nothing
    End Sub

End Class


Sub Die(msg)
    WScript.Echo "AutoLISP: " & msg
    WScript.Quit(1)
End Sub

Sub Help()
    WScript.Echo "Usage: autolisp [options] filename"
    WScript.Echo
    WScript.Echo "  -h, --help  Show this help"
    WScript.Echo "  --acad      Use AutoCAD as a host application"
    WScript.Echo "  --bcad      Use BricsCAD as a host application"
    WScript.Echo "  --icad      Use IntelliCAD as a host application"
    WScript.Quit()
End Sub

If WScript.Arguments.Count = 0 Then
    Die "No LISP file specified."
Else
    Dim arg, i
    For i = 0 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        If Left(arg, 1) = "-" Then
            Select Case arg
                Case "-h", "--help"
                    Help
                Case "--acad"
                    happ = ha_AutoCAD
                Case "--bcad"
                    happ = ha_BricsCAD
                Case "--icad"
                    happ = ha_IntelliCAD
                Case Else
                    Die "Bad option: " & arg
            End Select
        Else
            Dim filename
            filename = arg
        End If
    Next
End If

Dim al
Set al = (New AutoLisp)(happ)
al.Load filename
Set al = Nothing
