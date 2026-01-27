Imports System.IO
Imports Bricscad.ApplicationServices
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors

' Put this OUTSIDE any class or namespace
<Assembly: ExtensionApplication(GetType(LoggingTrial.CommandLogger))>
Namespace LoggingTrial
    Public Class CommandLogger
        Implements IExtensionApplication
        Public Sub Initialize() Implements IExtensionApplication.Initialize

            UNCPath()
            Module_Arcxis_TB.InitializeArcxisPaths()
        End Sub

        Public Sub Terminate() Implements IExtensionApplication.Terminate
        End Sub

    End Class
End Namespace