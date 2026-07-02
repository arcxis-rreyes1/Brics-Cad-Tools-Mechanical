Imports Bricscad.ApplicationServices
Imports Teigha.Runtime

<Assembly: ExtensionApplication(GetType(Arcxis_Cad_Tools.PluginApplication))>
Namespace Arcxis_Cad_Tools
    Public Class PluginApplication
        Implements IExtensionApplication

        Public Sub Initialize() Implements IExtensionApplication.Initialize
            ArcxisPaths.InitializeArcxisPaths()
        End Sub

        Public Sub Terminate() Implements IExtensionApplication.Terminate
        End Sub
    End Class
End Namespace
