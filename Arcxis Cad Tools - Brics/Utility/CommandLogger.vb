Imports System.IO
Imports Bricscad.ApplicationServices
Imports Teigha.Runtime
Imports Teigha.DatabaseServices
Imports Bricscad.EditorInput
Imports Teigha.Geometry
Imports Teigha.Colors
Imports Application = Bricscad.ApplicationServices.Application
Imports Document = Bricscad.ApplicationServices.Document

' Put this OUTSIDE any class or namespace
<Assembly: ExtensionApplication(GetType(LoggingTrial.CommandLogger))>
Namespace LoggingTrial
    Public Class CommandLogger
        Implements IExtensionApplication

        Private Shared _pendingSealFixDoc As Document
        Private Shared _idleHandlerAttached As Boolean
        Private Shared _lastIdleResumeCheckUtc As DateTime = DateTime.MinValue
        Private Shared ReadOnly IdleResumeCheckInterval As TimeSpan = TimeSpan.FromSeconds(30)

        Public Sub Initialize() Implements IExtensionApplication.Initialize
            Arcxis_Cad_Tools.ArcxisActivityLog.LogPluginInit()

            UNCPath()
            Module_Arcxis_TB.InitializeArcxisPaths()

            AddHandler Application.DocumentManager.DocumentCreated, AddressOf OnDocumentCreated
            AddHandler Application.Idle, AddressOf OnIdleCheckResume
            Arcxis_Cad_Tools.ArcxisActivityLog.WriteLog("EVENT|DocumentCreated handler attached")
        End Sub

        Public Sub Terminate() Implements IExtensionApplication.Terminate
            Arcxis_Cad_Tools.ArcxisActivityLog.WriteLog("EVENT|DocumentCreated handler detaching")
            RemoveHandler Application.Idle, AddressOf OnIdleCheckResume
            RemoveHandler Application.DocumentManager.DocumentCreated, AddressOf OnDocumentCreated
            If _idleHandlerAttached Then
                RemoveHandler Application.Idle, AddressOf OnIdleFixMasterSealXref
                _idleHandlerAttached = False
            End If
            _pendingSealFixDoc = Nothing
            Arcxis_Cad_Tools.ArcxisActivityLog.LogPluginTerminate()
        End Sub

        Private Shared Sub OnDocumentCreated(sender As Object, e As DocumentCollectionEventArgs)
            If e Is Nothing OrElse e.Document Is Nothing Then Return

            Arcxis_Cad_Tools.ArcxisActivityLog.WriteLog("EVENT|DocumentCreated|" & e.Document.Name)
            _pendingSealFixDoc = e.Document

            If Not _idleHandlerAttached Then
                AddHandler Application.Idle, AddressOf OnIdleFixMasterSealXref
                _idleHandlerAttached = True
            End If
        End Sub

        Private Shared Sub OnIdleFixMasterSealXref(sender As Object, e As EventArgs)
            If _idleHandlerAttached Then
                RemoveHandler Application.Idle, AddressOf OnIdleFixMasterSealXref
                _idleHandlerAttached = False
            End If

            Dim doc As Document = _pendingSealFixDoc
            _pendingSealFixDoc = Nothing

            If doc Is Nothing Then Return

            Arcxis_Cad_Tools.ArcxisActivityLog.WriteLog("EVENT|Idle seal xref fix start|" & doc.Name)
            Try
                Arcxis_Cad_Tools.ExportToCad.FixMasterSealFileXrefPath(doc)
                Arcxis_Cad_Tools.ArcxisActivityLog.WriteLog("EVENT|Idle seal xref fix end|" & doc.Name)
            Catch ex As System.Exception
                Arcxis_Cad_Tools.ArcxisActivityLog.WriteLog(ex, "Idle seal xref fix|" & doc.Name)
            End Try
        End Sub

        Private Shared Sub OnIdleCheckResume(sender As Object, e As EventArgs)
            Dim now = DateTime.UtcNow
            If _lastIdleResumeCheckUtc <> DateTime.MinValue AndAlso
               (now - _lastIdleResumeCheckUtc) < IdleResumeCheckInterval Then
                Return
            End If

            _lastIdleResumeCheckUtc = now
            Arcxis_Cad_Tools.ArcxisActivityLog.CheckIdleResume()
        End Sub

    End Class
End Namespace
