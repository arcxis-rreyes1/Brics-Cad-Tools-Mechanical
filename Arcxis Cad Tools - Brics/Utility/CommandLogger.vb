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

        Public Sub Initialize() Implements IExtensionApplication.Initialize

            UNCPath()
            Module_Arcxis_TB.InitializeArcxisPaths()

            AddHandler Application.DocumentManager.DocumentCreated, AddressOf OnDocumentCreated
        End Sub

        Public Sub Terminate() Implements IExtensionApplication.Terminate
            RemoveHandler Application.DocumentManager.DocumentCreated, AddressOf OnDocumentCreated
            If _idleHandlerAttached Then
                RemoveHandler Application.Idle, AddressOf OnIdleFixMasterSealXref
                _idleHandlerAttached = False
            End If
            _pendingSealFixDoc = Nothing
        End Sub

        Private Shared Sub OnDocumentCreated(sender As Object, e As DocumentCollectionEventArgs)
            If e Is Nothing OrElse e.Document Is Nothing Then Return

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

            Try
                Arcxis_Cad_Tools.ExportToCad.FixMasterSealFileXrefPath(doc)
            Catch
            End Try
        End Sub

    End Class
End Namespace
