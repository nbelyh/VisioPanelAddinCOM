Imports System
Imports System.Runtime.InteropServices
Imports Extensibility

<GuidAttribute("$clsid$")> _
<ProgId("$progid$")> _
Partial Public Class Addin
    Implements IDTExtensibility2

#Region "IDTExtensibility2"

    Public Sub OnConnection(app As Object, connectMode As ext_ConnectMode, addInInst As Object, ByRef [custom] As Array) Implements IDTExtensibility2.OnConnection
        Startup(app)
    End Sub

    Public Sub OnDisconnection(disconnectMode As ext_DisconnectMode, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection
        Shutdown()
    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As Array) Implements IDTExtensibility2.OnAddInsUpdate
    End Sub

    Public Sub OnStartupComplete(ByRef custom As Array) Implements IDTExtensibility2.OnStartupComplete
    End Sub

    Public Sub OnBeginShutdown(ByRef custom As Array) Implements IDTExtensibility2.OnBeginShutdown
    End Sub

#End Region

End Class