﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On


Namespace My
    
    <Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "17.9.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Partial Friend NotInheritable Class MySettings
        Inherits Global.System.Configuration.ApplicationSettingsBase
        
        Private Shared defaultInstance As MySettings = CType(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New MySettings()),MySettings)
        
#Region "My.Settings Auto-Save Functionality"
#If _MyType = "WindowsForms" Then
    Private Shared addedHandler As Boolean

    Private Shared addedHandlerLockObject As New Object

    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(), Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Private Shared Sub AutoSaveSettings(sender As Global.System.Object, e As Global.System.EventArgs)
        If My.Application.SaveMySettingsOnExit Then
            My.Settings.Save()
        End If
    End Sub
#End If
#End Region
        
        Public Shared ReadOnly Property [Default]() As MySettings
            Get
                
#If _MyType = "WindowsForms" Then
               If Not addedHandler Then
                    SyncLock addedHandlerLockObject
                        If Not addedHandler Then
                            AddHandler My.Application.Shutdown, AddressOf AutoSaveSettings
                            addedHandler = True
                        End If
                    End SyncLock
                End If
#End If
                Return defaultInstance
            End Get
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property AlwaysOnTop() As Boolean
            Get
                Return CType(Me("AlwaysOnTop"),Boolean)
            End Get
            Set
                Me("AlwaysOnTop") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("not yet set")>  _
        Public Property ApplicationVersion() As String
            Get
                Return CType(Me("ApplicationVersion"),String)
            End Get
            Set
                Me("ApplicationVersion") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("150")>  _
        Public Property ColumnWidthDescription() As Double
            Get
                Return CType(Me("ColumnWidthDescription"),Double)
            End Get
            Set
                Me("ColumnWidthDescription") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("150")>  _
        Public Property ColumnWidthListenFor() As Double
            Get
                Return CType(Me("ColumnWidthListenFor"),Double)
            End Get
            Set
                Me("ColumnWidthListenFor") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("100")>  _
        Public Property ColumnWidthOpen() As Double
            Get
                Return CType(Me("ColumnWidthOpen"),Double)
            End Get
            Set
                Me("ColumnWidthOpen") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("100")>  _
        Public Property ColumnWidthStartIn() As Double
            Get
                Return CType(Me("ColumnWidthStartIn"),Double)
            End Get
            Set
                Me("ColumnWidthStartIn") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property CheckForUpdate() As Boolean
            Get
                Return CType(Me("CheckForUpdate"),Boolean)
            End Get
            Set
                Me("CheckForUpdate") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("2")>  _
        Public Property CheckForUpdateFrequency() As Integer
            Get
                Return CType(Me("CheckForUpdateFrequency"),Integer)
            End Get
            Set
                Me("CheckForUpdateFrequency") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("100")>  _
        Public Property Left() As Integer
            Get
                Return CType(Me("Left"),Integer)
            End Get
            Set
                Me("Left") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("700, 450")>  _
        Public Property MainWindowSize() As Global.System.Drawing.Size
            Get
                Return CType(Me("MainWindowSize"),Global.System.Drawing.Size)
            End Get
            Set
                Me("MainWindowSize") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property PushBulletAPI() As String
            Get
                Return CType(Me("PushBulletAPI"),String)
            End Get
            Set
                Me("PushBulletAPI") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property ShowPush2RunAtStartup() As Boolean
            Get
                Return CType(Me("ShowPush2RunAtStartup"),Boolean)
            End Get
            Set
                Me("ShowPush2RunAtStartup") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property StartBossAsAdministratorByDefault() As Boolean
            Get
                Return CType(Me("StartBossAsAdministratorByDefault"),Boolean)
            End Get
            Set
                Me("StartBossAsAdministratorByDefault") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property StartBossOnLogon() As Boolean
            Get
                Return CType(Me("StartBossOnLogon"),Boolean)
            End Get
            Set
                Me("StartBossOnLogon") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("100")>  _
        Public Property Top() As Integer
            Get
                Return CType(Me("Top"),Integer)
            End Get
            Set
                Me("Top") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property ViewDescription() As Boolean
            Get
                Return CType(Me("ViewDescription"),Boolean)
            End Get
            Set
                Me("ViewDescription") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property ViewListenFor() As Boolean
            Get
                Return CType(Me("ViewListenFor"),Boolean)
            End Get
            Set
                Me("ViewListenFor") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property ViewOpen() As Boolean
            Get
                Return CType(Me("ViewOpen"),Boolean)
            End Get
            Set
                Me("ViewOpen") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property ViewParameters() As Boolean
            Get
                Return CType(Me("ViewParameters"),Boolean)
            End Get
            Set
                Me("ViewParameters") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Property NextVersionCheckDate() As Date
            Get
                Return CType(Me("NextVersionCheckDate"),Date)
            End Get
            Set
                Me("NextVersionCheckDate") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property PushBulletTitleFilter() As String
            Get
                Return CType(Me("PushBulletTitleFilter"),String)
            End Get
            Set
                Me("PushBulletTitleFilter") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property MyUniqueID() As String
            Get
                Return CType(Me("MyUniqueID"),String)
            End Get
            Set
                Me("MyUniqueID") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property SortByDescription() As Boolean
            Get
                Return CType(Me("SortByDescription"),Boolean)
            End Get
            Set
                Me("SortByDescription") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property ViewStartIn() As Boolean
            Get
                Return CType(Me("ViewStartIn"),Boolean)
            End Get
            Set
                Me("ViewStartIn") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property ViewAdmin() As Boolean
            Get
                Return CType(Me("ViewAdmin"),Boolean)
            End Get
            Set
                Me("ViewAdmin") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("-1")>  _
        Public Property ColumnWidthParameters() As Double
            Get
                Return CType(Me("ColumnWidthParameters"),Double)
            End Get
            Set
                Me("ColumnWidthParameters") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("25")>  _
        Public Property ColumnWidthAdmin() As Double
            Get
                Return CType(Me("ColumnWidthAdmin"),Double)
            End Get
            Set
                Me("ColumnWidthAdmin") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property UACLimit() As Boolean
            Get
                Return CType(Me("UACLimit"),Boolean)
            End Get
            Set
                Me("UACLimit") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property ConfirmExit() As Boolean
            Get
                Return CType(Me("ConfirmExit"),Boolean)
            End Get
            Set
                Me("ConfirmExit") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property ConfirmRedX() As Boolean
            Get
                Return CType(Me("ConfirmRedX"),Boolean)
            End Get
            Set
                Me("ConfirmRedX") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property ConfirmDelete() As Boolean
            Get
                Return CType(Me("ConfirmDelete"),Boolean)
            End Get
            Set
                Me("ConfirmDelete") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property FirstTimeRun() As Boolean
            Get
                Return CType(Me("FirstTimeRun"),Boolean)
            End Get
            Set
                Me("FirstTimeRun") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property ViewFilters() As String
            Get
                Return CType(Me("ViewFilters"),String)
            End Get
            Set
                Me("ViewFilters") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property AutoScroll() As Boolean
            Get
                Return CType(Me("AutoScroll"),Boolean)
            End Get
            Set
                Me("AutoScroll") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property ViewStartingWindowState() As Boolean
            Get
                Return CType(Me("ViewStartingWindowState"),Boolean)
            End Get
            Set
                Me("ViewStartingWindowState") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("100")>  _
        Public Property ColumnWidthStartingWindowState() As Double
            Get
                Return CType(Me("ColumnWidthStartingWindowState"),Double)
            End Get
            Set
                Me("ColumnWidthStartingWindowState") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("150")>  _
        Public Property ColumnWidthKeysToSend() As Double
            Get
                Return CType(Me("ColumnWidthKeysToSend"),Double)
            End Get
            Set
                Me("ColumnWidthKeysToSend") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property ViewKeysToSend() As Boolean
            Get
                Return CType(Me("ViewKeysToSend"),Boolean)
            End Get
            Set
                Me("ViewKeysToSend") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property WriteLogToDisk() As Boolean
            Get
                Return CType(Me("WriteLogToDisk"),Boolean)
            End Get
            Set
                Me("WriteLogToDisk") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("not available")>  _
        Public Property PushoverSecret() As String
            Get
                Return CType(Me("PushoverSecret"),String)
            End Get
            Set
                Me("PushoverSecret") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("not available")>  _
        Public Property PushoverID() As String
            Get
                Return CType(Me("PushoverID"),String)
            End Get
            Set
                Me("PushoverID") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("not available")>  _
        Public Property PushoverDeviceName() As String
            Get
                Return CType(Me("PushoverDeviceName"),String)
            End Get
            Set
                Me("PushoverDeviceName") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("not available")>  _
        Public Property PushoverDeviceID() As String
            Get
                Return CType(Me("PushoverDeviceID"),String)
            End Get
            Set
                Me("PushoverDeviceID") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property PushoverUserID() As String
            Get
                Return CType(Me("PushoverUserID"),String)
            End Get
            Set
                Me("PushoverUserID") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property UsePushbullet() As Boolean
            Get
                Return CType(Me("UsePushbullet"),Boolean)
            End Get
            Set
                Me("UsePushbullet") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property UsePushover() As Boolean
            Get
                Return CType(Me("UsePushover"),Boolean)
            End Get
            Set
                Me("UsePushover") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("1066-10-14")>  _
        Public Property LastRequesttoKeepPushbulletToKeepAccountActive() As Date
            Get
                Return CType(Me("LastRequesttoKeepPushbulletToKeepAccountActive"),Date)
            End Get
            Set
                Me("LastRequesttoKeepPushbulletToKeepAccountActive") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("60")>  _
        Public Property CompetingPushThreshold() As Integer
            Get
                Return CType(Me("CompetingPushThreshold"),Integer)
            End Get
            Set
                Me("CompetingPushThreshold") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("not available")>  _
        Public Property DropboxPath() As String
            Get
                Return CType(Me("DropboxPath"),String)
            End Get
            Set
                Me("DropboxPath") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Command.txt")>  _
        Public Property DropboxFileName() As String
            Get
                Return CType(Me("DropboxFileName"),String)
            End Get
            Set
                Me("DropboxFileName") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("not available")>  _
        Public Property DropboxDeviceName() As String
            Get
                Return CType(Me("DropboxDeviceName"),String)
            End Get
            Set
                Me("DropboxDeviceName") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property UseDropbox() As Boolean
            Get
                Return CType(Me("UseDropbox"),Boolean)
            End Get
            Set
                Me("UseDropbox") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("c:\")>  _
        Public Property ImportExportDirectory() As String
            Get
                Return CType(Me("ImportExportDirectory"),String)
            End Get
            Set
                Me("ImportExportDirectory") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("1066-10-14")>  _
        Public Property LastDownload() As Date
            Get
                Return CType(Me("LastDownload"),Date)
            End Get
            Set
                Me("LastDownload") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("do not skip update")>  _
        Public Property SkipUpdateFor() As String
            Get
                Return CType(Me("SkipUpdateFor"),String)
            End Get
            Set
                Me("SkipUpdateFor") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property ImportTag() As Boolean
            Get
                Return CType(Me("ImportTag"),Boolean)
            End Get
            Set
                Me("ImportTag") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property ImportOnByDefault() As Boolean
            Get
                Return CType(Me("ImportOnByDefault"),Boolean)
            End Get
            Set
                Me("ImportOnByDefault") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property ImportConfirmation() As Boolean
            Get
                Return CType(Me("ImportConfirmation"),Boolean)
            End Get
            Set
                Me("ImportConfirmation") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("and, then")>  _
        Public Property SeparatingWords() As String
            Get
                Return CType(Me("SeparatingWords"),String)
            End Get
            Set
                Me("SeparatingWords") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property ShowNotifications() As Boolean
            Get
                Return CType(Me("ShowNotifications"),Boolean)
            End Get
            Set
                Me("ShowNotifications") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property ShowNotificationSource() As Boolean
            Get
                Return CType(Me("ShowNotificationSource"),Boolean)
            End Get
            Set
                Me("ShowNotificationSource") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property IncludeDisconnectAndReconnect() As Boolean
            Get
                Return CType(Me("IncludeDisconnectAndReconnect"),Boolean)
            End Get
            Set
                Me("IncludeDisconnectAndReconnect") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property ShowNotificationResult() As Boolean
            Get
                Return CType(Me("ShowNotificationResult"),Boolean)
            End Get
            Set
                Me("ShowNotificationResult") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property UseMQTT() As Boolean
            Get
                Return CType(Me("UseMQTT"),Boolean)
            End Get
            Set
                Me("UseMQTT") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property MQTTBroker() As String
            Get
                Return CType(Me("MQTTBroker"),String)
            End Get
            Set
                Me("MQTTBroker") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("")>  _
        Public Property MQTTUser() As String
            Get
                Return CType(Me("MQTTUser"),String)
            End Get
            Set
                Me("MQTTUser") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("not available")>  _
        Public Property MQTTPassword() As String
            Get
                Return CType(Me("MQTTPassword"),String)
            End Get
            Set
                Me("MQTTPassword") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("1883")>  _
        Public Property MQTTPort() As Integer
            Get
                Return CType(Me("MQTTPort"),Integer)
            End Get
            Set
                Me("MQTTPort") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("#")>  _
        Public Property MQTTFilter() As String
            Get
                Return CType(Me("MQTTFilter"),String)
            End Get
            Set
                Me("MQTTFilter") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property SuppressStartupNotice() As Boolean
            Get
                Return CType(Me("SuppressStartupNotice"),Boolean)
            End Get
            Set
                Me("SuppressStartupNotice") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property SandboxMessageHasBeenShown() As Boolean
            Get
                Return CType(Me("SandboxMessageHasBeenShown"),Boolean)
            End Get
            Set
                Me("SandboxMessageHasBeenShown") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("False")>  _
        Public Property WeclomeMessageHasBeenShown() As Boolean
            Get
                Return CType(Me("WeclomeMessageHasBeenShown"),Boolean)
            End Get
            Set
                Me("WeclomeMessageHasBeenShown") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property ViewDisabledCards() As Boolean
            Get
                Return CType(Me("ViewDisabledCards"),Boolean)
            End Get
            Set
                Me("ViewDisabledCards") = value
            End Set
        End Property
    End Class
End Namespace

Namespace My
    
    <Global.Microsoft.VisualBasic.HideModuleNameAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute()>  _
    Friend Module MySettingsProperty
        
        <Global.System.ComponentModel.Design.HelpKeywordAttribute("My.Settings")>  _
        Friend ReadOnly Property Settings() As Global.Push2Run.My.MySettings
            Get
                Return Global.Push2Run.My.MySettings.Default
            End Get
        End Property
    End Module
End Namespace