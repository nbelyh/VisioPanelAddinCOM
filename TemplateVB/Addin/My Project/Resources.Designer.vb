﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.18444
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources
    
    'This class was auto-generated by the StronglyTypedResourceBuilder
    'class via a tool like ResGen or Visual Studio.
    'To add or remove a member, edit your .ResX file then rerun ResGen
    'with the /str option, or rebuild your VS project.
    '''<summary>
    '''  A strongly-typed resource class, for looking up localized strings, etc.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0"),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.Microsoft.VisualBasic.HideModuleNameAttribute()>  _
    Friend Module Resources
        
        Private resourceMan As Global.System.Resources.ResourceManager
        
        Private resourceCulture As Global.System.Globalization.CultureInfo
        
        '''<summary>
        '''  Returns the cached ResourceManager instance used by this class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("$csprojectname$.Resources", GetType(Resources).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property
        
        '''<summary>
        '''  Overrides the current thread's CurrentUICulture property for all
        '''  resource lookups using this strongly typed resource class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = value
            End Set
        End Property
        $if$ ($uiCallbacks$ == true)
        '''<summary>
        '''  Looks up a localized string similar to command 1.
        '''</summary>
        Friend ReadOnly Property Command1_Label() As String
            Get
                Return ResourceManager.GetString("Command1_Label", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to command 2.
        '''</summary>
        Friend ReadOnly Property Command2_Label() As String
            Get
                Return ResourceManager.GetString("Command2_Label", resourceCulture)
            End Get
        End Property
        $endif$ $if$ ($ribbon$ == true)
        '''<summary>
        '''  Looks up a localized resource of type System.Drawing.Bitmap similar to (Bitmap).
        '''</summary>
        Friend ReadOnly Property Command1() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("Command1", resourceCulture)
                Return CType(obj, System.Drawing.Bitmap)
            End Get
        End Property
        $endif$ $if$ ($ribbonXml$ == true)
        '''<summary>
        '''  Looks up a localized resource of type System.Drawing.Bitmap similar to (Bitmap).
        '''</summary>
        Friend ReadOnly Property Command2() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("Command2", resourceCulture)
                Return CType(obj, System.Drawing.Bitmap)
            End Get
        End Property
        $endif$ $if$ ($commandbars$ == true)
        '''<summary>
        '''  Looks up a localized resource of type System.Drawing.Bitmap similar to (Bitmap).
        '''</summary>
        Friend ReadOnly Property Command1_sm() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("Command1_sm", resourceCulture)
                Return CType(obj, System.Drawing.Bitmap)
            End Get
        End Property

        '''<summary>
        '''  Looks up a localized resource of type System.Drawing.Bitmap similar to (Bitmap).
        '''</summary>
        Friend ReadOnly Property Command2_sm() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("Command2_sm", resourceCulture)
                Return CType(obj, System.Drawing.Bitmap)
            End Get
        End Property
        $endif$ $if$ ($taskpaneANDui$ == true)
        '''<summary>
        '''  Looks up a localized resource of type System.Drawing.Bitmap similar to (Bitmap).
        '''</summary>
        Friend ReadOnly Property TogglePanel() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("TogglePanel", resourceCulture)
                Return CType(obj, System.Drawing.Bitmap)
            End Get
        End Property
        $endif$ $if$ ($taskpaneANDuiCallbacks$ == true)
        '''<summary>
        '''  Looks up a localized string similar to Toggle panel.
        '''</summary>
        Friend ReadOnly Property TogglePanel_Label() As String
            Get
                Return ResourceManager.GetString("TogglePanel_Label", resourceCulture)
            End Get
        End Property
        $endif$ $if$ ($commandbarsANDtaskpane$ == true)
        '''<summary>
        '''  Looks up a localized resource of type System.Drawing.Bitmap similar to (Bitmap).
        '''</summary>
        Friend ReadOnly Property TogglePanel_sm() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("TogglePanel_sm", resourceCulture)
                Return CType(obj, System.Drawing.Bitmap)
            End Get
        End Property
  		$endif$ $if$ ($ribbonXml$ == true)
        '''<summary>
        '''  Looks up a localized string similar to &lt;?xml version=&quot;1.0&quot; encoding=&quot;UTF-8&quot;?&gt;
        '''&lt;customUI xmlns=&quot;http://schemas.microsoft.com/office/2009/07/customui&quot; onLoad=&quot;OnRibbonLoad&quot;&gt;
        '''  &lt;ribbon&gt;
        '''    &lt;tabs&gt;
        '''      &lt;tab idMso=&quot;TabHome&quot;&gt;
        '''        &lt;group id=&quot;Group1&quot; label=&quot;$csprojectname$&quot;&gt;
        '''
        '''          &lt;button id=&quot;Command1&quot; size=&quot;large&quot; getLabel=&quot;OnGetRibbonLabel&quot; onAction=&quot;OnRibbonButtonClick&quot; getEnabled=&quot;IsRibbonCommandEnabled&quot; getImage=&quot;GetRibbonImage&quot; /&gt;
        '''          &lt;button id=&quot;Command2&quot; size=&quot;large&quot; getLabel=&quot;OnGetRibbonLabel&quot; onAction=&quot;OnRibbonButtonC [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property Ribbon() As String
            Get
                Return ResourceManager.GetString("Ribbon", resourceCulture)
            End Get
        End Property
        $endif$
    End Module
End Namespace
