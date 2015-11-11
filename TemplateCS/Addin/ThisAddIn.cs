﻿using System;
using Office = Microsoft.Office.Core;
$if$ ($uiCallbacks$ == true)using System.Drawing;
$endif$$if$ ($ribbonANDcommandbars$ == true)using System.Globalization;
$endif$$if$ ($ui$ == true)using System.Windows.Forms;
$endif$$if$ ($uiCallbacks$ == true)using $csprojectname$.Properties;
$endif$using Visio = Microsoft.Office.Interop.Visio;

namespace $csprojectname$
{
    [GuidAttribute("$csprojectguid$")]
    [ProgId("$csprojectname$.Addin")]
    public partial class ThisAddIn : Extensibility.IDTExtensibility2
    {
        $if$ ($commandbars$ == true)
        private readonly AddinCommandBars _addinCommandBars = new AddinCommandBars();
        $endif$$if$ ($ribbonXml$ == true)
        private readonly AddinRibbon _addinRibbon = new AddinRibbon();
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return _addinRibbon;
        }
        $endif$$if$ ($ui$ == true)
        /// <summary>
        /// A simple command
        /// </summary>
        public void Command1()
        {
            MessageBox.Show("Hello from command 1!");
        }
        $endif$$if$ ($uiCallbacks$ == true)
        /// <summary>
        /// A command to demonstrate conditionally enabling/disabling.
        /// The command gets enabled only when a shape is selected
        /// </summary>
        public void Command2()
        {
            if (Application == null || Application.ActiveWindow == null || Application.ActiveWindow.Selection == null)
                return;

            MessageBox.Show(string.Format("Hello from (conditional) command 2! You have {0} shapes selected.", Application.ActiveWindow.Selection.Count));
        }
        $endif$$if$ ($uiCallbacks$ == true)
        /// <summary>
        /// Callback called by the UI manager when user clicks a button
        /// Should do something meaningful when corresponding action is called.
        /// </summary>
        public void OnCommand(string commandId)
        {
            switch (commandId)
            {
                case "Command1":
                    Command1();
                    return;
                $endif$$if$ ($uiCallbacks$ == true)
                case "Command2":
                    Command2();
                    return;
                $endif$$if$ ($taskpaneANDuiCallbacks$ == true)
                case "TogglePanel":
                    TogglePanel();
                    return;
            $endif$$if$ ($uiCallbacks$ == true)}
        }
        $endif$$if$ ($uiCallbacks$ == true)
        /// <summary>
        /// Callback called by UI manager.
        /// Should return if corresponding command should be enabled in the user interface.
        /// By default, all commands are enabled.
        /// </summary>
        public bool IsCommandEnabled(string commandId)
        {
            switch (commandId)
            {
                case "Command1":    // make command1 always enabled
                    return true;

                case "Command2":    // make command2 enabled only if a drawing is opened
                    return Application != null 
                        && Application.ActiveWindow != null
                        && Application.ActiveWindow.Selection.Count > 0;
                $endif$$if$ ($taskpaneANDuiCallbacks$ == true)
                case "TogglePanel": // make panel enabled only if we have selected shape(s).
                    return IsPanelEnabled();

                $endif$$if$ ($uiCallbacks$ == true) default:
                    return true;
            }
        }

        /// <summary>
        /// Callback called by UI manager.
        /// Should return if corresponding command (button) is pressed or not (makes sense for toggle buttons)
        /// </summary>
        public bool IsCommandChecked(string command)
        {
            $endif$$if$ ($taskpaneANDuiCallbacks$ == true)
            if (command == "TogglePanel")
                return IsPanelVisible();
            
            $endif$$if$($uiCallbacks$ == true)return false;
        }
        /// <summary>
        /// Callback called by UI manager.
        /// Returns a label associated with given command.
        /// We assume for simplicity taht command labels are named simply named as [commandId]_Label (see resources)
        /// </summary>
        public string GetCommandLabel(string command)
        {
            return Resources.ResourceManager.GetString(command + "_Label");
        }

        /// <summary>
        /// Returns a bitmap associated with given command.
        /// We assume for simplicity that bitmap ids are named after command id.
        /// </summary>
        public Bitmap GetCommandBitmap(string id)
        {
            return (Bitmap)Resources.ResourceManager.GetObject(id);
        }
        $endif$$if$ ($taskpane$ == true)
        public void TogglePanel()
        {
            _panelManager.TogglePanel(Application.ActiveWindow);
        }
        $endif$$if$ ($taskpaneANDuiCallbacks$ == true)
        public bool IsPanelEnabled()
        {
            return Application != null && Application.ActiveWindow != null;
        }

        public bool IsPanelVisible()
        {
            return Application != null 
                && _panelManager != null 
                && _panelManager.IsPanelOpened(Application.ActiveWindow);
        }
        $endif$$if$ ($taskpane$ == true)
        private PanelManager _panelManager;
        $endif$$if$ ($uiCallbacks$ == true)
        internal void UpdateUI()
        {
            $endif$$if$ ($commandbars$ == true) _addinCommandBars.UpdateCommandBars();
            $endif$$if$ ($ribbonXml$ == true)_addinRibbon.UpdateRibbon();
        $endif$$if$ ($uiCallbacks$ == true)}
        $endif$$if$ ($uiCallbacks$ == true)
        private void Application_SelectionChanged(Visio.Window window)
        {
            UpdateUI();
        }
        $endif$
        private void Startup(object application)
        {
            $if$ ($taskpane$ == true)_panelManager = new PanelManager();
            $endif$$if$ ($ribbonANDcommandbars$ == true)var version = int.Parse(Application.Version, NumberStyles.AllowDecimalPoint);
            if (version < 14)
                $endif$$if$ ($commandbars$ == true)_addinCommandBars.StartupCommandBars("$csprojectname$", new[] { "Command1", "Command2" $endif$$if$ ($commandbarsANDtaskpane$ == true) , "TogglePanel"$endif$$if$ ($commandbars$ == true)});
            $endif$$if$ ($uiCallbacks$ == true)Application.SelectionChanged += Application_SelectionChanged;
            $endif$
		}

        private void Shutdown()
        {
            $if$ ($commandbars$ == true)_addinCommandBars.ShutdownCommandBars();
            $endif$$if$ ($taskpane$ == true)_panelManager.Dispose();
            $endif$$if$ ($uiCallbacks$ == true)Application.SelectionChanged -= Application_SelectionChanged;
            $endif$
		}

        #region IDTExtensibility2

        public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            Startup(application);
        }

        public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref Array custom)
        {
            Shutdown();
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        #endregion // IDTExtensibility2
	}
}
