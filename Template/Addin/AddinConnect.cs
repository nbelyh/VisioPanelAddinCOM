using System;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace $csprojectname$
{
    [GuidAttribute("$csprojectguid$")]
    [ProgId("$csprojectname$.Addin")]
    public partial class Addin : Extensibility.IDTExtensibility2
    {
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