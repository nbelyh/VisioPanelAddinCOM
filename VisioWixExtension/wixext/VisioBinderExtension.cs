using System.Runtime.InteropServices;
using Microsoft.Tools.WindowsInstallerXml;

namespace VisioWixExtension
{
    /// <summary>
    /// The decompiler for the Windows Installer XML Toolset Gaming Extension.
    /// </summary>
    public sealed class VisioBinderExtension : BinderExtension
    {
        public override void DatabaseFinalize(Output output)
        {
            var filesTable = output.Tables["WixFile"];

            foreach (WixFileRow row in filesTable.Rows)
            {
                Core.OnMessage(new WixGenericMessageEventArgs(null, 0, MessageLevel.Information, "{0} => {1}", row.File, row.Source));
            }

            base.DatabaseFinalize(output);
        }
    }
}
