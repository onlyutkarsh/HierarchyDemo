using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;

namespace Utkarsh.HierarchyDemo
{
    /// <summary>
    /// This class implements the tool window exposed by this package and hosts a user control.
    ///
    /// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane, 
    /// usually implemented by the package implementer.
    ///
    /// This class derives from the ToolWindowPane class provided from the MPF in order to use its 
    /// implementation of the IVsUIElementPane interface.
    /// </summary>
    [Guid("c3fc9d92-4a03-46a4-abee-047529674048")]
    public class MyToolWindow : ToolWindowPane
    {
        /// <summary>
        /// Standard constructor for the tool window.
        /// </summary>
        public MyToolWindow() :
            base(null)
        {
            // Set the window title reading it from the resources.
            this.Caption = Resources.ToolWindowTitle;
            // Set the image that will appear on the tab of the window frame
            // when docked with an other window
            // The resource ID correspond to the one defined in the resx file
            // while the Index is the offset in the bitmap strip. Each image in
            // the strip being 16x16.
            this.BitmapResourceID = 301;
            this.BitmapIndex = 1;

            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on 
            // the object returned by the Content property.
            //base.Content = new MyControl();
            this.ToolClsid = VSConstants.CLSID_VsUIHierarchyWindow;
        }

        public override void OnToolWindowCreated()
        {
            base.OnToolWindowCreated();

            // Initialize the hierarchy window with desired styles...
            Object unkObj;
            uint grfUIHWF = (uint)__UIHWINFLAGS.UIHWF_DoNotSortRootNodes |
                (uint)__UIHWINFLAGS.UIHWF_SupportToolWindowToolbars |
                (uint)__UIHWINFLAGS.UIHWF_RouteCmdidDelete |
                (uint)__UIHWINFLAGS.UIHWF_ActAsProjectTypeWin;

            // Initialize with custom hierarchy
            IVsUIHierarchy hierarchy = new SimpleHierarchy.SimpleHierarchy(HierarchyWindow) as IVsUIHierarchy;
            //IVsUIHierarchy hierarchy = new SimpleHierarchy(HierarchyWindow) as IVsUIHierarchy;
            HierarchyWindow.Init(hierarchy, grfUIHWF, out unkObj);

            // Add a toolbar to the toolwindow
            IVsToolWindowToolbarHost toolbarHost = unkObj as IVsToolWindowToolbarHost;
            if (toolbarHost != null)
            {
                Guid guidToolbar = GuidList.guidHierarchyDemoCmdSet;
                uint toolbarID = 0x0500;
                toolbarHost.AddToolbar(VSTWT_LOCATION.VSTWT_TOP, ref guidToolbar, toolbarID);
                toolbarHost.ShowHideToolbar(ref guidToolbar, toolbarID, 1);
            }

            // Or add the custom hierarchy here... 
            //HierarchyWindow.AddUIHierarchy(new SimpleHierarchy(HierarchyWindow) as IVsUIHierarchy, 0);
        }

        public IVsUIHierarchyWindow HierarchyWindow
        {
            get
            {
                IVsWindowFrame frame = this.Frame as IVsWindowFrame;
                if (frame != null)
                {
                    Object docView;
                    int hr = frame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out docView);
                    if (hr == VSConstants.S_OK)
                        return (IVsUIHierarchyWindow)docView;
                }
                return null;
            }
        }
    }


}
