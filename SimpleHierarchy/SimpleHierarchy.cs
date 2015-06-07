using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using IOleServiceProvider = Microsoft.VisualStudio.OLE.Interop.IServiceProvider;


namespace Utkarsh.HierarchyDemo.SimpleHierarchy
{

    class SimpleHierarchy : IVsUIHierarchy, IVsHierarchy
    {
        internal static ImageList _imageList = null;
        internal static SimpleHierarchy Hierarchy = null;

        internal SimpleItem rootItem = new SimpleItem("Simple Hierarchy", unchecked((uint)-2));
        internal SimpleItem childItem1 = new SimpleItem("MSDN", 1, VSConstants.VSITEMID_ROOT);
        internal SimpleItem childItem2 = new SimpleItem("MSNBC", 2, VSConstants.VSITEMID_ROOT);

        internal static ServiceProvider _serviceProvider;
        private IVsUIHierarchyWindow _vsHierarchyWindow = null;

        public SimpleHierarchy(IVsUIHierarchyWindow hierWin)
        {
            _vsHierarchyWindow = hierWin;
            SimpleHierarchy.Hierarchy = this;

            if (_imageList == null)
            {
                _imageList = new ImageList();
                _imageList.ColorDepth = ColorDepth.Depth24Bit;
                _imageList.ImageSize = new Size(16, 16);
                _imageList.TransparentColor = Color.Magenta;
                _imageList.Images.AddStrip(Resources.SimpleHierarchyImages);

            }
        }

        SimpleItem GetItem(uint itemid)
        {
            switch (itemid)
            {
                case VSConstants.VSITEMID_ROOT:
                    return rootItem;
                case 1:
                    return childItem1;
                case 2:
                    return childItem2;
            }
            return null;
        }

        #region IVsUIHierarchy Members

        public int AdviseHierarchyEvents(IVsHierarchyEvents pEventSink, out uint pdwCookie)
        {
            pdwCookie = 0;
            return VSConstants.S_OK;
        }

        public int Close()
        {
            return VSConstants.S_OK;
        }

        public int ExecCommand(uint itemid, ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            if (pguidCmdGroup.Equals(VSConstants.GUID_VsUIHierarchyWindowCmds))
            {
                switch (nCmdID)
                {
                    case (uint)VSConstants.VsUIHierarchyWindowCmdIds.UIHWCMDID_DoubleClick:
                        {
                            switch (itemid)
                            {
                                case 1:
                                    childItem1.NavigateTo("http://www.msdn.microsoft.com");
                                    break;
                                case 2:
                                    childItem2.NavigateTo("http://www.msnbc.com");
                                    break;
                            }
                        }
                        break;
                }

                return (int)Microsoft.VisualStudio.OLE.Interop.Constants.OLECMDERR_E_NOTSUPPORTED;
            }

            return (int)Microsoft.VisualStudio.OLE.Interop.Constants.OLECMDERR_E_NOTSUPPORTED;
        }

        public int GetCanonicalName(uint itemid, out string pbstrName)
        {
            pbstrName = GetItem(itemid)._Caption;
            if (pbstrName != null)
                return VSConstants.S_OK;

            return VSConstants.E_INVALIDARG;
        }

        public int GetGuidProperty(uint itemid, int propid, out Guid pguid)
        {
            string s = string.Format("GetGuidProperty for itemID({0})", itemid.ToString());
            Trace.WriteLine(s);
            pguid = Guid.Empty;
            return VSConstants.DISP_E_MEMBERNOTFOUND;
        }

        public int GetNestedHierarchy(uint itemid, ref Guid iidHierarchyNested, out IntPtr ppHierarchyNested, out uint pitemidNested)
        {
            ppHierarchyNested = IntPtr.Zero;
            pitemidNested = 0;
            return VSConstants.E_NOTIMPL;
        }

        public int GetProperty(uint itemid, int propid, out object pvar)
        {
            // GetProperty is called many many times for this particular property
            if (propid != (int)__VSHPROPID.VSHPROPID_ParentHierarchy)
            {
                string s = string.Format("GetProperty for itemId ({0}) called with propid = {1}", itemid.ToString(), propid.ToString());
                Trace.WriteLine(s);
            }

            pvar = null;
            switch (propid)
            {
                case (int)__VSHPROPID.VSHPROPID_CmdUIGuid:
                    pvar = typeof(HierarchyDemoPackage).GUID;
                    break;

                case (int)__VSHPROPID.VSHPROPID_Parent:
                    if (itemid == VSConstants.VSITEMID_ROOT)
                        pvar = VSConstants.VSITEMID_NIL;
                    else
                        pvar = VSConstants.VSITEMID_ROOT;
                    break;

                case (int)__VSHPROPID.VSHPROPID_FirstChild:
                    if (itemid == VSConstants.VSITEMID_ROOT)
                        pvar = childItem1._Id;
                    else
                        pvar = VSConstants.VSITEMID_NIL;
                    break;

                case (int)__VSHPROPID.VSHPROPID_NextSibling:
                    if (itemid == childItem1._Id)
                        pvar = childItem2._Id;
                    else
                        pvar = VSConstants.VSITEMID_NIL;
                    break;

                case (int)__VSHPROPID.VSHPROPID_Expandable:
                    if (itemid == VSConstants.VSITEMID_ROOT)
                        pvar = true;
                    else
                        pvar = false;
                    break;

                case (int)__VSHPROPID.VSHPROPID_IconImgList:
                case (int)__VSHPROPID.VSHPROPID_OpenFolderIconHandle:
                    pvar = (int)_imageList.Handle;
                    break;

                case (int)__VSHPROPID.VSHPROPID_IconIndex:
                case (int)__VSHPROPID.VSHPROPID_OpenFolderIconIndex:
                    pvar = GetItem(itemid)._IconIndex;
                    break;

                case (int)__VSHPROPID.VSHPROPID_Caption:
                case (int)__VSHPROPID.VSHPROPID_SaveName:
                    pvar = GetItem(itemid)._Caption;
                    break;

                case (int)__VSHPROPID.VSHPROPID_ShowOnlyItemCaption:
                    pvar = true;
                    break;

                case (int)__VSHPROPID.VSHPROPID_ParentHierarchy:
                    if (itemid == childItem1._Id || itemid == childItem2._Id)
                        pvar = this as IVsHierarchy;
                    break;
            }

            if (pvar != null)
                return VSConstants.S_OK;

            return VSConstants.DISP_E_MEMBERNOTFOUND;
        }

        public int GetSite(out IOleServiceProvider ppSP)
        {
            ppSP = _serviceProvider.GetService(typeof(IOleServiceProvider)) as IOleServiceProvider;
            return VSConstants.S_OK;
        }

        public int ParseCanonicalName(string pszName, out uint pitemid)
        {
            pitemid = 0;
            return VSConstants.E_NOTIMPL;
        }

        public int QueryClose(out int pfCanClose)
        {
            pfCanClose = 1;
            return VSConstants.S_OK;
        }

        public int QueryStatusCommand(uint itemid, ref Guid pguidCmdGroup, uint cCmds, Microsoft.VisualStudio.OLE.Interop.OLECMD[] prgCmds, IntPtr pCmdText)
        {
            return (int)Microsoft.VisualStudio.OLE.Interop.Constants.OLECMDERR_E_UNKNOWNGROUP;
        }

        public int SetGuidProperty(uint itemid, int propid, ref Guid rguid)
        {
            return VSConstants.E_NOTIMPL;
        }

        public int SetProperty(uint itemid, int propid, object var)
        {
            return VSConstants.E_NOTIMPL;
        }

        public int SetSite(Microsoft.VisualStudio.OLE.Interop.IServiceProvider psp)
        {
            _serviceProvider = new ServiceProvider(psp, true);
            return VSConstants.S_OK;
        }

        public int UnadviseHierarchyEvents(uint dwCookie)
        {
            return VSConstants.S_OK;
        }

        public int Unused0()
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public int Unused1()
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public int Unused2()
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public int Unused3()
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public int Unused4()
        {
            throw new Exception("The method or operation is not implemented.");
        }

        #endregion

        internal class SimpleItem : IVsWebBrowserUser
        {
            public string _Caption;
            public uint _Id;
            public uint _ParentId;
            public int _IconIndex;

            private IVsWebBrowser _webBrowser = null;
            private IVsWindowFrame _frameWindow = null;

            public SimpleItem(string caption, uint id)
            {
                _Caption = caption;
                _Id = id;
                _ParentId = VSConstants.VSITEMID_ROOT;
                _IconIndex = 0;
            }

            public SimpleItem(string caption, uint id, uint parentId)
            {
                _Caption = caption;
                _Id = id;
                _ParentId = parentId;
                _IconIndex = 1;
            }

            public void NavigateTo(string strURL)
            {
                int hr = 0;

                // create a webbrowser instance and tie it to our hierarchy item
                if (_webBrowser == null)
                {
                    IVsWebBrowsingService svc = HierarchyDemoPackage.Package.GetSvc(typeof(SVsWebBrowsingService)) as IVsWebBrowsingService;
                    uint dwCreateFlags = (uint)__VSCREATEWEBBROWSER.VSCWB_FrameMdiChild | (uint)__VSCREATEWEBBROWSER.VSCWB_StartCustom;
                    Guid guidEmpty = GuidList.guidHierAnarchyPersistance;
                    hr = svc.CreateWebBrowser(dwCreateFlags, ref guidEmpty, _Caption, strURL, (IVsWebBrowserUser)this, out _webBrowser, out _frameWindow);

                    _frameWindow.SetProperty((int)__VSFPROPID.VSFPROPID_Hierarchy, SimpleHierarchy.Hierarchy);
                    _frameWindow.SetProperty((int)__VSFPROPID.VSFPROPID_ItemID, this._Id);

                }
                else
                    _webBrowser.Navigate(0, strURL);


                if (_frameWindow != null)
                    _frameWindow.Show();
            }


            #region IVsWebBrowserUser Members

            public int Disconnect()
            {
                return VSConstants.S_OK;
            }

            public int FilterDataObject(Microsoft.VisualStudio.OLE.Interop.IDataObject pDataObjIn, out Microsoft.VisualStudio.OLE.Interop.IDataObject ppDataObjOut)
            {
                throw new Exception("The method or operation is not implemented.");
            }

            public int GetCmdUIGuid(out Guid pguidCmdUI)
            {
                pguidCmdUI = GuidList.guidHierarchyDemoCmdSet;
                return VSConstants.S_OK;
            }

            public int GetCustomMenuInfo(object pUnkCmdReserved, object pDispReserved, uint dwType, uint dwPosition, out Guid pguidCmdGroup, out int pdwMenuID)
            {
                throw new Exception("The method or operation is not implemented.");
            }

            public int GetCustomURL(uint nPage, out string pbstrURL)
            {
                throw new Exception("The method or operation is not implemented.");
            }

            public int GetDropTarget(Microsoft.VisualStudio.OLE.Interop.IDropTarget pDropTgtIn, out Microsoft.VisualStudio.OLE.Interop.IDropTarget ppDropTgtOut)
            {
                throw new Exception("The method or operation is not implemented.");
            }

            public int GetExternalObject(out object ppDispObject)
            {
                throw new Exception("The method or operation is not implemented.");
            }

            public int GetOptionKeyPath(uint dwReserved, out string pbstrKey)
            {
                throw new Exception("The method or operation is not implemented.");
            }

            public int Resize(int cx, int cy)
            {
                throw new Exception("The method or operation is not implemented.");
            }

            public int TranslateAccelarator(MSG[] lpmsg)
            {
                throw new Exception("The method or operation is not implemented.");
            }

            public int TranslateUrl(uint dwReserved, string lpszURLIn, out string lppszURLOut)
            {
                throw new Exception("The method or operation is not implemented.");
            }

            #endregion
        }
    }


}
