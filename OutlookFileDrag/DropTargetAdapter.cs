using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.IO;
using System.Runtime.InteropServices.ComTypes;

namespace OutlookFileDrag
{
    class DropTargetAdapter : NativeMethods.IDropTarget
    {
        private static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private NativeMethods.IDropTarget _delegate;

        public DropTargetAdapter(NativeMethods.IDropTarget dropTarget)
        {
            _delegate = dropTarget;
        }

        public int OleDragEnter([In, MarshalAs(UnmanagedType.Interface)] NativeMethods.IDataObject pDataObj, [In, MarshalAs(UnmanagedType.U4)] int grfKeyState, [In, MarshalAs(UnmanagedType.U8)] long pt, [In, Out] ref int pdwEffect)
        {
            log.Info("OleDragEnter");
            return _delegate.OleDragEnter(Wrap(pDataObj), grfKeyState, pt, ref pdwEffect);
        }

        public int OleDragOver([In, MarshalAs(UnmanagedType.U4)] int grfKeyState, [In, MarshalAs(UnmanagedType.U8)] long pt, [In, Out] ref int pdwEffect)
        {
            log.Info("OleDragOver");
            return _delegate.OleDragOver(grfKeyState, pt, ref pdwEffect);
        }

        public int OleDragLeave()
        {
            log.Info("OleLeave");
            return _delegate.OleDragLeave();
        }

        public int OleDrop([In, MarshalAs(UnmanagedType.Interface)] NativeMethods.IDataObject pDataObj, [In, MarshalAs(UnmanagedType.U4)] int grfKeyState, [In, MarshalAs(UnmanagedType.U8)] long pt, [In, Out] ref int pdwEffect)
        {
            log.Info("OleDrop");
            return _delegate.OleDrop(Wrap(pDataObj), grfKeyState, pt, ref pdwEffect);
        }

        [return: MarshalAs(UnmanagedType.Interface)]
        internal NativeMethods.IDataObject Wrap([In, MarshalAs(UnmanagedType.Interface)] NativeMethods.IDataObject pDataObj)
        {
            int r = pDataObj.EnumFormatEtc(DATADIR.DATADIR_GET, out IEnumFORMATETC f);
            if (r != NativeMethods.S_OK)
            {
                log.Error("Could not EnumFormatEtc");
            }
            else
            {
                log.Info("Wrap - EnumFormatEtc Ok");
                FORMATETC[] buffer = new FORMATETC[] { new FORMATETC() };
                while (f.Next(1, buffer, null) == NativeMethods.S_OK)
                {
                    FORMATETC c = buffer[0];
                    log.InfoFormat("Format {0}  tymed {1}", c.cfFormat, c.tymed);
                    System.Windows.Forms.DataFormats.Format z = System.Windows.Forms.DataFormats.GetFormat(c.cfFormat);
                    log.InfoFormat("Format name {0}", z.Name);
                }
            }

            string url = DataObjectHelper.GetContentUnicode(pDataObj, "UniformResourceLocatorW");
            if (!(url is null) && url.StartsWith("about:"))
            {
                url = DataObjectHelper.GetContentUnicode(pDataObj, "UnicodeText");
            }
            if (url is null)
            {
                log.Info("No UniformResourceLocatorW found");

                // TODO Check for ANSI URL
            }
            log.InfoFormat("UnicodeText: {0}", DataObjectHelper.GetContentUnicode(pDataObj, "UnicodeText"));
            log.InfoFormat("Text: {0}", DataObjectHelper.GetContentAnsi(pDataObj, "Text"));
            log.InfoFormat("UniformResourceLocatorW: {0}", DataObjectHelper.GetContentUnicode(pDataObj, "UniformResourceLocatorW"));
            log.InfoFormat("UniformResourceLocator: {0}", DataObjectHelper.GetContentAnsi(pDataObj, "UniformResourceLocator"));
            log.InfoFormat("text/html: {0}", DataObjectHelper.GetContentAnsi(pDataObj, "text/html"));
            log.InfoFormat("HTML format: {0}", DataObjectHelper.GetContentAnsi(pDataObj, "HTML Format"));
            log.InfoFormat("Object Descriptor: {0}", DataObjectHelper.GetContentAnsi(pDataObj, "Object Descriptor"));
            log.InfoFormat("FileName: {0}", DataObjectHelper.GetContentAnsi(pDataObj, "FileName", 0));
     //       log.InfoFormat("FileContents: {0}", DataObjectHelper.GetContentAnsi(pDataObj, "FileContents", 0));
            log.InfoFormat("Link: {0}", DataObjectHelper.GetContentAnsi(pDataObj, "Link", 0));


            log.Info("==============");

            if (url is null)
            {
                return pDataObj;
            }
            else
            {
                // TODO Check the URL is a case records document


                // return new OutlookDataObjectBase(pDataObj, new string[] { "\\\\AP-WIN.ap.local\\Apshared\\IT\\temp\\2.pdf" });
                //return new OutlookDataObjectBase(pDataObj, new string[] { "C:\\Users\\PhilScull\\Desktop\\cnipa.html" });
                return new OutlookDataObjectBase(pDataObj, new string[] { url });
            }
        }
    }
}
