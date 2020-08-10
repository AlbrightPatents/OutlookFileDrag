using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using log4net;

namespace OutlookFileDrag
{
    // Class that wraps a drop data object converts URLs to file contents
    class OutlookDataObjectDrop : NativeMethods.IDataObject, ICustomQueryInterface  
    {
        // TODO - Do we need this here any more?
        private NativeMethods.IDataObject innerData;

        private string[] urls;
        private static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public bool FilesDropped { get; private set; }

        public OutlookDataObjectDrop(NativeMethods.IDataObject innerData, string[] urls)
        {
            log.DebugFormat("OutlookDataObjectDrop {0}", urls);
            this.innerData = innerData;
            this.urls = urls;
        }

        public static string UrlToFilePath(string url)
        {
            Uri uri = UrlToUri(url);
            return UriToFilePath(uri);
        }

        public static string UriToFilePath(Uri uri)
        {
            string query = uri.Query;
            log.DebugFormat("query of {0} is {1}", uri, query);
            string path = uri.LocalPath;
            log.DebugFormat("Read path from URI {0}", path);
            return path.Substring(0, path.Length - query.Length); ;
        }

        public static string UriToFileName(Uri uri)
        {
            string[] segments = uri.Segments;
            string lastSegment = segments[segments.Length - 1];
            log.DebugFormat("Last segment of {0} is {1}", uri, lastSegment);
            return lastSegment;
        }

        public static Uri UrlToUri(string url)
        {
            string[] s = url.Split('?');
            return new Uri(s[0]);
        }

        public static string QueryParamsToFilename(string query)
        {
            System.Collections.Specialized.NameValueCollection x = Mono.Web.HttpUtility.ParseQueryString(query);
            return x.Get("filename");
        }

        public static string UrlToFileName(string url)
        {
            // TODO Improve the way the query params are separated from the URL
            string[] s = url.Split('?');
            if (s.Length > 1)
            {
                string filename = QueryParamsToFilename(s[1]);
                if (!(filename is null)) return filename;
            }
            Uri uri = UrlToUri(url);
            return UriToFileName(uri);
        }

        public static string[] UrlsToFilenames(string[] urls)
        {
            string[] filenames = new string[urls.Length];
            for(int i = 0; i < urls.Length; ++i)
            {
                filenames[i] = UrlToFileName(urls[i]);
            }
            return filenames;
        }

        // TODO Not actually called by outlook. Needs testing.
        public int EnumFormatEtc(DATADIR direction, out IEnumFORMATETC ppenumFormatEtc)
        {
            IEnumFORMATETC origEnum = null;
            try
            {
                log.InfoFormat("IDataObject.EnumFormatEtc called -- direction {0}", direction);
                switch (direction)
                {
                    case DATADIR.DATADIR_GET:
                        // Enumerate formats
                        List<FORMATETC> formats = new List<FORMATETC>();
                                               
                        // FileContents
                        formats.Add(new FORMATETC
                        {
                            cfFormat = (short)System.Windows.Forms.DataFormats.GetFormat("FileContents").Id,
                            dwAspect = DVASPECT.DVASPECT_CONTENT,
                            lindex = -1,
                            ptd = IntPtr.Zero,
                            tymed = TYMED.TYMED_HGLOBAL
                        });

                        // FileGroupDescriptorW
                        formats.Add(new FORMATETC
                        {
                            cfFormat = (short)System.Windows.Forms.DataFormats.GetFormat("FileGroupDescriptorW").Id,
                            dwAspect = DVASPECT.DVASPECT_CONTENT,
                            lindex = -1,
                            ptd = IntPtr.Zero,
                            tymed = TYMED.TYMED_HGLOBAL
                        });

                        // FileGroupDescriptor
                        formats.Add(new FORMATETC
                        {
                            cfFormat = (short)System.Windows.Forms.DataFormats.GetFormat("FileGroupDescriptor").Id,
                            dwAspect = DVASPECT.DVASPECT_CONTENT,
                            lindex = -1,
                            ptd = IntPtr.Zero,
                            tymed = TYMED.TYMED_HGLOBAL
                        });

                        //Return new enumerator for available formats
                        ppenumFormatEtc = new FormatEtcEnumerator(formats.ToArray());
                        return NativeMethods.S_OK;

                    case DATADIR.DATADIR_SET:
                        //Return original enumerator
                        // TODO This should probably just return an error code
                        return innerData.EnumFormatEtc(direction, out ppenumFormatEtc);

                    default:
                        //Invalid direction
                        ppenumFormatEtc = null;
                        return NativeMethods.E_INVALIDARG;
                }
            }
            catch (Exception ex)
            {
                log.Error("Exception in IDataObject.EnumFormatEtc", ex);
                ppenumFormatEtc = null;
                return NativeMethods.E_UNEXPECTED;
            }
            finally
            {
                // Release all unmanaged objects
                if (origEnum != null)
                    Marshal.ReleaseComObject(origEnum);
            }
        }

        // TODO Not actually called by outlook. Needs testing.
        public int GetCanonicalFormatEtc(ref FORMATETC formatIn, out FORMATETC formatOut)
        {
            try
            {
                log.InfoFormat("IDataObject.GetCanonicalFormatEtc called -- cfFormat {0} dwAspect {1} lindex {2} ptd {3} tymed {4}", formatIn.cfFormat, formatIn.dwAspect, formatIn.lindex, formatIn.ptd, formatIn.tymed);
                string formatName = System.Windows.Forms.DataFormats.GetFormat((ushort)formatIn.cfFormat).Name;
                log.InfoFormat("IDataObject.GetCanonicalFormatEtc Format name: {0}", formatName);

                switch(formatName) {
                    case "FileContents":
                    case "FileGroupDescriptor":
                    case "FileGroupDescriptorW":
                        // Copy input format to output format
                        formatOut = new FORMATETC
                        {
                            cfFormat = formatIn.cfFormat,
                            dwAspect = formatIn.dwAspect,
                            lindex = formatIn.lindex,
                            ptd = IntPtr.Zero,
                            tymed = formatIn.tymed
                        };

                        return NativeMethods.DATA_S_SAMEFORMATETC;
                    default:
                        formatOut = new FORMATETC();
                        return NativeMethods.E_UNEXPECTED;
                }
            }
            catch (Exception ex)
            {
                log.Error("Exception in IDataObject.GetCanonicalFormatEtc", ex);
                formatOut = new FORMATETC();
                return NativeMethods.E_UNEXPECTED;
            }
        }

        public int GetData(ref FORMATETC format, out STGMEDIUM medium)
        {
            try
            {
                // Get data into passed medium
                log.DebugFormat("IDataObject.GetData called -- cfFormat {0} dwAspect {1} lindex {2} ptd {3} tymed {4}", format.cfFormat, format.dwAspect, format.lindex, format.ptd, format.tymed);
                string formatName = System.Windows.Forms.DataFormats.GetFormat((ushort)format.cfFormat).Name;
                log.DebugFormat("IDataObject.GetData Format name: {0}", formatName);

                if (formatName == "FileGroupDescriptorW")
                {
                    medium = new STGMEDIUM();
                    DataObjectHelper.SetFileGroupDescriptorW(ref medium, UrlsToFilenames(urls));
                    return NativeMethods.S_OK;
                }
                else if (formatName == "FileGroupDescriptor")
                {
                    medium = new STGMEDIUM();
                    DataObjectHelper.SetFileGroupDescriptor(ref medium, UrlsToFilenames(urls));
                    return NativeMethods.S_OK;
                }
                else if (formatName == "FileContents")
                {
                    medium = new STGMEDIUM();
                    int index = format.lindex;
                    if (index >=0 && index < urls.Length) {
                        // TODO Currently this only works with locally accessible file: URLs
                        // TODO Add support for HTTP/HTTPS fetches
                        DataObjectHelper.SetFileContents(ref medium, UrlToFilePath(urls[index]));
                    }
                    else
                    {
                        return NativeMethods.DV_E_LINDEX;
                    }
                    return NativeMethods.S_OK;
                }
                else
                {
                    medium = new STGMEDIUM();
                    return NativeMethods.DV_E_FORMATETC;
                }

            }
            catch (Exception ex)
            {
                log.Error("Exception in IDataObject.GetData", ex);
                medium = new STGMEDIUM();
                return NativeMethods.E_UNEXPECTED;
            }
        }

        public int GetDataHere(ref FORMATETC format, ref STGMEDIUM medium)
        {
            log.InfoFormat("IDataObject.QueryGetData called -- cfFormat {0} dwAspect {1} lindex {2} ptd {3} tymed {4}", format.cfFormat, format.dwAspect, format.lindex, format.ptd, format.tymed);
            return NativeMethods.E_NOTIMPL;
        }

        public int QueryGetData(ref FORMATETC format)
        {
            try
            {
                log.DebugFormat("IDataObject.QueryGetData called -- cfFormat {0} dwAspect {1} lindex {2} ptd {3} tymed {4}", format.cfFormat, format.dwAspect, format.lindex, format.ptd, format.tymed);
                string formatName = System.Windows.Forms.DataFormats.GetFormat((ushort)format.cfFormat).Name;
                log.DebugFormat("IDataObject.QueryGetData Format name: {0}", formatName);

                int r;
                switch(formatName)
                {
                    case "FileContents":
                    case "FileGroupDescriptor":
                    case "FileGroupDescriptorW":
                        r = NativeMethods.S_OK;
                        break;
                    default:
                      r =  NativeMethods.DV_E_FORMATETC;
                      break;
                }

                log.DebugFormat("IDataObject.QueryGetData response {0}", r);
                return r;
            }
            catch (Exception ex)
            {
                log.Error("Exception in IDataObject.QueryGetData", ex);
                return NativeMethods.E_UNEXPECTED;
            }
        }

        public int SetData(ref FORMATETC formatIn, ref STGMEDIUM medium, bool release)
        {
            try
            {
                log.InfoFormat("IDataObject.SetData called -- cfFormat {0} dwAspect {1} lindex {2} ptd {3} tymed {4}", formatIn.cfFormat, formatIn.dwAspect, formatIn.lindex, formatIn.ptd, formatIn.tymed);
                log.InfoFormat("IDataObject.SetData Format name: {0}", System.Windows.Forms.DataFormats.GetFormat((ushort)formatIn.cfFormat).Name);
                int result = innerData.SetData(formatIn, medium, release);
                log.InfoFormat("Result: {0}", result);
                return result;

            }
            catch (Exception ex)
            {
                log.Error("Exception in IDataObject.SetData", ex);
                return NativeMethods.E_UNEXPECTED;
            }
        }

        public int DAdvise(ref FORMATETC pFormatetc, ADVF advf, IAdviseSink adviseSink, out int connection)
        {
            log.Info("DAdvise");
            return innerData.DAdvise(pFormatetc, advf, adviseSink, out connection);
        }

        public int DUnadvise(int connection)
        {
            log.Info("DUnadvise");
            return innerData.DUnadvise(connection);
        }

        public int EnumDAdvise(out IEnumSTATDATA enumAdvise)
        {
            log.Info("EnumDAdvise");
            return innerData.EnumDAdvise(out enumAdvise);
        }

        public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
        {
            ppv = IntPtr.Zero;
            try
            {
                log.InfoFormat("Get COM interface {0}", iid);

                //For IDataObject interface, use interface on this object
                if (iid == new Guid("0000010E-0000-0000-C000-000000000046"))
                {
                    log.InfoFormat("Interface handled");
                    return CustomQueryInterfaceResult.NotHandled;
                }

                else
                {
                    //For all other interfaces, use interface on original object
                    IntPtr pUnk = Marshal.GetIUnknownForObject(this.innerData);
                    int retVal = Marshal.QueryInterface(pUnk, ref iid, out ppv);
                    if (retVal == NativeMethods.S_OK)
                    {
                        log.InfoFormat("Interface handled by inner object");
                        return CustomQueryInterfaceResult.Handled;
                    }
                    else
                    {
                        log.InfoFormat("Interface not handled by inner object");
                        return CustomQueryInterfaceResult.Failed;
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error("Exception in ICustomQueryInterface", ex);
                return CustomQueryInterfaceResult.Failed;
            }
        }
    }
}
