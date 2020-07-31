using System;
using EasyHook;
using log4net;

namespace OutlookFileDrag
{
    class RegisterDragDropHook : IDisposable
    {
        private static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private LocalHook hook;
        private bool disposed = false;
        private bool isHooked = false;

        public RegisterDragDropHook()
        {
            try
            {
                //Hook OLE RegisterDragDrop event
                log.Info("Creating hook for RegisterDragDrop method of ole32.dll");
                hook = EasyHook.LocalHook.Create(EasyHook.LocalHook.GetProcAddress("ole32.dll", "RegisterDragDrop"),
                    new NativeMethods.RegisterDragDropDelegate(RegisterDragDropHook.DoRegisterDragDropHook), null);
            }
            catch (Exception ex)
            {
                log.Error("Error creating RegisterDragDrop hook", ex);
                throw;
            }
        }

        public bool IsHooked
        {
            get
            {
                return isHooked;
            }
        }

        public void Start()
        {
            try
            {
                if (isHooked)
                    return;

                log.Info("Starting RegisterDragDrop hook");
                //Hook current (UI) thread
                // TODO the following line prevents outlook from starting
                hook.ThreadACL.SetInclusiveACL(new Int32[] { 0 });
                isHooked = true;
            }
            catch (Exception ex)
            {
                log.Error("Error starting hook", ex);
                throw;
            }
        }

        public void Stop()
        {
            try
            {
                if (!isHooked)
                    return;

                log.Info("Stopping hook");
                //Stop hooking all threads
                hook.ThreadACL.SetInclusiveACL(new Int32[] { });
                isHooked = false;
                log.Info("Stopped hook");
            }
            catch (Exception ex)
            {
                log.Error("Error stopping hook", ex);
                throw;
            }
        }

        public static int DoRegisterDragDropHook(IntPtr hwnd, NativeMethods.IDropTarget pDropTarget)
        {
            try
            {
                log.Info("RegisterDragDrop started");

                log.Info("Calling real RegisterDragDrop...");

                return NativeMethods.RegisterDragDrop(hwnd, new DropTargetAdapter(pDropTarget));
            }
            catch (Exception ex)
            {
                log.Warn("RegisterDragDrop error", ex);
                return NativeMethods.S_OK;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                hook.Dispose();
            }

            disposed = true;
        }
    }

}
