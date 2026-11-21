
using System;
using System.Threading;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointVideoPlayer
{
    public partial class ThisAddIn
    {
        private Thread glThread;
        private volatile OpenGLWindow glWindow;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            this.Application.SlideShowBegin += Application_SlideShowBegin;
            this.Application.SlideShowEnd += Application_SlideShowEnd;
        }

        private void Application_SlideShowBegin(PowerPoint.SlideShowWindow Wn)
        {
            if (glThread != null && glThread.IsAlive)
                return;

            IntPtr parent = new IntPtr(Wn.HWND);

            glThread = new Thread(() =>
            {
                using (var window = new OpenGLWindow(parent))
                {
                    glWindow = window;

                    // 0 = без ограничения частоты кадров
                    window.Run(0.0);

                    glWindow = null;
                }
            });

            glThread.IsBackground = true;
            glThread.SetApartmentState(ApartmentState.STA);
            glThread.Start();
        }

        private void Application_SlideShowEnd(PowerPoint.Presentation Pres)
        {
            OpenGLWindow window = glWindow;

            if (window != null)
                window.Exit();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            OpenGLWindow window = glWindow;

            if (window != null)
                window.Exit();
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
