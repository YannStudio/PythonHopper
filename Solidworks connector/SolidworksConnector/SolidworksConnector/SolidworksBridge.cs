using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;


namespace SolidworksConnector
{
    public static class SolidworksBridge
    {
        public static SldWorks Connect(bool visible = true)
        {
            SldWorks app;
            try
            {
                app = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            }
            catch (COMException)
            {
                app = new SldWorks();
            }

            app.Visible = visible;
            return app;
        }

    }
}
