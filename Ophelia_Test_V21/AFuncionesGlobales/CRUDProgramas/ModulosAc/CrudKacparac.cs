using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Drawing;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales;
using OpenQA.Selenium;
using System.Drawing;
using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;


namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosAc
{
    class CrudKacparac : FuncionesVitales
    {
        static public WindowsDriver<WindowsElement> RootSession()
        {
            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }


        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string parametro)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(parametro);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex

                var appCapabilities = new AppiumOptions();
                //appCapabilities.AddAdditionalCapability("app", "Root");
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }

        

        static public void CrudPrincipal(WindowsDriver<WindowsElement> desktopSession, List<string> crudPrincipal)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 190, 44);
            desktopSession.Mouse.DoubleClick(null);
            desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
            Thread.Sleep(2000);          
        }




        static public void CrudDet1(WindowsDriver<WindowsElement> desktopSession, List<string> crudDet1)
        {
            // Se ingresa info
            var btnTDVI1 = desktopSession.FindElementsByClassName("TScrollBox");
            desktopSession.Mouse.MouseMove(btnTDVI1[0].Coordinates, 140, 162);
            desktopSession.Mouse.DoubleClick(null);
            desktopSession.Keyboard.SendKeys(crudDet1[0]);
            Thread.Sleep(1000);
        }



        //CrudDetalle del centro  
        static public void BotonesCrudDetalleCentro(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            switch (bandera)
            {               
                case 1:

                    //Click btn Aplicar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 944, 253);
                    desktopSession.Mouse.Click(null);

                    break;
                case 2:

                    //Click btn Cancelar
                    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 968, 253);
                    desktopSession.Mouse.Click(null);

                    break;

            }


        }



        static public void CrudDet2(WindowsDriver<WindowsElement> desktopSession, List<string> crudDet2)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 118, 327);
            desktopSession.Mouse.DoubleClick(null);
            desktopSession.Keyboard.SendKeys(crudDet2[0]);
            Thread.Sleep(2000);
;
        }
        
    }
}


