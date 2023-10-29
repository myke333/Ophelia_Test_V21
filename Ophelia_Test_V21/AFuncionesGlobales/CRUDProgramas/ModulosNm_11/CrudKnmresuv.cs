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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_11
{
    class CrudKnmresuv : FuncionesVitales
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



        static public void AgregarDatos(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 375, 73);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys("1");
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);

            }

            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 701, 410);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 790, 263);
                desktopSession.Mouse.Click(null);

            }



            //Thread.Sleep(1000);




        }

        //static public void AceptarPreliminar(WindowsDriver<WindowsElement> desktopSession)
        //{
        //    Thread.Sleep(1500);
        //    WindowsDriver<WindowsElement> RootSession = null;
        //    RootSession = PruebaCRUD.RootSession();
        //    RootSession = ReloadSession(RootSession, "TFrmInfoP50");

        //    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

        //    //var btnTDVI2 = RootSession.FindElementsByClassName("TMessageForm");
        //    //Thread.Sleep(1500);
        //    //RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 254, 92);
        //    //RootSession.Mouse.Click(null);
        //    //Thread.Sleep(1500);


        //}

    }
}

