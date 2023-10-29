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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_9
{
    class CrudKnmranma : FuncionesVitales
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

        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");

            /*for (int i = 0; i < btnTDVI.Count; i++)
            {
                btnTDVI[i].SendKeys(i.ToString());
            }
            Thread.Sleep(4000);*/
            if (bandera == 0)
            {
                btnTDVI[9].SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);
                btnTDVI[8].SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);
                btnTDVI[7].SendKeys(crudPrincipal[2]);
                Thread.Sleep(1000);
                btnTDVI[6].SendKeys(crudPrincipal[3]);
                Thread.Sleep(1000);
                btnTDVI[5].SendKeys(crudPrincipal[4]);
                Thread.Sleep(1000);
                btnTDVI[3].SendKeys(crudPrincipal[5]);
                Thread.Sleep(1000);
                btnTDVI[2].SendKeys(crudPrincipal[6]);
                Thread.Sleep(1000);
             }
            else
            {
                btnTDVI[5].Clear();
                btnTDVI[5].SendKeys(crudPrincipal[7]);
                Thread.Sleep(1000);
            }

        }

        static public void AgregarCrud1(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet1)
        {
            var btnTDVIgrid = desktopSession.FindElementsByClassName("TDBGrid");

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 350, 30);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                WindowsDriver<WindowsElement> rootSesison = null;
                rootSesison = PruebaCRUD.RootSession();
                Thread.Sleep(2000);
                rootSesison.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 410, 30);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                rootSesison = PruebaCRUD.RootSession();
                Thread.Sleep(2000);
                rootSesison.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);

                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 430, 30);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudDet1[0]);
            } 
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 430, 30);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudDet1[1]);
            }
        }

    }
}
