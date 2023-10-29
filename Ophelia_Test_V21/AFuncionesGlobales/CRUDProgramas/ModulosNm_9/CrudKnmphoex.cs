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
    class CrudKnmphoex : FuncionesVitales
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
            }*/

            if (bandera == 0)
            {
                btnTDVI[4].SendKeys(crudPrincipal[0]);
                btnTDVI[3].SendKeys(crudPrincipal[1]);
                btnTDVI[2].SendKeys(crudPrincipal[2]);
            }
            else
            {
                //btnTDVI[4].Clear();
                btnTDVI[3].Clear();
                //btnTDVI[2].Clear();
                //btnTDVI[4].SendKeys("394");
                btnTDVI[3].SendKeys(crudPrincipal[3]);
                //btnTDVI[2].SendKeys("30/10/2021");
            }

        }

        static public void AgregarCrud1(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet1)
        {
            var btnTDVIgrid = desktopSession.FindElementsByClassName("TDBGrid");

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 44,30 );
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                WindowsDriver<WindowsElement> rootSession = null;
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TCaptura");
                Thread.Sleep(4500);
                rootSession.FindElementByName("Aceptar").Click();
                //desktopSession.Keyboard.SendKeys("1");
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 319, 30);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(3000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmGnArbol");
                Thread.Sleep(4000);
                var Tfrmgnarbol = rootSession.FindElementsByClassName("TVirtualStringTree");
                Thread.Sleep(4000);
                rootSession.Mouse.MouseMove(Tfrmgnarbol[0].Coordinates, 109, 8);
                Thread.Sleep(2000);
                rootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                rootSession.Mouse.Click(null);
                Thread.Sleep(2000);
                rootSession.FindElementByName("Aceptar").Click();
                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 629, 30);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudDet1[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {

                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 629, 30);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudDet1[3]);

            }

        }

        static public void ClickButton(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TDBGrid");

            if (bandera == 0)
            {
                desktopSession.FindElementByName("Compensatorios").Click();
            }
            else if (bandera == 1)
            {
                desktopSession.FindElementByName("Sin Horas Extras").Click();
            }
            else
            {
                desktopSession.FindElementByName("Horas Extras").Click();
            }

        }

        static public void AgregarCrud2(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet2)
        {
            var btnTDVIgrid = desktopSession.FindElementsByClassName("TDBGrid");

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 53, 28);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                WindowsDriver<WindowsElement> rootSession = null;
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TCaptura");
                var TKactusDBgrid = rootSession.FindElementsByClassName("TKactusDBgrid");
                Thread.Sleep(3000);
                rootSession.Mouse.MouseMove(TKactusDBgrid[0].Coordinates, 111, 70);
                Thread.Sleep(3000);
                rootSession.Mouse.Click(null);
                rootSession.FindElementByName("Aceptar").Click();
                Thread.Sleep(3000);
                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 339, 31);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudDet2[1]);
                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 435, 31);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudDet2[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {

                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 435, 31);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudDet2[3]);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }

        }

        static public void AgregarCrud3(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet3)
        {
            var btnTDVIgrid = desktopSession.FindElementsByClassName("TDBGrid");

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 54, 29);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(3000);
                WindowsDriver<WindowsElement> rootSession = null;
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TCaptura");
                var TKactusDBgrid = rootSession.FindElementsByClassName("TKactusDBgrid");
                Thread.Sleep(3000);
                rootSession.Mouse.MouseMove(TKactusDBgrid[0].Coordinates, 73, 174);
                Thread.Sleep(4000);
                rootSession.Mouse.Click(null);
                rootSession.FindElementByName("Aceptar").Click();
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 338, 33);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmGnArbol");
                Thread.Sleep(4000);
                var Tfrmgnarbol = rootSession.FindElementsByClassName("TVirtualStringTree");
                Thread.Sleep(3000);
                rootSession.Mouse.MouseMove(Tfrmgnarbol[0].Coordinates, 83, 47);
                Thread.Sleep(2000);
                rootSession.Mouse.Click(null);
                Thread.Sleep(2000);
                rootSession.FindElementByName("Aceptar").Click();
            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVIgrid[0].Coordinates, 338, 33);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                WindowsDriver<WindowsElement> rootSession = null;
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmGnArbol");
                Thread.Sleep(4000);
                var Tfrmgnarbol = rootSession.FindElementsByClassName("TVirtualStringTree");
                Thread.Sleep(3000);
                rootSession.Mouse.MouseMove(Tfrmgnarbol[0].Coordinates, 87, 25);
                Thread.Sleep(2000);
                rootSession.Mouse.Click(null);
                Thread.Sleep(2000);
                rootSession.FindElementByName("Aceptar").Click();
            }

        }

    }

    
}
