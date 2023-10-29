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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSl
{
    class CrudKslreqpe : FuncionesVitales
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

            //var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");
            //Debugger.Launch();
            //for (int i = 0; i < btnTDVI.Count; i++)
            //{
            //    btnTDVI[i].SendKeys(i.ToString());
            //}


            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 40, 40);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 426, 40);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 585, 40);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 585, 87);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 750, 40);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 750, 87);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 40, 95);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 620, 95);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                WindowsDriver<WindowsElement> RootSession2 = null;
                RootSession2 = PruebaCRUD.RootSession();
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                RootSession2.Keyboard.SendKeys(crudPrincipal[3]);
                Thread.Sleep(1000);
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 40, 143);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[4]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 40, 285);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 40, 324);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[10]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 310, 324);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 45, 556);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[6]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[7]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[8]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 680, 100);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[9]);
                Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 680, 100);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[10]);
                Thread.Sleep(1000);

                //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 101, 27);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                //Thread.Sleep(1500);

            }
        }

        static public void Click(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 115, -20);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 160, 78);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            if (bandera == 1)
            {

                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession = ReloadSession(RootSession, "TFrmNmFirMail");

                //var btnTDVI2 = RootSession.FindElementsByClassName("TTabSheet");

                var ElementList = RootSession.FindElementsByName("Cancelar");
                ElementList[0].Click();
                Thread.Sleep(2000);

                //RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 318, -54);
                //RootSession.Mouse.Click(null);
                //Thread.Sleep(2000);

            }
            if (bandera == 2)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 222, 78);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            if (bandera == 3)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 410, -20);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            if (bandera == 4)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 115, -20);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            if (bandera == 5)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 160, 78);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            if (bandera == 6)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 40, -20);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
        }

        static public void AgregarCrud1(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet)
        {

            //var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            //var btnTDVI = desktopSession.FindElementsByClassName("TDBGrid");
            var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");
            //Debugger.Launch();

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 40, 152);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet[0]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet[1]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet[2]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet[3]);
                Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 786, 152);//95
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet[4]);
                Thread.Sleep(1000);

            }
        }

        static public void AgregarCrud2(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet2)
        {

            //var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            //var btnTDVI = desktopSession.FindElementsByClassName("TDBGrid");
            var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");
            //Debugger.Launch();

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 128, 152);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet2[0]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet2[1]);
                Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 556, 152);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet2[2]);
                Thread.Sleep(1000);

            }
        }

        static public void AgregarCrud3(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet3)
        {

            //var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            //var btnTDVI = desktopSession.FindElementsByClassName("TDBGrid");
            var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");
            //Debugger.Launch();

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 120, 60);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet3[0]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 566, 60);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 566, 80);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 673, 60);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 673, 80);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 740, 60);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet3[1]);
                Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 740, 60);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet3[2]);
                Thread.Sleep(1000);

            }
        }

        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {
            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);

        }
    }
}