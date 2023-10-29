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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_10
{
    class CrudKrlhvext : FuncionesVitales
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
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 91, 35);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 60, 355);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 507, 201);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1000);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);

                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(2000);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);

                //WindowsDriver<WindowsElement> RootSession2 = null;
                //RootSession2 = PruebaCRUD.RootSession();
                //RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(1000);
                //RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(1000);
                //RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(1000);
                //RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(1000);
                //RootSession2.Keyboard.SendKeys(crudPrincipal[2]);
                //Thread.Sleep(1000);
                //RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 507, 201);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
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
            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 407, 66);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            if (bandera == 1)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 336, 66);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            if (bandera == 2)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 115, 95);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            if (bandera == 3)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 223, 66);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            if (bandera == 4)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 50, 95);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

            }
            if (bandera == 5)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 50, 66);
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
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 167, 28);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
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
                RootSession2.Keyboard.SendKeys(crudDet[0]);
                Thread.Sleep(1000);
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 226, 28);
                desktopSession.Mouse.Click(null);
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
                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 499, 272);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 926, 28);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet[4]);
                Thread.Sleep(1000);

            }
        }

        static public void AgregarCrud2(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet2)
        {

            var btnTDVI2 = desktopSession.FindElementsByClassName("TDBMemo"); 
            //var btnTDVI2 = desktopSession.FindElementsByClassName("TFrmDesComun");
            var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");
            //Debugger.Launch();

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 75, 63);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet2[0]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet2[1]);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 665, 63);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

                WindowsDriver<WindowsElement> RootSession2 = null;
                RootSession2 = PruebaCRUD.RootSession();
                RootSession2 = ReloadSession(RootSession2, "TFrmDesComun");
                RootSession2.Keyboard.SendKeys(crudDet2[2]);
                Thread.Sleep(1000);
                //RootSession2.Mouse.MouseMove(btnTDVI2[0].Coordinates, 360, 310);
                //RootSession2.Mouse.Click(null);
                //Thread.Sleep(1000);

                //var ElementList = desktopSession.FindElementByXPath("xpath_LeftClickButtonAceptar_57_10");
                var ElementList = RootSession2.FindElementsByName("Aceptar");
                ElementList[0].Click();
                Thread.Sleep(2000);

                //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 499, 272);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 204, 63);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudDet2[3]);
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
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 75, 63);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet3[0]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 232, 63);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 232, 83);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 298, 63);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 298, 83);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 373, 63);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 373, 83);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 440, 63);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 440, 83);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);


            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 75, 63);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet3[1]);
                Thread.Sleep(1000);

            }
        }

        static public void AgregarCrud4(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudDet4)
        {

            //var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");
            //var btnTDVI = desktopSession.FindElementsByClassName("TDBGrid");
            var btnTDVI = desktopSession.FindElementsByClassName("TTabSheet");
            //Debugger.Launch();

            if (bandera == 0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 124, 63);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudDet4[0]);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 316, 63);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                


            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 124, 63);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(crudDet4[1]);
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