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
    class CrudKnmcesan : FuncionesVitales
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

        static public void Lupa(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            //var btnTDVI = desktopSession.FindElementsByClassName("TGroupBox");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();

            Thread.Sleep(2000);

            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 692, 33);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000); 
            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TKQBE");

            var btnTDVI = RootSession.FindElementsByClassName("TStringGrid");

            RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 302, 34);
            RootSession.Mouse.Click(null);
            RootSession.Keyboard.SendKeys(crudPrincipal[0]);
            Thread.Sleep(2000);

            var btnTDVI3 = RootSession.FindElementsByClassName("TPanel");

            RootSession.Mouse.MouseMove(btnTDVI3[0].Coordinates, 55, 15);
            RootSession.Mouse.Click(null);

            Thread.Sleep(2000);

            Enter(desktopSession);

        }

        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {

            //var btnTDVI = desktopSession.FindElementsByClassName("TGroupBox");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();

            if (bandera == 0)
            {
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(1000);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);

                //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 281, 317);
                //desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);


                //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 460, 303);
                //desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(2000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 435, 301);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(2000);
            }

        }

        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {

            //var btnTDVI = desktopSession.FindElementsByClassName("TGroupBox");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
        }

        static public void Excel(WindowsDriver<WindowsElement> desktopSession)
        {

            //var btnTDVI = desktopSession.FindElementsByClassName("TGroupBox");
         
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();

            //if (bandera == 0)
            //{
                var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 575, 24);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);

                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            ////desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            ////Thread.Sleep(1000);
            ////desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            ////Thread.Sleep(1000);
            //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            //Thread.Sleep(1000);
            //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            //Thread.Sleep(1000);
            //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            //Thread.Sleep(1000);
            //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            //Thread.Sleep(1000);
            //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            //Thread.Sleep(1000);
            //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            //Thread.Sleep(1000);
            //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            //Thread.Sleep(1000);
            //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            //Thread.Sleep(1000);
            //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            //Thread.Sleep(1000);
            //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            //Thread.Sleep(1000);
            //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            //Thread.Sleep(1000);
            //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            //Thread.Sleep(1000);
            //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            //Thread.Sleep(1000);
            //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            //Thread.Sleep(1000);

            ////desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 281, 317);
            ////desktopSession.Mouse.Click(null);
            //desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
            //Thread.Sleep(1000);
            //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            //Thread.Sleep(1000);


            ////desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 460, 303);
            ////desktopSession.Mouse.Click(null);
            //desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
            //Thread.Sleep(2000);

            //}
            //else
            //{
            //    var btnTDVI4 = desktopSession.FindElementsByClassName("TForm");



            //    //Coordenadas boton qbe2
            //    //string filterqb = "13";



            //    desktopSession.Mouse.MouseMove(btnTDVI4[1].Coordinates, 242, 51);
            //    desktopSession.Mouse.Click(null);
            //    Thread.Sleep(2000);

            //}

        }


        static public void Ventana(WindowsDriver<WindowsElement> desktopSession)
        {




            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");



            //Coordenadas boton qbe2
            //string filterqb = "13";



            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 887, 322);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);

        }

        }
    }
