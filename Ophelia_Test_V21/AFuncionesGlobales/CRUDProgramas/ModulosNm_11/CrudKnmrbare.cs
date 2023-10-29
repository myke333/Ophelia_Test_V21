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
    class CrudKNmRbare : FuncionesVitales
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

        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera==0)
            {
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 69, 66);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys("01/10/2020");
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            }
            else if (bandera==1)
            {
                desktopSession.Keyboard.SendKeys("1");
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys("76");
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys("76");
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(2500);
            }
            else
            {
                Thread.Sleep(2500);
                desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 710, 590);
                desktopSession.Mouse.Click(null);
            }
            



            //if (bandera == 0)
            //{
            //    //desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 66, 82);
            //    //desktopSession.Mouse.DoubleClick(null);
            //    //desktopSession.Keyboard.SendKeys("01/01/2020");

            //    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 162, 76);
            //    desktopSession.Mouse.Click(null);
            //    Thread.Sleep(1500);
            //    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 718, 48);
            //    desktopSession.Mouse.DoubleClick(null);
            //    Thread.Sleep(1800);
            //    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 740, 55);
            //    desktopSession.Mouse.DoubleClick(null);
            //    Thread.Sleep(2000);
            //    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 674, 212);
            //    desktopSession.Mouse.Click(null);
            //}
            //else if(bandera == 1)
            //{
            //    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 69, 108);
            //    desktopSession.Mouse.Click(null);
            //    desktopSession.Keyboard.SendKeys("1");
            //    Thread.Sleep(1500);
            //    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 564, 190);
            //    desktopSession.Mouse.Click(null);
            //    Thread.Sleep(2500);
            //    WindowsDriver<WindowsElement> RootSession2 = null;
            //    RootSession2 = PruebaCRUD.RootSession();
            //    RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            //    Thread.Sleep(2500);
            //    desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 568, 258);
            //    desktopSession.Mouse.Click(null);
            //    Thread.Sleep(1500);
            //    RootSession2 = PruebaCRUD.RootSession();
            //    RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            //    Thread.Sleep(2000);
            //}
            //else
            //{

            //}
        }

        static public void qbe2(WindowsDriver<WindowsElement> desktopSession)
        {



            var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");

            //Coordenadas boton qbe2
            //string filterqb = "13";

            //protegido

            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 832, 593);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);



            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TKQBE");



            var btnTDVI2 = RootSession.FindElementsByClassName("TTabSheet");



            RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 255, 40);
            RootSession.Mouse.Click(null);
            Thread.Sleep(2000);
            RootSession.Mouse.Click(null);
            Thread.Sleep(1000);
            RootSession.Keyboard.SendKeys("2");
            Thread.Sleep(1000);
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);





            //WindowsDriver<WindowsElement> RootSession2 = null;
            //RootSession2 = PruebaCRUD.RootSession();
            //Thread.Sleep(500);
            //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 247, 91);



            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(1000);
            //desktopSession.Keyboard.SendKeys(filterqb);
            //Thread.Sleep(500);
            //RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);



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
