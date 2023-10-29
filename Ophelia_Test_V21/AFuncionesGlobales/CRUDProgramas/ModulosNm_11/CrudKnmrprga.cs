﻿using OpenQA.Selenium.Appium;
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
    class CrudKnmrprga : FuncionesVitales
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

        static public void SeleccionDatos(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal)
        {
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");

            if (bandera == 0)
            {


                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 190, 42);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 350, 42);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 500, 42);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1000);
            }
        }

        static public void Click(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");

            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 415, 335);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(1500);
            newFunctions_4.ScreenshotNuevo("Click QBE Det 1", true, file);
            //desktopSession.Mouse.Click(null);
            //Thread.Sleep(1500);

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TKQBE");

            var btnTDVI = RootSession.FindElementsByClassName("TTabSheet");

            RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 600, 16);
            RootSession.Mouse.Click(null);
            Thread.Sleep(2000);

            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);
            newFunctions_4.ScreenshotNuevo("Registros QBE", true, file);

            WindowsDriver<WindowsElement> RootSession3 = null;
            RootSession3 = PruebaCRUD.RootSession();
            RootSession3.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);
        }


        public static void ClickButtonPreliminar(WindowsDriver<WindowsElement> desktopSession, int banderas, string file)
        {
            var btnTDVI0 = desktopSession.FindElementsByClassName("TScrollBox");

            desktopSession.Mouse.MouseMove(btnTDVI0[0].Coordinates, 697, 310);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession = ReloadSession(RootSession, "TFrmImpRep");

            //[@ClassName=\"TRadioGroup\"]/RadioButton[@ClassName=\"TGroupButton\

            var btnTDVI2 = RootSession.FindElementsByClassName("TRadioGroup");


            switch (banderas)
            {
                case 0:
                    //primer opcion preliminar
                    Thread.Sleep(2000);
                    //RootSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 61, 12);
                    //RootSession.Mouse.Click(null);
                    newFunctions_4.ScreenshotNuevo("Primera opcion del preliminar", true, file);
                    Thread.Sleep(2000);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    break;

                case 1:
                    ////Segunda opcion preliminar
                    Thread.Sleep(2000);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                    newFunctions_4.ScreenshotNuevo("Segunda opcion del preliminar", true, file);
                    Thread.Sleep(2000);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                    break;

                case 2:
                    //Tercera opcion preliminar
                    Thread.Sleep(2000);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                    newFunctions_4.ScreenshotNuevo("Tercera opcion del preliminar", true, file);
                    Thread.Sleep(2000);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                    break;

                case 3:
                    //Cuarta opcion preliminar
                    Thread.Sleep(2000);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                    newFunctions_4.ScreenshotNuevo("Cuarta opcion del preliminar", true, file);
                    Thread.Sleep(2000);
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    break;

                case 4:
                    //Dar enter
                    RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    break;
            }
          
           
        }
    }
}
