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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosEd
{
    class CurdKedpmoev : FuncionesVitales
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

        static public void Click(WindowsDriver<WindowsElement> desktopSession)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TGroupBox");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();


            desktopSession.FindElementByName("Aceptar").Click();

        }

        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TGroupBox");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");

            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 144, 46);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 136, 58);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            Thread.Sleep(2000);

            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 334, 46);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 326, 58);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(2000);
            Thread.Sleep(2000);

        }

        static public void Enter(WindowsDriver<WindowsElement> desktopSession)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TGroupBox");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();

            WindowsDriver<WindowsElement> RootSession = null;
            RootSession = PruebaCRUD.RootSession();
            RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

        }
    }
}
