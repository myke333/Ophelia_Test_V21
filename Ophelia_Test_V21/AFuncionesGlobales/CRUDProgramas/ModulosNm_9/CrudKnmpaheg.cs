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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_9
{
    class CrudKnmpaheg : FuncionesVitales
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
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();
            //for (int i = 0; i < btnTDVI.Count; i++)
            //{
            //    btnTDVI[i].SendKeys(i.ToString());
            //}

            if (bandera == 0)
            {

                btnTDVI[12].SendKeys(crudPrincipal[1]);
                btnTDVI[10].SendKeys(crudPrincipal[2]);
                btnTDVI[8].SendKeys(crudPrincipal[3]);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 25, 96);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 152, 96);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 278, 96);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 391, 96);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 524, 96);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 634, 96);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[5]);
                Thread.Sleep(1500);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 25, 149);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[6]);
                Thread.Sleep(1500);
                btnTDVI[6].SendKeys(crudPrincipal[7]);
                btnTDVI[34].SendKeys(crudPrincipal[8]);
                btnTDVI[33].SendKeys(crudPrincipal[9]);
                btnTDVI[32].SendKeys(crudPrincipal[10]);
                btnTDVI[31].SendKeys(crudPrincipal[11]);
                btnTDVI[30].SendKeys(crudPrincipal[12]);
                btnTDVI[29].SendKeys(crudPrincipal[13]);
                btnTDVI[28].SendKeys(crudPrincipal[14]);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 389, 356);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(crudPrincipal[15]);
                Thread.Sleep(1500);
            }
            else
            {
                btnTDVI[8].Clear();
                btnTDVI[8].SendKeys(crudPrincipal[4]);

                //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 25, 149);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[6]);
            }

        }
    }
}
