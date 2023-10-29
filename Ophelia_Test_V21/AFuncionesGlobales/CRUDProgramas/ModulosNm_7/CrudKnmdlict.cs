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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_7
{
    class CrudKnmdlict : FuncionesVitales
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
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, string tipo, List<String> data, string file)
        {
            var Elementlist = desktopSession.FindElementsByClassName("TDBEdit");
            var Elementlist1 = desktopSession.FindElementsByName("Abrir");
            switch (tipo)
            {
                case ("0"):
                    Elementlist[9].SendKeys(data[0]);
                    Elementlist[9].SendKeys(OpenQA.Selenium.Keys.Tab);
                    Elementlist[4].SendKeys(data[1]);
                    Elementlist[4].SendKeys(OpenQA.Selenium.Keys.Tab);

                    Elementlist1[2].Click();
                    Elementlist1[2].SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    Elementlist1[2].SendKeys(OpenQA.Selenium.Keys.Tab);

                    Elementlist1[1].Click();
                    Elementlist1[1].SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    Elementlist1[1].SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    Elementlist1[1].SendKeys(OpenQA.Selenium.Keys.Tab);

                    Elementlist[3].SendKeys(data[2]);
                    Elementlist[3].SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
                case ("1"):
                    Elementlist[9].Clear();
                    Elementlist[9].SendKeys(data[3]);
                    Elementlist[9].SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
            }
        }
    }
}