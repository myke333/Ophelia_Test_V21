﻿using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsoadpmc : FuncionesVitales
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






        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, String file)
        {
            List<string> errorMessages = new List<string>();
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            WindowsDriver<WindowsElement> rootSession = null;

            rootSession = RootSession();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            List<int> lupasOff = new List<int>() { 0, 1, 2 };


            if (tipo == 1)
            {

                {

                    PruebaCRUD.LupaDinamica(desktopSession, crudVariables);
                    ElementList[9].SendKeys(crudVariables[3]);
                    ElementList[12].SendKeys(crudVariables[4]);
                    


                }

            }

            else
            {

                ElementList[12].Clear();
                ElementList[12].SendKeys(crudVariables[5]);

            }
        }
    }
}