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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_9
{
    class CrudKnmCaaut : FuncionesVitales
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

                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }
        static public void AgregarCrud(WindowsDriver<WindowsElement> desktopSession/* int bandera, List<string> crudPrincipal*/)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("TDBEdit");

            for (int i = 0; i < btnTDVI.Count; i++)
            {
                btnTDVI[i].SendKeys(i.ToString());
            }
            //List<string> crudLupa = new List<string>() { "", "", crudPrincipal[3], "" };

            //if (bandera == 0)
            //{
            //    LupaDinamica(desktopSession, crudLupa);
            //    btnTDVI[6].SendKeys(crudPrincipal[0]);
            //    btnTDVI[7].SendKeys(crudPrincipal[1]);
            //    btnTDVI[9].SendKeys(crudPrincipal[2]);
            //    btnTDVI[17].SendKeys(crudPrincipal[4]);
            //    btnTDVI[13].SendKeys(crudPrincipal[5]);
            //    btnTDVI[14].SendKeys(crudPrincipal[6]);
            //    btnTDVI[2].SendKeys(crudPrincipal[7]);
            //    btnTDVI[12].SendKeys(crudPrincipal[8]);
            //    btnTDVI[11].SendKeys(crudPrincipal[9]);
            //    btnTDVI[10].SendKeys(crudPrincipal[10]);


            //}
            //else
            //{
            //    btnTDVI[14].Clear();
            //    btnTDVI[14].SendKeys(crudPrincipal[11]);
            //}

        }
    }
}
