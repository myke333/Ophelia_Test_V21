using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmpadia : PruebaCRUD
    {
        public static void CRUDKnmpadia(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                ElementList[4].SendKeys(data[0]);
                ElementList[2].SendKeys(data[1]);
                ElementList[3].SendKeys(data[2]);
                ElementList[5].SendKeys(data[3]);
                desktopSession.FindElementByName("Pasivos Laborales").Click();
                rootSession = RootSession();
                var ElementList2 = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList2[2].SendKeys(data[4]);
                desktopSession.FindElementByName("Datos de la Resolución").Click();
            }
            else
            {
                ElementList[2].Clear();
                ElementList[2].SendKeys(data[5]);
            }
        }
    }
}
