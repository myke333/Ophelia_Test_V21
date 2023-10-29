using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
    class CrudKnmpbper : PruebaCRUD
    {
        public static void CRUDKnmpbper(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList2 = desktopSession.FindElementsByClassName("TDBCheckBox");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                ElementList[8].SendKeys(data[0]);
                ElementList[4].SendKeys(data[1]);
                ElementList2[0].Click();
            }
            else
            {
                ElementList[2].Clear();
                ElementList[2].SendKeys(data[2]);
            }
        }

        public static void CRUDDetalle(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TKactusDBgrid");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 50, 11);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[3]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[4]);
            }
            else
            {
                desktopSession.Keyboard.SendKeys(data[5]);
            }
        }
    }
}
