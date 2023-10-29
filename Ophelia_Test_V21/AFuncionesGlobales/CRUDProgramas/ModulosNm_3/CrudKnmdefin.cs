using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmdefin
    {
        public static void CrudKNmdefin(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                ElementList[3].SendKeys(data[0]);
                Thread.Sleep(500);
                ElementList[2].SendKeys(data[1]);
            }
            else if (Tipo == 2)
            {
                ElementList[2].Clear();
                Thread.Sleep(500);
                ElementList[2].SendKeys(data[2]);
            }
        }
        public static void CrudKNmdefinDetalle(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 66, 26);
            desktopSession.Mouse.Click(null);
            if (Tipo == 1)
            {
                ElementList[0].SendKeys(data[0]);
                Thread.Sleep(200);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                ElementList[0].SendKeys(data[1]);
                Thread.Sleep(200);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                ElementList[0].SendKeys(data[2]);
                Thread.Sleep(200);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                ElementList[0].SendKeys(data[3]);
            }
            else if (Tipo == 2)
            {
                for(int i = 0; i < 3; i++)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                }
                Thread.Sleep(200);
                ElementList[0].Clear();
                Thread.Sleep(500);
                ElementList[0].SendKeys(data[4]);
            }
        }
    }
}
