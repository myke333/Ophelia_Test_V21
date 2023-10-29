using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
    class CrudKnmpcate
    {
        public static void CrudKNmpcate(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            if (Tipo == 1)
            {
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Acumulados y Prenómina").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                PruebaCRUD.LupaDinamica2(desktopSession, data);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Maneja Motivo de Dejación").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                List<string> data1 = new List<string> { data[1] };
                PruebaCRUD.LupaDinamica2(desktopSession, data1);
            }
        }
        public static void CrudKNmpcateDetalle(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TKactusDBgrid");
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 66, 26);
            desktopSession.Mouse.Click(null);

            if (Tipo == 1)
            {
                ElementList[1].SendKeys(data[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList[1].SendKeys(data[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList[1].SendKeys(data[2]);
            }
            else if (Tipo == 2)
            {
                ElementList[1].Clear();
                ElementList[1].SendKeys(data[3]);
            }
        }
        public static void CrudKNmpcateOra(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            if (Tipo == 1)
            {
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Autoliquidación Mes Anterior").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Descuento de vacaciones").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                PruebaCRUD.LupaDinamica2(desktopSession, data);
                Thread.Sleep(1000);
            }
            else if (Tipo == 2)
            {
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Maneja Motivo de Dejación").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
            }
        }
        public static void CrudKNmpcateDetalleOra(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TKactusDBgrid");
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 66, 26);
            desktopSession.Mouse.Click(null);

            if (Tipo == 1)
            {
                ElementList[1].SendKeys(data[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList[1].SendKeys(data[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList[1].SendKeys(data[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList[1].SendKeys(data[3]);
            }
            else if (Tipo == 2)
            {
                ElementList[1].Clear();
                ElementList[1].SendKeys(data[4]);
            }
        }
    }
}
