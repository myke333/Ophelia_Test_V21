using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System.Drawing;
using OpenQA.Selenium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosCo
{
    class CrudKcorcoen
    {
        public static void KCoRcoenCRUD(WindowsDriver<WindowsElement> desktopSession, int tabs, string radioName, List<string> lupasData) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            List<string> pageNames = new List<string>() { "Encuesta Base", "Encuesta de Comparación 1" };
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> desplegables = desktopSession.FindElementsByClassName("TKsCmbCodigo");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> radioButtons = desktopSession.FindElementsByClassName("TGroupButton");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> buttons = desktopSession.FindElementsByClassName("TButton");

            //Agregar modulos
            desktopSession.Mouse.MouseMove(desplegables[1].Coordinates);
            Thread.Sleep(1000);
            desktopSession.Mouse.Click(null);
            for (int i = 0; i < tabs; i++) {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                Thread.Sleep(700);
            }
            Thread.Sleep(1000);
            desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            //Elegir Reporte
            foreach (var rad in radioButtons) {
                if (rad.Text == radioName) {
                    desktopSession.Mouse.MouseMove(rad.Coordinates);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(800);
                }
            }

            //LupaPestañas
            PruebaCRUD.LupaDinamica(desktopSession, lupasData);
            Thread.Sleep(1000);
            newFunctions_4.changeWindow(desktopSession, "Encuesta de Comparación 1");
            List<string> lupasData2 = new List<string>() { lupasData[1] };
            PruebaCRUD.LupaDinamica(desktopSession, lupasData2);
            Thread.Sleep(2000);
            rootSession = newFunctions_4.RootSessionNew();
            Thread.Sleep(2000);
            WebDriverWait error = new WebDriverWait(rootSession, TimeSpan.FromSeconds(3));
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowLeft);
            Thread.Sleep(1000);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);


            //Presionar boton de datos
            foreach (var button in buttons) {
                if (button.Text == "Datos") {
                    desktopSession.Mouse.MouseMove(button.Coordinates);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                }
            }
        }
        //ClickWindow
        static public void ClickWindow (WindowsDriver<WindowsElement> desktopSession)
        {
            var btnTDVI = desktopSession.FindElementsByClassName("FrameGrabHandle");
            //TPanel
            desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 525, 17);
            Thread.Sleep(2000);
            desktopSession.Mouse.Click(null);
        }
    }
}
