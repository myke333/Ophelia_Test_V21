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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmmlsec : FuncionesGlobales
    {
        public static void KNmMlseCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {

            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> radButtons = desktopSession.FindElementsByClassName("TGroupButton");


            if (flag == 0)
            {

                PruebaCRUD.LupaDinamica(desktopSession, variables);

                desktopSession.Mouse.MouseMove(editFields[9].Coordinates);
                desktopSession.Mouse.Click(null);
                editFields[9].Clear();
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(variables[variables.Count - 2]);
                Thread.Sleep(1000);

                foreach (var elem in radButtons)
                {
                    if (elem.Text == variables[variables.Count - 3])
                    {
                        desktopSession.Mouse.MouseMove(elem.Coordinates);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);
                    }
                }
            }
            else {

                foreach (var elem in radButtons)
                {
                    if (elem.Text == variables[variables.Count - 1])
                    {
                        desktopSession.Mouse.MouseMove(elem.Coordinates);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);
                    }
                }

            }


        }


        public static void ReportePreliminarKNmMlsec(WindowsDriver<WindowsElement> desktopSession, string ruta, List<string> buttonNames, string file)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WebDriverWait printScreen;
            WindowsElement printerField;
            WindowsDriver<WindowsElement> rootSession = null;


            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            ////////////////////////////////////////////////


            Thread.Sleep(2000);
            List<Point> coordenadas = FuncionesGlobales.coordinatesFinder(desktopSession, 227, 117, 0);
            int coordX = coordenadas[0].X;
            int coordY = coordenadas[0].Y;
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            Thread.Sleep(1000);
            try
            {
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                desktopSession.Mouse.Click(null);
                newFunctions_4.ScreenshotNuevo(buttonNames[0], true, file);
                WindowsDriver<WindowsElement> rootSession1 = null;
                rootSession1 = RootSession();
                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }

            rootSession = newFunctions_4.RootSessionNew();
            Thread.Sleep(5000);
            WebDriverWait reporteWindow = new WebDriverWait(rootSession, TimeSpan.FromSeconds(4));


            //ventana de reporte preliminar: "/Pane[@ClassName=\"#32769\"][@Name=\"Escritorio 1\"]/Window[@ClassName=\"TFrmBiProxExa\"][@Name=\"Reportes\"]"
            //radio butons: /RadioButton[@ClassName=\"TGroupButton\"][@Name=\"Próximo Examen\"]"
            //fechasfields: "/Pane[@ClassName=\"#32769\"][@Name=\"Escritorio 1\"]/Window[@ClassName=\"TFrmBiProxExa\"][@Name=\"Reportes\"]/Edit[@ClassName=\"TEdit\"]"
            //boton Aceptar: "/Pane[@ClassName=\"#32769\"][@Name=\"Escritorio 1\"]/Window[@ClassName=\"TFrmBiProxExa\"][@Name=\"Reportes\"]/Button[@ClassName=\"TBitBtn\"][@Name=\"Aceptar\"]"

            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> radioButton = rootSession.FindElementsByClassName("TGroupButton");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> aceptButton = rootSession.FindElementsByName(" Aceptar");
            reporteWindow.Until(res => aceptButton[0].Displayed);


            foreach (var but in radioButton) {
                if(but.Text == buttonNames[0] || but.Text == buttonNames[buttonNames.Count-1])
                rootSession.Mouse.MouseMove(but.Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(2000);
            }

            rootSession.Mouse.MouseMove(aceptButton[0].Coordinates);
            rootSession.Mouse.Click(null);
            Thread.Sleep(2000);

            try
            {
                errors = newFunctions_5.GenerarReportes(buttonNames[0], file, ruta);
            }
            catch (Exception e) {
                newFunctions_4.ScreenshotNuevo("Error al generar Preliminar", true, file);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(2000);
            }


        }


    }
}
