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
using System.IO;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsoseyde
    {
        public static void KSoseydeCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1),
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");

            if (flag == 0)
            {
                List<int> fieldIndex = new List<int>() { 10, 12, 13};
                for (int i = 0; i < fieldIndex.Count; i++) {
                    editFields[fieldIndex[i]].Click();
                    editFields[fieldIndex[i]].Clear();
                    Thread.Sleep(1000);
                    editFields[fieldIndex[i]].SendKeys(variables[i]);
                    Thread.Sleep(1000);
                }
            }
            else {
                Thread.Sleep(1000);
                editFields[10].Click();
                editFields[10].Clear();
                Thread.Sleep(1000);
                editFields[10].SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
            }

        }

        public static void KSoseydeCRUDDet1(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1),
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> grids = desktopSession.FindElementsByClassName("TDBGrid");
            if (flag == 0)
            {
                desktopSession.Mouse.MouseMove(grids[0].Coordinates, 20, grids[0].Size.Height/4);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                rootSession = PruebaCRUD.RootSession();
                Thread.Sleep(4000);
                System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> dots = rootSession.FindElementsByClassName("TDBGridInplaceEdit");
                rootSession.Mouse.MouseMove(dots[1].Coordinates, dots[1].Size.Width-10, dots[1].Size.Height/2);
                Thread.Sleep(1000);
                rootSession.Mouse.Click(null);
                Thread.Sleep(1000);
                rootSession = null;
                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                var campo = rootSession.FindElementsByClassName("TEdit");
                ////Debugger.Launch();
                rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                rootSession.Mouse.Click(null);
                rootSession.Keyboard.SendKeys(variables[0]);
                //Aqui se puede enviar un valor

                var btn = rootSession.FindElementsByClassName("TBitBtn");
                ////Debugger.Launch();
                foreach (var boton in btn)
                {
                    if (boton.Text == "Aceptar")
                    {
                        rootSession.Mouse.MouseMove(boton.Coordinates, 10, 10);
                        rootSession.Mouse.Click(null);
                    }
                }

                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(grids[0].Coordinates, 20, grids[0].Size.Height / 4);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);

                for (int i = 1; i < variables.Count-1; i++) {
                    desktopSession.Keyboard.SendKeys(variables[i]);
                    Thread.Sleep(1000);
                    if (i < variables.Count - 2) desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                }

            }
            else {
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
            }

        }

        public static void KSoseydeCRUDDet2(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1),
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> grids = desktopSession.FindElementsByClassName("TDBGrid");

            if (flag == 0)
            {
                desktopSession.Mouse.MouseMove(grids[0].Coordinates, 20, grids[0].Size.Height / 4);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                List<int> exepcions = new List<int>() { 2, 4, 5 };
                for (int i = 0; i < variables.Count-1; i++) {
                    desktopSession.Keyboard.SendKeys(variables[i]);
                    Thread.Sleep(1000);
                    if (i < variables.Count - 2) desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(1000);
                    if (exepcions.Contains(i)) desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(1000);
                }

            }
            else {
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
            }
        }


        public static List<string> ReportePreliminarIfnotOptions(WindowsDriver<WindowsElement> desktopSession, string ruta, string file, string buttonName, bool isDateNeccessary, List<string> dates = null)
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
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            
            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "TFrmOpcRepo");
            Thread.Sleep(5000);
            WebDriverWait reporteWindow = new WebDriverWait(rootSession, TimeSpan.FromSeconds(4));


            //ventana de reporte preliminar: "/Pane[@ClassName=\"#32769\"][@Name=\"Escritorio 1\"]/Window[@ClassName=\"TFrmBiProxExa\"][@Name=\"Reportes\"]"
            //radio butons: /RadioButton[@ClassName=\"TGroupButton\"][@Name=\"Próximo Examen\"]"
            //fechasfields: "/Pane[@ClassName=\"#32769\"][@Name=\"Escritorio 1\"]/Window[@ClassName=\"TFrmBiProxExa\"][@Name=\"Reportes\"]/Edit[@ClassName=\"TEdit\"]"
            //boton Aceptar: "/Pane[@ClassName=\"#32769\"][@Name=\"Escritorio 1\"]/Window[@ClassName=\"TFrmBiProxExa\"][@Name=\"Reportes\"]/Button[@ClassName=\"TBitBtn\"][@Name=\"Aceptar\"]"

            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> radioButton = rootSession.FindElementsByClassName("TGroupButton");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> aceptButton = rootSession.FindElementsByName("OK");
            reporteWindow.Until(res => aceptButton[0].Displayed);

            //foreach (var but in radioButton) {
            //    rootSession.Mouse.MouseMove(but.Coordinates);
            //    Thread.Sleep(2000);
            //}

            foreach (var elem in radioButton)
            {
                if (elem.Text == buttonName && isDateNeccessary == false)
                {
                    rootSession.Mouse.MouseMove(elem.Coordinates);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo(buttonName, true, file);
                    rootSession.Mouse.MouseMove(aceptButton[0].Coordinates);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(3000);
                    try
                    {
                        errors = newFunctions_5.GenerarReportes("Preliminar", file, ruta);
                        if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                    }
                    catch
                    {
                        newFunctions_4.ScreenshotNuevo("Error en Preliminar", true, file);
                        Thread.Sleep(1000);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    }
                }
                else
                {
                    if (elem.Text == buttonName && isDateNeccessary == true && dates != null)
                    {
                        rootSession.Mouse.MouseMove(elem.Coordinates);
                        rootSession.Mouse.Click(null);
                        System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> dateField = rootSession.FindElementsByClassName("TEdit");
                        int indexDates = dates.Count - 1;
                        foreach (var field in dateField)
                        {

                            rootSession.Mouse.MouseMove(field.Coordinates);
                            rootSession.Mouse.Click(null);
                            field.Clear();
                            rootSession.Keyboard.SendKeys(dates[indexDates]);
                            Thread.Sleep(1000);
                            indexDates--;
                        }

                        rootSession.Mouse.MouseMove(aceptButton[0].Coordinates);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(3000);
                        try
                        {
                            errors = newFunctions_5.GenerarReportes("Preliminar", file, ruta);
                            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                        }
                        catch
                        {
                            newFunctions_4.ScreenshotNuevo("Error en Preliminar", true, file);
                            Thread.Sleep(1000);
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        }

                    }
                }
            }
            return errorMessages;
        }

    }
}
