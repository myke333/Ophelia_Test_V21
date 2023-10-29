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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmpreno
    {
        public static void KNMPrenoCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1),
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");
            List<int> fieldIndex = new List<int>() { 17, 16, 21, 15, 14 };

            if (flag == 0)
            {
                List<int> lupasOff = new List<int>() { 2, 3, 4 };
                newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, variables, lupasOff);
                Thread.Sleep(1000);
                for (int i = 0; i < fieldIndex.Count; i++) {
                    desktopSession.Mouse.MouseMove(editFields[fieldIndex[i]].Coordinates);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    editFields[fieldIndex[i]].Clear();
                    Thread.Sleep(1000);
                    desktopSession.Keyboard.SendKeys(variables[2 + i]);
                    Thread.Sleep(1000);
                }
            }
            else {
                List<int> lupasOff = new List<int>() { 0, 2, 3, 4 };
                List<string> editvalue = new List<string>() { variables[variables.Count - 1] };
                newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, editvalue, lupasOff);
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

            rootSession = newFunctions_4.RootSessionNew();
            rootSession = newFunctions_4.ReloadXSession(rootSession, "TFrmRep");
            Thread.Sleep(5000);
            WebDriverWait reporteWindow = new WebDriverWait(rootSession, TimeSpan.FromSeconds(4));


           
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> checkBoxes = rootSession.FindElementsByClassName("TCheckBox");
            reporteWindow.Until(res => checkBoxes[0].Displayed);

            //foreach (var but in radioButton) {
            //    rootSession.Mouse.MouseMove(but.Coordinates);
            //    Thread.Sleep(2000);
            //}

            foreach (var elem in checkBoxes)
            {
                if (elem.Text == buttonName && isDateNeccessary == false)
                {
                    rootSession.Mouse.MouseMove(elem.Coordinates);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                    newFunctions_4.ScreenshotNuevo(buttonName, true, file);
                    Thread.Sleep(2000);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(2000);

                    rootSession = newFunctions_4.RootSessionNew();
                    Thread.Sleep(5000);
                    reporteWindow = new WebDriverWait(rootSession, TimeSpan.FromSeconds(4));
                    System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> message = rootSession.FindElementsByClassName("TFrmNmMsjInst");
                    reporteWindow.Until(res => message[0].Displayed);

                    rootSession.Mouse.MouseMove(message[0].Coordinates, 282, 168);
                    Thread.Sleep(1000);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    rootSession = null;
                    rootSession = newFunctions_4.RootSessionNew();
                    Thread.Sleep(5000);
                    reporteWindow = new WebDriverWait(rootSession, TimeSpan.FromSeconds(4));
                    System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> menuPrem = rootSession.FindElementsByClassName("TFrmRep");
                    reporteWindow.Until(res => message[0].Displayed);
                    Thread.Sleep(3000);
                   
            
                    rootSession.Mouse.MouseMove(menuPrem[0].Coordinates, menuPrem[0].Size.Width/2, menuPrem[0].Size.Height-20);
                    Thread.Sleep(1000);
                    rootSession.Mouse.Click(null);
                    
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

                    rootSession = null;
                    Thread.Sleep(2000);
                    rootSession = newFunctions_4.RootSessionNew();
                    rootSession = newFunctions_4.ReloadXSession(rootSession, "TFrmRep");
                    Thread.Sleep(5000);
                    reporteWindow = new WebDriverWait(rootSession, TimeSpan.FromSeconds(4));


                    //ventana de reporte preliminar: "/Pane[@ClassName=\"#32769\"][@Name=\"Escritorio 1\"]/Window[@ClassName=\"TFrmBiProxExa\"][@Name=\"Reportes\"]"
                    //radio butons: /RadioButton[@ClassName=\"TGroupButton\"][@Name=\"Próximo Examen\"]"
                    //fechasfields: "/Pane[@ClassName=\"#32769\"][@Name=\"Escritorio 1\"]/Window[@ClassName=\"TFrmBiProxExa\"][@Name=\"Reportes\"]/Edit[@ClassName=\"TEdit\"]"
                    //boton Aceptar: "/Pane[@ClassName=\"#32769\"][@Name=\"Escritorio 1\"]/Window[@ClassName=\"TFrmBiProxExa\"][@Name=\"Reportes\"]/Button[@ClassName=\"TBitBtn\"][@Name=\"Aceptar\"]"

                    System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> cerrarButtons = rootSession.FindElementsByName("Cerrar");
                    reporteWindow.Until(res => cerrarButtons[0].Displayed);

                    rootSession.Mouse.MouseMove(cerrarButtons[0].Coordinates);
                    Thread.Sleep(1000);
                    rootSession.Mouse.Click(null);

                }
                else
                {
               
                }
            }
            return errorMessages;
        }


    }
}
