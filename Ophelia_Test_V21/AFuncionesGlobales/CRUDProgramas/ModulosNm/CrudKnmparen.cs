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
    class CrudKnmparen
    {
        public static void KNmParenCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> lupasConceptos, List<string> lupasReintegro, List<string> lupasEncargos, string motor, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            //System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> pageControls = desktopSession.FindElementsByClassName("TPageControl");
            List<string> pagesNames = new List<string>() { "Conceptos", "Conceptos de Reintegro", "Ausentismos", "Encargos Tipo Comisión", "Otros Parámetros" };

            if (flag == 0)
            {
                procesoConceptos(desktopSession, lupasConceptos);
                newFunctions_4.changeWindow(desktopSession, pagesNames[1]);
                Thread.Sleep(1000);
                procesoReintegro(desktopSession, lupasReintegro);
                newFunctions_4.changeWindow(desktopSession, pagesNames[2]);
                Thread.Sleep(1000);
                procesoAusentismos(desktopSession);
                newFunctions_4.changeWindow(desktopSession, pagesNames[3]);
                Thread.Sleep(1000);
                procesoEncargos(desktopSession, lupasEncargos);
                newFunctions_4.changeWindow(desktopSession, pagesNames[4]);
                Thread.Sleep(1000);
                procesoParametros(desktopSession, motor);
                newFunctions_4.changeWindow(desktopSession, pagesNames[0]);
                Thread.Sleep(1000);
            }
            else {
                List<int> lupasOff = new List<int>() { 0, 2, 3 };
                List<string> editValue = new List<string>() { lupasConceptos[lupasConceptos.Count - 1] };
                newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, editValue, lupasOff);
            }

        }

        private static void procesoConceptos(WindowsDriver<WindowsElement> desktopSession, List<string> lupasConceptos) {
            List<int> lupasOff = new List<int>() { 2, 3 };
            newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, lupasConceptos, lupasOff);
        }

        private static void procesoReintegro(WindowsDriver<WindowsElement> desktopSession, List<string> lupasReintegro)
        {
            List<int> lupasOff = new List<int>() { 2, 3 };
            newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, lupasReintegro, lupasOff);
        }

        private static void procesoAusentismos(WindowsDriver<WindowsElement> desktopSession)
        {
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> boxes = desktopSession.FindElementsByClassName("TDBCheckBox");
            foreach (var elem in boxes) {
                if (elem.Text == "Permite Ingresar Encargos sin Ausentismos") {
                    Thread.Sleep(1000);
                    desktopSession.Mouse.MouseMove(elem.Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                }
            }
        }

        private static void procesoEncargos(WindowsDriver<WindowsElement> desktopSession, List<string> lupasEncargos)
        {
           
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> boxes = desktopSession.FindElementsByClassName("TDBCheckBox");
            foreach (var elem in boxes) {
                if (elem.Text == "Validar Tipo Cargo ") {
                    desktopSession.Mouse.MouseMove(elem.Coordinates);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                }
            }
            PruebaCRUD.LupaDinamica(desktopSession, lupasEncargos);
        }


        private static void procesoParametros(WindowsDriver<WindowsElement> desktopSession, string motor)
        {
            List<string> checkboxNamesSQL = new List<string>() {
                "Validar sueldo de contrato de empleado a reemplazar",
                "Actualizar Planta",
                "Actualizar Resolución de Contratos",
                "Validar Ocupaciones",
                "Inactivar Cadena en Retiro",
                "No Tomar Ubicación Empleado a Reemplazar",
                "Actualizar Cargo en Encargo por Funciones",
                "Validar Empleado en Encargo",
                "Actualizar Cargo Relacionado"
            };

            List<string> checkboxNamesORA = new List<string>() {
                "Actualizar Planta",
                "Actualizar Resolución de Contratos",
                "Validar Ocupaciones",
                "No Tomar Ubicación Empleado a Reemplazar",
                "Actualizar Cargo en Encargo por Funciones",
                "Validar Empleado en Encargo",
                "Actualizar Cargo Relacionado"
            };

            List<string> checkboxNames = motor == "SQL" ? checkboxNamesSQL : checkboxNamesORA;

            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> boxes = desktopSession.FindElementsByClassName("TDBCheckBox");
            foreach (var elem in boxes)
            {
                if (checkboxNames.Contains(elem.Text))
                {
                    desktopSession.Mouse.MouseMove(elem.Coordinates);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                }
            }
        }


    }
}
