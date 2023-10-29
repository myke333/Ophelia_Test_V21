using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
    class CrudKnmpcesa
    {
        public static void CrudKNmpcesa(WindowsDriver<WindowsElement> desktopSession, int Tipo, string motor)
        {
            if (Tipo == 1)
            {
                if (motor == "ORA")
                {
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Dias Parciales en Intereses").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Inserta Liquidaciones Parciales en Prenomina").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Inserta Liquidaciones Definitivas en Prenomina").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Consolidar empleados activos e inactivos del mes").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Usar Rango del Concepto para Promedio").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Redondea Base").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Base Parcial").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Base Fin de Año").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Base Definitiva").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Base Consolidada").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                }
                else
                {
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Exige Solicitud para Liquidación Parcial").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                }
            }
            else if (Tipo == 2)
            {
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Dias Parciales en Intereses").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Parámetros Generales").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);                
            }
        }
    }
}
