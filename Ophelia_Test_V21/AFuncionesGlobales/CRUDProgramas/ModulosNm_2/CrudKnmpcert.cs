using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_2
{
    class CrudKnmpcert
    {
        public static void CrudKNmpcert(WindowsDriver<WindowsElement> desktopSession, int Tipo, string motor)
        {
            if (Tipo == 1)
            {
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Nombre Empleado").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);

                if (motor == "ORA")
                {
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Fecha Ingreso").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Fecha vencimiento").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500); 
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Tipo Contrato").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Código Cargo").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Nombre del Cargo").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Grado del Cargo").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Sueldo Básico").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Monto Escrito Sueldo Básico").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Municipio Expedición Identificación").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Fecha Expedición").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Valor Concepto Fijo 1").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Valor Concepto Fijo 2").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Valor Concepto Fijo 3").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Codigo Interno").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Total").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);                    
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Monto Escrito Total").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Estado Funcionario").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Nombre Nivel").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Código Centro de Costos").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Nombre Centro de Costos").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Monto Escrito Sueldo Básico + S. Variable").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Usar Decimales").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Sueldo Variable").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Monto Escrito Sueldo Variable").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Distribuir Conceptos").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Sueldo Básico + Extrasalarial").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Monto Escrito Sueldo Básico + Extrasalarial").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Fecha de Nombramiento").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Fecha de Posesión").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Resolución de Nombramiento").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Resolución de Posesión").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Monto Escrito Concepto Fijo1").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Monto Escrito Concepto Fijo2").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Monto Escrito Concepto Fijo3").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Fecha Inicio del Encargo").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Fecha Fin del Encargo").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Cargo Relacionado").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                    desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Fecha De Contrato").Coordinates);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(500);
                }
            }
            else if (Tipo == 2)
            {
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByName("Identificación").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(500);
            }
        }
    }
}
