using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales;
using OpenQA.Selenium;
using System.Drawing;
using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;


namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_4
{
    class CrudKnmejpro : FuncionesVitales
    {
        public static void CRUDKNmEjpro(WindowsDriver<WindowsElement> desktopSession,List<string> crudVars, string file)
        {

            List<string> data = new List<string>() { crudVars[0] };
                PruebaCRUD.LupaDinamica2(desktopSession, data);
        }
        public static void AceptarKNmEjpro(WindowsDriver<WindowsElement> desktopSession, string file,string motor)
        {
            var Aceptar = desktopSession.FindElementsByName("Aceptar");
            WindowsDriver<WindowsElement> rootSession = null;

            Thread.Sleep(500);
            Aceptar[0].Click();
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Botón Aceptar", true, file);
            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
            Thread.Sleep(500);
            Enter(desktopSession);
            Screenshot("Mensaje", true, file, desktopSession);
            Enter(desktopSession);
            Screenshot("Proceso finalizado", true, file, desktopSession);
            var allFrame = rootSession.FindElementsByClassName("IEFrame");
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2), (allFrame[0].Size.Height / 2) + 30).DoubleClick().Perform();
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Proceso Finalizado", true, file);
            Thread.Sleep(500);

            if (motor=="ORA")
            { 
            //BOTON ERRORES
            var Errores = desktopSession.FindElementsByName("Errores");
            Thread.Sleep(500);
            Errores[0].Click();
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Reporte Errores", true, file);
            string pdf = @"C:\Reportes\ReportePDF1_" + "Errores" + "_" + Hora();
            newFunctions_5.GenerarReportes("Errores", file, pdf);
            }

        }

        public static void Enter(WindowsDriver<WindowsElement> desktopSession)
        {

            WindowsDriver<WindowsElement> RootSession2 = null;
            RootSession2 = PruebaCRUD.RootSession();
            RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(2000);

        }
    }
}
