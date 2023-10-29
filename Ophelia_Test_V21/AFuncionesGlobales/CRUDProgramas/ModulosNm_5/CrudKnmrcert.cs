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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_5
{
    class CrudKnmrcert : FuncionesVitales
    {
        public static void CRUDKNmRcert(WindowsDriver<WindowsElement> desktopSession, string file)
        {            
            var Element = desktopSession.FindElementsByClassName("TEdit");
            WindowsDriver<WindowsElement> rootSession = null;
            string Ruta = @"C:\Kactus\Garh\Nm\wrk";
            Thread.Sleep(500);
            desktopSession.Mouse.MouseMove(Element[1].Coordinates, 350, 10);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "#32770");
            Thread.Sleep(500);
            var Element4 = rootSession.FindElementsByClassName("Edit");
            var Element5 = rootSession.FindElementsByClassName("Button");
            rootSession = PruebaCRUD.RootSession();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            //Element4[0].SendKeys(Ruta);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Selccionar carpeta", true, file);
            //Thread.Sleep(500);
            //Element5[0].Click();
            //Thread.Sleep(500);

        }
        public static void ImpimirKNmRcert(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            var Element = desktopSession.FindElementsByName("Imprimir");
            //BOTON IMPRIMIR
            desktopSession.Mouse.MouseMove(Element[0].Coordinates, 60, 20);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Reporte Imprimir", true, file);
            string pdf = @"C:\Reportes\ReportePDF1_" + "Imprimir" + "_" + Hora();
            newFunctions_5.GenerarReportes("Imprimir", file, pdf);
            Thread.Sleep(500);
        }
        public static List<string> PreliminarKNmRcert(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            List<string> errorMessages = new List<string>();
            string pdf = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
            WindowsDriver<WindowsElement> rootSession = null;
            try
            {
                rootSession = PruebaCRUD.RootSession();
            rootSession = PruebaCRUD.ReloadSession(rootSession, "IEFrame");
            var allFrame = rootSession.FindElementsByClassName("IEFrame");
            Thread.Sleep(500);
            new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset((allFrame[0].Size.Width / 2), (allFrame[0].Size.Height / 2) + 40).DoubleClick().Perform();
            Thread.Sleep(500);
            try
            {
                newFunctions_5.GenerarReportes("Preliminar", file, pdf);
            }
            catch (Exception)
            {
                PruebaCRUD.VentanaAzul(desktopSession);
            }
            
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }
            return errorMessages;
        }
    }
}
