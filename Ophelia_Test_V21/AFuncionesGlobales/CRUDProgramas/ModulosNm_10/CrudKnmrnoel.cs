using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Drawing;
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
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_10
{
    class CrudKnmrnoel : FuncionesVitales
    {
        static public WindowsDriver<WindowsElement> RootSession()
        {
            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }

        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string parametro)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(parametro);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex

                var appCapabilities = new AppiumOptions();
                //appCapabilities.AddAdditionalCapability("app", "Root");
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }

        static public void ClickButtonAceptarEnter(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {

            var btnTDVI = desktopSession.FindElementsByClassName("TBitBtn");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //var btnTDVI = desktopSession.FindElementsByClassName("TPanel");
            //Debugger.Launch();


            if (bandera == 0)
            {
                var ElementList = desktopSession.FindElementsByName("Aceptar");
                ElementList[0].Click();
                Thread.Sleep(3000);


            }
            else
            {
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1500);
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1500);

                //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 101, 27);
                //desktopSession.Mouse.Click(null);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[3]);
                //Thread.Sleep(1500);

            }
        }


        static public void Click(WindowsDriver<WindowsElement> desktopSession, int bandera)
        {

            //var btnTDVI2 = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //Debugger.Launch();

            if (bandera == 0)
            {

                //desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 472, 40);
                //desktopSession.Mouse.DoubleClick(null);
                //Thread.Sleep(1000);
                ////desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                ////Thread.Sleep(1500);
                //desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                //Thread.Sleep(1000);
                //desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                //Thread.Sleep(1500);
                //newFunctions_4.ScreenshotNuevo("Datos de configuracion de conexiones", true, file);
                //Thread.Sleep(1000);
                ////////////////////////////////
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession = ReloadSession(RootSession, "XLMAIN");

                //var btnTDVI = RootSession.FindElementsByClassName("TTabSheet");

                var ElementList = RootSession.FindElementsByName("Cerrar");
                ElementList[0].Click();
                Thread.Sleep(2000);
                //WindowsDriver<WindowsElement> RootSession2 = null;
                //RootSession2 = PruebaCRUD.RootSession();
                //RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(1000);

            }
            if (bandera == 1)
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 900, 383);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession.Keyboard.SendKeys("ReportePDF1_Preliminar_08022022_161632");
                //Thread.Sleep(1000);
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            }
            if (bandera == 2)
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 430, -30);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession.Keyboard.SendKeys("ReportePDF1_Preliminar_08022022_161632");
                //Thread.Sleep(1000);
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            }

            if (bandera == 3)
            {
                WindowsDriver<WindowsElement> RootSession2 = null;
                RootSession2 = PruebaCRUD.RootSession();
                RootSession2.Keyboard.SendKeys("Prueba");
                Thread.Sleep(1500);
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                RootSession2.Keyboard.SendKeys("Esto es una prueba");
                Thread.Sleep(1500);
                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession.Keyboard.SendKeys("ReportePDF1_Preliminar_08022022_161632");
                //Thread.Sleep(1000);
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            }

            if (bandera == 4)
            {
                WindowsDriver<WindowsElement> RootSession2 = null;
                RootSession2 = PruebaCRUD.RootSession();
                RootSession2 = ReloadSession(RootSession2, "TFrmMensaje");

                //var btnTDVI = RootSession.FindElementsByClassName("TTabSheet");

                var ElementList = RootSession2.FindElementsByName("Aceptar");
                ElementList[0].Click();
                Thread.Sleep(5000);
                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession.Keyboard.SendKeys("ReportePDF1_Preliminar_08022022_161632");
                //Thread.Sleep(1000);
                //RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

            }
            if (bandera == 5)
            {

                
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession = ReloadSession(RootSession, "Notepad");

                //var btnTDVI = RootSession.FindElementsByClassName("TTabSheet");

                var ElementList = RootSession.FindElementsByName("Close");
                ElementList[0].Click();
                Thread.Sleep(2000);
                //WindowsDriver<WindowsElement> RootSession2 = null;
                //RootSession2 = PruebaCRUD.RootSession();
                //RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(1000);

            }

        }

        static public void ClickTotalGeneral(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal, string file)
        {

            //var btnTDVI2 = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //Debugger.Launch();

            if (bandera == 0)
            {

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 472, 40);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[0]);
                Thread.Sleep(1500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(crudPrincipal[1]);
                Thread.Sleep(1500);
                newFunctions_4.ScreenshotNuevo("Datos de fechas", true, file);
                Thread.Sleep(1000);
                ////////////////////////////////
                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession = ReloadSession(RootSession, "TfrmConfig");

                //var btnTDVI = RootSession.FindElementsByClassName("TTabSheet");

                //RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 318, -54);
                //RootSession.Mouse.Click(null);
                //Thread.Sleep(2000);
                //WindowsDriver<WindowsElement> RootSession2 = null;
                //RootSession2 = PruebaCRUD.RootSession();
                //RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 293, 73);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 701, 83);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1500);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
            }
        }

        static public void ClickDetallePorEmpleado(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal, string file)
        {

            //var btnTDVI2 = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //Debugger.Launch();

            if (bandera == 0)
            {

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 27, 51);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Detalle por Empleado", true, file);
                Thread.Sleep(1000);
                ////////////////////////////////
                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession = ReloadSession(RootSession, "TfrmConfig");

                //var btnTDVI = RootSession.FindElementsByClassName("TTabSheet");

                //RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 318, -54);
                //RootSession.Mouse.Click(null);
                //Thread.Sleep(2000);
                //WindowsDriver<WindowsElement> RootSession2 = null;
                //RootSession2 = PruebaCRUD.RootSession();
                //RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 293, 73);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 701, 83);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1500);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
            }
        }

        static public void ClickDiferencia(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal, string file)
        {

            //var btnTDVI2 = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //Debugger.Launch();

            if (bandera == 0)
            {

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 27, 64);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Diferencia Acumulados VS Nomina Electronica", true, file);
                Thread.Sleep(1000);
                ////////////////////////////////
                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession = ReloadSession(RootSession, "TfrmConfig");

                //var btnTDVI = RootSession.FindElementsByClassName("TTabSheet");

                //RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 318, -54);
                //RootSession.Mouse.Click(null);
                //Thread.Sleep(2000);
                //WindowsDriver<WindowsElement> RootSession2 = null;
                //RootSession2 = PruebaCRUD.RootSession();
                //RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 293, 73);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 701, 83);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1500);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
            }
        }

        static public void ClickTotalDetalleEmpleado(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal, string file)
        {

            //var btnTDVI2 = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //Debugger.Launch();

            if (bandera == 0)
            {

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 27, 81);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Total de detalle por empleado", true, file);
                Thread.Sleep(1000);
                ////////////////////////////////
                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession = ReloadSession(RootSession, "TfrmConfig");

                //var btnTDVI = RootSession.FindElementsByClassName("TTabSheet");

                //RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 318, -54);
                //RootSession.Mouse.Click(null);
                //Thread.Sleep(2000);
                //WindowsDriver<WindowsElement> RootSession2 = null;
                //RootSession2 = PruebaCRUD.RootSession();
                //RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 293, 73);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 701, 83);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                WindowsDriver<WindowsElement> RootSession = null;
                RootSession = PruebaCRUD.RootSession();
                RootSession.Keyboard.SendKeys(crudPrincipal[2]);
                Thread.Sleep(1500);
                RootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
            }
        }

        static public void ClickDesprendible(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudPrincipal, string file)
        {

            //var btnTDVI2 = desktopSession.FindElementsByClassName("TDBEdit");
            var btnTDVI2 = desktopSession.FindElementsByClassName("TScrollBox");
            //Debugger.Launch();

            if (bandera == 0)
            {

                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 27, 116);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
                newFunctions_4.ScreenshotNuevo("Desprendible de pago nomina electronica", true, file);
                Thread.Sleep(1000);
                ////////////////////////////////
                //WindowsDriver<WindowsElement> RootSession = null;
                //RootSession = PruebaCRUD.RootSession();
                //RootSession = ReloadSession(RootSession, "TfrmConfig");

                //var btnTDVI = RootSession.FindElementsByClassName("TTabSheet");

                //RootSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 318, -54);
                //RootSession.Mouse.Click(null);
                //Thread.Sleep(2000);
                //WindowsDriver<WindowsElement> RootSession2 = null;
                //RootSession2 = PruebaCRUD.RootSession();
                //RootSession2.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                //Thread.Sleep(1000);

            }
            else
            {
                desktopSession.Mouse.MouseMove(btnTDVI2[0].Coordinates, 293, 73);
                desktopSession.Mouse.DoubleClick(null);
                Thread.Sleep(1000);
            }
        }

        public static void ClickButtonErrores(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByName("Errores");
            newFunctions_4.ScreenshotNuevo("Errores", true, file);
            ElementList[0].Click();
        }

        private static List<Point> CoordinatesFinder(WindowsDriver<WindowsElement> session, int R, int G, int B)
        {
            Screenshot image = ((ITakesScreenshot)session).GetScreenshot();
            string name = FuncionesVitales.Hora();
            string path = "C:\\imagenesTest\\" + name + ".Png";
            Image imgSource;
            try
            {
                image.SaveAsFile(string.Format("C:\\imagenesTest\\{0}.Png", name), ScreenshotImageFormat.Png);
                Thread.Sleep(1000);
                imgSource = Image.FromFile(path);
            }
            catch (Exception e)
            {
                image.SaveAsFile(string.Format("C:\\imagenesTest\\{0}.Png", name), ScreenshotImageFormat.Png);
                Thread.Sleep(1000);
                imgSource = Image.FromFile(path);
            }

            Point puntos = new Point();
            int lastx = 0;
            int lasty = 0;
            List<Point> coordenadas = new List<Point>();
            Bitmap bmpSource = new Bitmap(imgSource);
            Bitmap bmpDest = new Bitmap(bmpSource.Width, bmpSource.Height);
            int Romper = 0;

            for (int x = 0; x < bmpSource.Width; x++)
            {
                for (int y = 0; y < bmpSource.Height; y++)
                {
                    Color clrPixel = bmpSource.GetPixel(x, y);
                    ////Debugger.Launch();
                    if (clrPixel.R == R && clrPixel.G == G && clrPixel.B == B)
                    {
                        puntos.X = x;
                        puntos.Y = y;
                        if (x > lastx + 20 || y > lasty + 20)
                        {
                            coordenadas.Add(puntos);
                            lastx = x;
                            lasty = y;
                            Romper++;
                            break;
                        }
                    }
                }
                if (Romper != 0)
                {
                    break;
                }
            }
            ////Debugger.Launch();
            return coordenadas;
        }

        private static void ExportarReporte(string NomReporte, string Ruta, string Extension, string TipoReporte, string file, string nomPrograma = null)
        {
            WindowsDriver<WindowsElement> rootSession;

            if (NomReporte.Contains("Word") || NomReporte.Contains("Excel"))
            {
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                var ElementsPreview = rootSession.FindElementsByClassName("TToolBar");

                Thread.Sleep(1000);
                //Clic en Disquet
                rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 57, 17);
                rootSession.Mouse.Click(null);
                Thread.Sleep(1000);
            }

            rootSession = RootSession();
            //clic en Tipo Reporte
            var ClicReport = rootSession.FindElementByName(NomReporte);
            ClicReport.Click();
            Thread.Sleep(1000);
            rootSession = RootSession();
            var AcceptPDF = rootSession.FindElementByName("Aceptar");

            //clic en aceptar exportar Reporte
            AcceptPDF.Click();
            Thread.Sleep(2000);

            //ingresar la ruta
            rootSession.Keyboard.SendKeys(Ruta);
            Thread.Sleep(1000);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(3000);            
               
            if (Extension == ".pdf")
                {
                  rootSession = RootSession();
                  rootSession = ReloadSession(rootSession, "AcrobatSDIWindow");
                  newFunctions_4.ScreenshotNuevo("Reporte " + TipoReporte + " Exportado a Pdf", true, file);
                }
            LimpiarProcesos();
            
        }

        public static List<string> GenerarReportes(string TipoReporte, string file, string ruta, string nomPrograma = null, string codProgram = null)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession;
            string ReportePdf = "";
            string ReporteW = "";
            string ReporteE = "";
            string Ruta1 = "";
            string Extension1 = "";
            string Ruta2 = "";
            string Extension2 = "";
            string Ruta3 = "";
            string Extension3 = "";
            bool Pdf = false;
            bool Word = false;
            bool Excel = false;

            // Revisar que se abrio el preview
            bool IsDisplayedReview = false;

            try
            {
                for (int i = 0; i < 20; i++)
                {
                    try
                    {
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                        var validar = rootSession.FindElementsByClassName("TfrxPreviewForm");
                        break;
                    }
                    catch
                    {
                        Thread.Sleep(1000);
                    }
                }
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxPreviewForm");

                int count = 0;
                while (count < 300)
                {
                    try
                    {
                        List<Point> carga = CoordinatesFinder(rootSession, 224, 32, 64);
                        break;
                    }
                    catch
                    {
                        Thread.Sleep(1000);
                        count++;
                    }
                }

                List<Point> coordenadas = CoordinatesFinder(rootSession, 247, 223, 99);
                int CoordX = coordenadas[0].X;
                int CoordY = coordenadas[0].Y;

                var PanelElements = rootSession.FindElement(By.Name(rootSession.Title));
                bool IsDisplayedReport = false;

                // Clic en el icono de editar reporte
                rootSession.Mouse.MouseMove(PanelElements.Coordinates, CoordX, CoordY);
                rootSession.Mouse.Click(null);
                Thread.Sleep(8000);
                newFunctions_4.ScreenshotNuevo("La opción 'Editar Reporte' " + TipoReporte + " está activada", true, file);

                errorMessages.Add("El Preview del reporte " + TipoReporte + " tiene habilitado la opcion 'Editar Reporte'");

                // Cerrar la ventana de editar Reporte
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxDesignerForm");

                var CerrarEditar = rootSession.FindElementByName("Cerrar");
                CerrarEditar.Click();
                Thread.Sleep(2000);
                newFunctions_4.ScreenshotNuevo("Ventana 'Editar Reporte' " + TipoReporte + " Cerrada", true, file);
            }
            catch
            {

            }

            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TfrxPreviewForm");

            var ElementsPreview = rootSession.FindElementsByClassName("TToolBar");

            if (ElementsPreview.Count > 0)
            {
                IsDisplayedReview = true;
            }
            else
            {
                errorMessages.Add("No se puede visualizar la ventana Preview del Reporte " + TipoReporte);
            }

            string fecha = Hora();
            if (IsDisplayedReview == true)
            {
                ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                int cerrar = 0;
                try
                {
                    Thread.Sleep(2000);
                    List<Point> coordenadas1 = CoordinatesFinder(rootSession, 224, 32, 64);
                    int CoordX = coordenadas1[0].X;
                    int CoordY = coordenadas1[0].Y;

                    var PanelElements = rootSession.FindElement(By.Name(rootSession.Title));

                    // Clic en el icono de exportar PDF
                    rootSession.Mouse.MouseMove(PanelElements.Coordinates, CoordX, CoordY);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TfrxPDFExportDialog");

                    var ExportPDF = rootSession.FindElementByName("Aceptar");

                    if (ExportPDF != null)
                    {
                        //Clic en Aceptar exportar PDF
                        ExportPDF.Click();
                        Thread.Sleep(2000);
                        rootSession = RootSession();

                        try
                        {
                            //ingresar la ruta
                            rootSession.Keyboard.SendKeys(ruta);
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            Thread.Sleep(2000);
                            //Abrir PDF
                            Process.Start(ruta + ".pdf");
                            Thread.Sleep(6000);
                            rootSession = RootSession();
                            rootSession = ReloadSession(rootSession, "AcrobatSDIWindow");
                            //var TraerPantalla = rootSession.FindElementsByName("AVL_AVView");
                            newFunctions_4.ScreenshotNuevo("Reporte " + TipoReporte + " exportado a PDF 1", true, file);
                            LimpiarProcesos();
                        }
                        catch
                        {
                            var Ok = rootSession.FindElementByName("OK");
                            Ok.Click();
                            Thread.Sleep(1000);
                            errorMessages.Add("No existe el botón exportar reporte a PDF");
                        }
                    }
                }
                catch
                {
                    errorMessages.Add("En el Preview  del Reporte " + TipoReporte + " NO Existe el Botón Exportar Reporte a PDF");
                    cerrar = 1;
                }

                Thread.Sleep(2000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                Thread.Sleep(1000);
                //Debugger.Launch();
                //if (codProgram != "KNmNoema")
                //{
                //    //Clic en Disquet
                //    rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 57, 17);
                //    rootSession.Mouse.Click(null);
                //    Thread.Sleep(2000);

                //    //Reportes
                //    rootSession = RootSession();
                //    rootSession = ReloadSession(rootSession, "#32768");
                //    ////Debugger.Launch();
                //    var MenuItem = rootSession.FindElementsByXPath("/Menu[@Name=\"Contexto\"][@ClassName=\"#32768\"]/*");
                //    if (MenuItem.Count == 0)
                //    {
                //        //MenuItem = rootSession.FindElementsByXPath("/Menu[@ClassName='#32768'][@Name='Context']/*");
                //        MenuItem = rootSession.FindElementsByXPath("//Menu[contains(@ClassName,#32768)][contains(@Name,Context)]");
                //        //Desktop 1
                //    }
                //    for (int i = 0; i < MenuItem.Count; i++)
                //    {
                //        if (MenuItem[i].GetAttribute("Name").Contains("PDF"))
                //        {
                //            Pdf = true;
                //            ReportePdf = MenuItem[i].GetAttribute("Name");
                //            Ruta1 = @"C:\Reportes\Reporte_PDF_" + TipoReporte + "_" + fecha;
                //            Extension1 = ".pdf";
                //        }

                //    }

                //    if (Pdf == true)
                //    {
                //        try
                //        {
                //            ExportarReporte(ReportePdf, Ruta1, Extension1, TipoReporte, file);
                //        }
                //        catch
                //        {
                //            errorMessages.Add("El Proceso de exportar el reporte " + TipoReporte + " a PDF Falló");
                //        }
                //    }                                       

                //}


                Thread.Sleep(2000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                Thread.Sleep(2500);

                if (cerrar == 0)
                {
                    //Cerrar Preview
                    rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 600, 17);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(2000);
                    try
                    {
                        ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                        //Cerrar Preview
                        rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 543, 17);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(2000);
                        try
                        {
                            rootSession.Mouse.MouseMove(rootSession.FindElementByName("Cerrar").Coordinates);
                            rootSession.Mouse.Click(null);
                            errorMessages.Add("El Botón 'Cerrar' no Existe en el Preview del Reporte " + TipoReporte);
                        }
                        catch
                        {

                        }
                    }
                    catch
                    {

                    }
                }
                else
                {
                    //Cerrar Preview
                    rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 520, 17);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(2000);
                    try
                    {
                        ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                        //Cerrar Preview
                        rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 543, 17);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(2000);
                        try
                        {
                            rootSession.Mouse.MouseMove(rootSession.FindElementByName("Cerrar").Coordinates);
                            rootSession.Mouse.Click(null);
                            errorMessages.Add("El Botón 'Cerrar' no Existe en el Preview del Reporte " + TipoReporte);
                        }
                        catch
                        {

                        }
                    }
                    catch
                    {

                    }
                }
            }
            return errorMessages;
        }


    }
}