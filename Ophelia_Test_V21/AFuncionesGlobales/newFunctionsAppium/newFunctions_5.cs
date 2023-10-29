using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Interactions;
using System.Threading;
using Ophelia_Test_V21.AFuncionesGlobales;
using System.Diagnostics;
using System.Drawing;
using OpenQA.Selenium;
using System.IO;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;

namespace Ophelia_Test_V21
{
    class newFunctions_5 : FuncionesVitales
    {
        public static WindowsDriver<WindowsElement> RootSession()
        {
            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }

        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string Clase)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(Clase);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                //appCapabilities.AddAdditionalCapability("app", "Root");
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                ////Debugger.Launch();
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }

        public static List<string> ReporteDinamico(WindowsDriver<WindowsElement> desktopSession, string file, string ruta, string nomPrograma = null)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();

            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            try
            {
                wait.Until(driver =>
                {
                    WindowsDriver<WindowsElement> rootSession = null;

                    List<Point> coordenadas = CoordinatesFinder(desktopSession, 241, 102, 57);
                    int CoordX = coordenadas[0].X;
                    int CoordY = coordenadas[0].Y;
                    
                    var PanelElements = desktopSession.FindElement(By.Name(desktopSession.Title));
                    bool IsDisplayedReport = false;               

                    // Clic en el icono de reporte dinamico
                    desktopSession.Mouse.MouseMove(PanelElements.Coordinates, CoordX, CoordY);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(3000);

                    for(int i = 0; i < 20; i++)
                    {
                        try
                        {
                            Thread.Sleep(3000);
                            rootSession = RootSession();                          
                            rootSession = ReloadSession(rootSession, "TFrmUDRepor");
                            var elements = rootSession.FindElementsByClassName("TFrmUDRepor");
                            if (elements.Count > 0)
                            {
                                IsDisplayedReport = true;
                            }
                            else
                            {
                                errorMessages.Add($"No se puede visualizar la ventana Generar Reporte Dinámico");
                            }
                            break;
                        }
                        catch
                        {
                            newFunctions_4.ScreenshotNuevo("Error en el reporte dinamico", true, file);
                            Thread.Sleep(500);
                            Console.WriteLine("Error en el reporte dinamico.");
                            desktopSession.Close();
                            desktopSession.Dispose();
                            Thread.Sleep(1000);
                            break;
                        }
                    }
                    
                    if (IsDisplayedReport == true)
                    {
                        try
                        {

                            var ReportElements = rootSession.FindElementsByClassName("TKsListCodigo");

                            if (ReportElements.Count > 0)
                            {
                                // Agregar Variables para el reporte
                                rootSession.Mouse.MouseMove(ReportElements[1].Coordinates, 76, 7);
                                rootSession.Mouse.DoubleClick(null);
                                rootSession.Mouse.DoubleClick(null);                                

                                rootSession = ReloadSession(rootSession, "TFrmUDRepor");
                                // Generar reporte
                                var ElementGenerar = rootSession.FindElementByName("Generar");
                                ElementGenerar.Click();
                                Thread.Sleep(3000);
                                newFunctions_4.ScreenshotNuevo("Datos Seleccionados para Generar el Reporte Dinámico", true, file);

                                string salir = null;
                                while(salir == null)
                                {
                                    // Generar Preview
                                    var ElementPreview = rootSession.FindElementByName("Preview");
                                    ElementPreview.Click();
                                    Thread.Sleep(1000);
                                    try
                                    {
                                        rootSession = RootSession();
                                        rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                                        var validar = rootSession.FindElementsByClassName("TfrxPreviewForm");
                                        if(validar.Count > 0)
                                        {
                                            salir = "saliendo";
                                        }
                                    }
                                    catch
                                    {
                                        rootSession = RootSession();
                                        rootSession = ReloadSession(rootSession, "TFrmUDRepor");
                                        Thread.Sleep(1000);
                                    }
                                }
                                newFunctions_4.ScreenshotNuevo("Preview del Reporte Dinámico Generado", true, file);

                                // Generar Reportes
                                errors = GenerarReportes("Dinámico", file, ruta, nomPrograma);
                                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }

                                // Salir ventana "Generar Reporte"
                                rootSession = RootSession();
                                rootSession = ReloadSession(rootSession, "TFrmUDRepor");
                                var ElementSalir = rootSession.FindElementByName("Salir");
                                ElementSalir.Click();
                            }
                        }
                        catch
                        {
                            errorMessages.Add($"Fallo en el Proceso");
                        }
                    }
                    return rootSession != null;
                });                
            }
            catch
            {
                
            }
            return errorMessages;            
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

            if (NomReporte.Contains("Excel"))
            {
                if (File.Exists(Ruta + Extension))
                {
                    if(nomPrograma == null)
                    {
                        //Abrir reporte
                        Process.Start(Ruta + Extension);
                        Thread.Sleep(4000);
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "XLMAIN");
                        newFunctions_4.ScreenshotNuevo("Reporte " + TipoReporte + " Exportado a Excel", true, file);
                        LimpiarProcesos();
                    }                    
                }
                else
                {
                    if (nomPrograma == null)
                    {
                        Extension = ".xls";
                        //Abrir reporte
                        Process.Start(Ruta + Extension);
                        Thread.Sleep(4000);
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "XLMAIN");
                        newFunctions_4.ScreenshotNuevo("Reporte " + TipoReporte + " Exportado a Excel", true, file);
                        LimpiarProcesos();
                    }                    
                }
            }
            else
            {
                //Abrir reporte
                Process.Start(Ruta + Extension);
                Thread.Sleep(4000);
                if (Extension == ".docx")
                {
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "OpusApp");
                    newFunctions_4.ScreenshotNuevo("Reporte " + TipoReporte + " Exportado a Word", true, file);
                }
                else if (Extension == ".pdf")
                {
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "AcrobatSDIWindow");
                    newFunctions_4.ScreenshotNuevo("Reporte " + TipoReporte + " Exportado a Pdf", true, file);
                }                
                LimpiarProcesos();
            }            
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

            //Debugger.Launch();
            // Revisar que se abrio el preview
            bool IsDisplayedReview = false;
            
            try
            {
                int ext = 0;
                for(int i = 0; i < 2; i++)
                {
                    try
                    {
                        Thread.Sleep(5000);
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                        var validar = rootSession.FindElementsByClassName("TfrxPreviewForm");
                        break;
                    }
                    catch
                    {
                        newFunctions_4.ScreenshotNuevo("Error en el reporte preliminar", true, file);
                        Thread.Sleep(500);
                        errorMessages.Add("Error en el reporte preliminar.");
                        //rootSession = RootSession();
                        //rootSession = ReloadSession(rootSession, "Chrome_WidgetWin_1");
                        //rootSession.Close();
                        //rootSession.Dispose();
                        //ext = 1;
                        //break;
                        string varNavigator = null;
                        varNavigator = "msedge";
                        Process[] processes = Process.GetProcessesByName(varNavigator);

                        if (processes.Length > 0)
                        {
                            for (int j = 0; j < processes.Length; j++)
                            {
                                processes[j].Kill();
                            }
                        }
                        return errorMessages;
                    }
                }

                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxPreviewForm");

                int count = 0;
                while(count < 300)
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
                if (codProgram != "KNmNoema")
                {
                    //Clic en Disquet
                    rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 57, 17);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(2000);

                    //Reportes
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "#32768");
                    ////Debugger.Launch();
                    var MenuItem = rootSession.FindElementsByXPath("/Menu[@Name=\"Contexto\"][@ClassName=\"#32768\"]/*");
                    if (MenuItem.Count == 0)
                    {
                        //MenuItem = rootSession.FindElementsByXPath("/Menu[@ClassName='#32768'][@Name='Context']/*");
                        MenuItem = rootSession.FindElementsByXPath("//Menu[contains(@ClassName,#32768)][contains(@Name,Context)]");
                        //Desktop 1
                    }
                    for (int i = 0; i < MenuItem.Count; i++)
                    {
                        if (MenuItem[i].GetAttribute("Name").Contains("PDF"))
                        {
                            Pdf = true;
                            ReportePdf = MenuItem[i].GetAttribute("Name");
                            Ruta1 = @"C:\Reportes\Reporte_PDF_" + TipoReporte + "_" + fecha;
                            Extension1 = ".pdf";
                        }

                        if (MenuItem[i].GetAttribute("Name").Contains("Word"))
                        {
                            Word = true;
                            ReporteW = MenuItem[i].GetAttribute("Name");
                            Ruta2 = @"C:\Reportes\Reporte_Word_" + TipoReporte + "_" + fecha;
                            Extension2 = ".docx";
                        }

                        if (MenuItem[i].GetAttribute("Name").Contains("Excel") && MenuItem[i].GetAttribute("Name").Contains("OLE"))
                        {

                        }
                        else if (MenuItem[i].GetAttribute("Name").Contains("Excel"))
                        {
                            Excel = true;
                            ReporteE = MenuItem[i].GetAttribute("Name");
                            Ruta3 = @"C:\Reportes\Reporte_Excel_" + TipoReporte + "_" + fecha;
                            Extension3 = ".xlsx";
                        }
                    }

                    if (Pdf == true)
                    {
                        try
                        {
                            ExportarReporte(ReportePdf, Ruta1, Extension1, TipoReporte, file);
                        }
                        catch
                        {
                            errorMessages.Add("El Proceso de exportar el reporte " + TipoReporte + " a PDF Falló");
                        }
                    }

                    if (Word == true)
                    {
                        try
                        {
                            ExportarReporte(ReporteW, Ruta2, Extension2, TipoReporte, file);
                        }
                        catch
                        {
                            errorMessages.Add("El Proceso de exportar el reporte " + TipoReporte + " a Word Falló");
                        }
                    }

                    if (Excel == true)
                    {
                        try
                        {
                            ExportarReporte(ReporteE, Ruta3, Extension3, TipoReporte, file, nomPrograma);
                        }
                        catch
                        {
                            errorMessages.Add("El Proceso de exportar el reporte " + TipoReporte + " a Excel Falló");
                        }
                    }
                }
                

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