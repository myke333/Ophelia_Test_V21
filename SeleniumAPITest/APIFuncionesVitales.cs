using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UITesting;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Serialization;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Diagnostics;
using OpenPop.Mime;
using OpenPop.Pop3;

namespace APITest
{
    public class APIFuncionesVitales
    {
        public APIFuncionesVitales()
        {
        }

        public void Screenshot(string maestro, bool bandera, string file)
        {
            Image MyImage = null;
            string UrlImage = @"C:\Reportes\" + maestro + ".bmp";
            MyImage = UITestControl.Desktop.CaptureImage();
            MyImage.Save(UrlImage, System.Drawing.Imaging.ImageFormat.Bmp);
            InsertAPicture(file, UrlImage, maestro, bandera);
        }

        public string CrearDocumentoWordDinamico(string app, string database, string modulo)
        {
            string NumCompilacion = ReadNumCompiler(app);
            string path = string.Format(@"C:\Reportes\Reportes{0}\Reportes_{1}\Ejecucion{2}", app, database, NumCompilacion);
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string file = string.Format(@"{0}\Reportes_{1}_{2}.docx", path, modulo, Hora());
            using (WordprocessingDocument wordDocument =
            WordprocessingDocument.Create(file, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Prueba de ejecución realizada en el módulo: " + modulo));
            }
            return file;
        }

        public string Hora()
        {
            DateTime dateAndTime = DateTime.Now;
            string fecha = dateAndTime.ToString("ddMMyyyy_HHmmss");
            return fecha;
        }

        public void ConvertWordToPDF(string RutaArchivo, string database)
        {
            string[] split = RutaArchivo.Split('\\');
            string Ruta = "";
            for (int i = 0; i < (split.Length - 1); i++)
            {
                Ruta = Ruta + split[i] + @"\";
            }
            Ruta = string.Format(Ruta + @"ArchivosPDF_{0}\", database);
            string Archivo = split[split.Length - 1];
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object oMissing = System.Reflection.Missing.Value;
            word.Visible = false;
            word.ScreenUpdating = false;
            Object FileName = (Object)RutaArchivo;
            Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(ref FileName, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                            ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            doc.Activate();
            if (!Directory.Exists(Ruta))
            {
                Directory.CreateDirectory(Ruta);
            }
            object Filter = Archivo.Replace(".docx", ".pdf");
            object outputFileName = Ruta + Filter;
            object fileFormat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF;
            doc.SaveAs(ref outputFileName,
                    ref fileFormat, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
            ((Microsoft.Office.Interop.Word.Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
            doc = null;
            ((Microsoft.Office.Interop.Word.Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
            word = null;
            Process[] processes2 = Process.GetProcessesByName("WINWORD");
            if (processes2.Length > 0)
            {
                for (int i = 0; i < processes2.Length; i++)
                {
                    processes2[i].Kill();
                }
            }
        }

        public void LimpiarProcesos()
        {
            Process[] processes = Process.GetProcessesByName("AcroRd32");
            if (processes.Length > 0)
            {
                for (int i = 0; i < processes.Length; i++)
                {
                    processes[i].Kill();
                }
            }
            Process[] processes1 = Process.GetProcessesByName("EXCEL");
            if (processes1.Length > 0)
            {
                for (int i = 0; i < processes1.Length; i++)
                {
                    processes1[i].Kill();
                }
            }
            Process[] processes2 = Process.GetProcessesByName("WINWORD");
            if (processes2.Length > 0)
            {
                for (int i = 0; i < processes2.Length; i++)
                {
                    processes2[i].Kill();
                }
            }

            Process[] processes3 = Process.GetProcessesByName("notepad");
            if (processes3.Length > 0)
            {
                for (int i = 0; i < processes3.Length; i++)
                {
                    processes3[i].Kill();
                }
            }
        }

        public void ReporteErrores(List<string> Error, string testname)
        {
            string ruta = string.Format(@"C:\Reportes\Errores\{0}.txt", testname);
            StreamWriter sw = new StreamWriter(ruta);
            for (int i = 0; i < (Error.ToArray()).Length; i++)
            {
                sw.WriteLine(Error[i]);
            }
            sw.Close();
        }

        public List<string> ValidarCorreo(string Correo, string Password, string DescripcionRemitente, string ContenidoCorreo)
        {
            List<string> Error = new List<string>();
            string UserEmail = $"recent:{Correo}";
            int port = 995;
            bool UseSSL = true;
            string Hostname = "pop.gmail.com";
            string DescRemitente = string.Empty;
            string Body = string.Empty;
            int count = 0;
            StringBuilder builder = new StringBuilder();

            Pop3Client popClient = new Pop3Client();
            try
            {    
                popClient.Connect(Hostname, port, UseSSL);
                popClient.Authenticate(UserEmail, Password);
                count = popClient.GetMessageCount();

                Message message = popClient.GetMessage(count);

                DescRemitente = message.Headers.From.DisplayName;
                if (DescRemitente != DescripcionRemitente)
                {
                    Error.Add("::::::NO SE ENCONTRO EL CORREO EN LA BANDEJA DE ENTRADA::::::");
                    Error.Add($"La descripcion del remitente del correo no coincide, se esperaba: {DescripcionRemitente} y se encontró {DescRemitente}");
                }

                MessagePart plainText = message.FindFirstPlainTextVersion();
                builder.Append(plainText.GetBodyAsText());
                Body = builder.ToString();
                if (!Body.Contains(ContenidoCorreo))
                {
                    if (Error.Count <= 0)
                    {
                        Error.Add("::::::NO SE ENCONTRO EL CORREO EN LA BANDEJA DE ENTRADA::::::");
                    }
                    Error.Add($"No se encontro el contenido esperado en el correo '{ContenidoCorreo}'");
                }
                Thread.Sleep(500);
                popClient.DeleteMessage(count);
                popClient.Disconnect();

                return Error;
            }
            catch (Exception ex)
            {
                Thread.Sleep(500);
                popClient.Disconnect();
                Error.Add("LA BANDEJA DE ENTRADA ESTA VACIA::::::" + ex.ToString());
                return Error;
            }
        }

        public static void InsertAPicture(string document, string UrlImage, string maestro, bool bandera)
        {
            using (WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(document, true))
            {
                MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

                using (FileStream stream = new FileStream(UrlImage, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart), maestro, bandera, UrlImage);
            }
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId, string maestro, bool bandera, string UrlImage)
        {
            int iWidth = 0;
            int iHeight = 0;
            using (Bitmap bmp = new Bitmap(UrlImage))
            {
                iWidth = bmp.Width;
                iHeight = bmp.Height;
            }
            iWidth = (int)Math.Round((decimal)iWidth * 4000);
            iHeight = (int)Math.Round((decimal)iHeight * 4000);

            var element =

                new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = iWidth, Cy = iHeight },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.bmp"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.None
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 10000L, Y = 10000L },
                                             new A.Extents() { Cx = iWidth, Cy = iHeight }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            if (bandera)
            {
                wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(new Text(maestro))));
            }
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));

        }

        public static string ReadNumCompiler(string app)
        {
            string PathXML = "";
            if (app.ToLower() == "selfservice" || app.ToLower() == "smartpeople")
            {
                PathXML = @"\\dwtfskscm\TfsScripts\TfsCompilerNumber\VariablesSelfService.xml";
            }
            else if (app.ToLower() == "reclutamiento")
            {
                PathXML = @"\\dwtfskscm\TfsScripts\TfsCompilerNumber\VariablesReclutamiento.xml";
            }
            XmlDocument xmlDoc = new XmlDocument();

            xmlDoc.Load(PathXML);
            string NumCambios = "";
            foreach (XmlNode nodeProject in xmlDoc.GetElementsByTagName("VARIABLES"))
            {
                foreach (XmlNode nodePropertyGroup in nodeProject.ChildNodes)
                {
                    foreach (XmlNode nodeVerInfo_Keys in nodePropertyGroup.ChildNodes)
                    {
                        if (nodeVerInfo_Keys.Name == "GRUPOCAMBIOS")
                        {
                            NumCambios = nodeVerInfo_Keys.InnerText;
                        }
                    }

                }
            }
            return NumCambios;
        }

    }
}
