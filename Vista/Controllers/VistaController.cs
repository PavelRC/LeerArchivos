using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Vista.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System.Collections;
using System.IO;
using System.Reflection;
using System.Xml;
using SpreadsheetLight;

namespace Vista.Controllers
{
    public class VistaController : Controller
    {


        public IActionResult Index()
        {
            // string xml = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"WikimediaExample.xml");
            string csv = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Libro1.csv");
            string xls = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Libro2.xlsx");
            //string cadenaXML = "";
            string cadenaXLS = "";


            cadenaXLS = LeerXls(xls);
            //cadenaXML = LeerXml(xml);
            //Console.WriteLine(cadenaXML);
            CabecerasClass obj = new CabecerasClass(csv);

            string cadenaCSV = obj.direc;

            string[] valores;

            if (cadenaXLS == "")
            {
                valores = cadenaCSV.Split(',');
            }

            else
            {
                valores = cadenaXLS.Split(',');
            }

            string[] ultimo = { null, null, null, null, null, null, null, null, null, null };
            for (int i = 0; i < valores.Length; i++)
            {
                ultimo[i] = valores[i];
            }

            string[] arreglo = { ultimo[0], ultimo[1], ultimo[2], ultimo[3], ultimo[4], ultimo[5], ultimo[6], ultimo[7], ultimo[8], ultimo[9] };
            string cad = "";
            for (int i = 0; i < (arreglo.Length); i++)
            {
                if (arreglo[i] == null)
                {
                    cad = cad + "," + "null";
                }
                else
                {
                    cad = cad + "," + arreglo[i];
                }

            }
            // int a=escribirNuevoXML(cad);
            CabecerasClass cabecera = new CabecerasClass { Arreglo = arreglo };


            return View(cabecera);
        }


        //Recibe datos de formulario
        [HttpPost]
        public IActionResult Index(IFormCollection formCollection)
        {
            //arreglo donde se guarda la seleccion del form
            string[] datos = new string[10];
            try
            {

                datos[0] = formCollection["Id"];
                datos[1] = formCollection["Name"];
                datos[2] = formCollection["LastName"];
                datos[3] = formCollection["Country"];
                datos[4] = formCollection["Phone"];
                datos[5] = formCollection["Addres"];
                datos[6] = formCollection["Email"];
                datos[7] = formCollection["Date"];
                datos[8] = formCollection["Dni"];
                datos[9] = formCollection["PostalCode"];


            }
            catch (NullReferenceException e)
            {
            }


            CabecerasClass selected = new CabecerasClass { Arreglo = datos };

            // Metodo para enviar string iría aca
            //string cad = "";

            //for (int i = 0; i < (datos.Length); i++)
            //{
            //    cad = cad + "," + datos[i];
            //}

            string cad = datos[0];

            for (int i = 1; i < (datos.Length); i++)
            {
                if (datos[i] == null)
                {
                    cad = cad + "," + "null";
                }
                else
                {
                    cad = cad + "," + datos[i];
                }

            }
            int a = escribirNuevoXML(cad);

            // escribirXML(cad);


            return Result(selected);
        }

        public IActionResult Result(CabecerasClass selected)
        {
            Console.WriteLine("------------.......................-");
            return View("Result");
        }


        public static void escribirXML(string cadena)
        {
            string[] arre = { "1", "2" };

            XmlDocument doc = new XmlDocument();
            for (int i = 0; i < (arre.Length); i++)
            {
                XmlElement raiz = doc.CreateElement(arre[i]);
            }
            doc.Save("NuevoXML.xml");
        }

        public static int escribirNuevoXML(string cadena)
        {

            string[] valores = cadena.Split(',');

            XmlWriter xmlWriter = XmlWriter.Create("Resources/NuevoXML.xml");

            xmlWriter.WriteStartDocument();
            xmlWriter.WriteStartElement(valores[0]);
            xmlWriter.WriteStartElement(valores[1]);
            xmlWriter.WriteStartElement(valores[2]);
            xmlWriter.WriteStartElement(valores[3]);
            xmlWriter.WriteStartElement(valores[4]);
            xmlWriter.WriteStartElement(valores[5]);
            xmlWriter.WriteStartElement(valores[6]);
            xmlWriter.WriteStartElement(valores[7]);
            xmlWriter.WriteStartElement(valores[8]);
            xmlWriter.WriteStartElement(valores[9]);
            xmlWriter.WriteEndDocument();
            xmlWriter.Close();
            int a = 1;
            return a;

        }

        public static string LeerXls(string full)
        {
            SLDocument documento = new SLDocument(full);
            SLWorksheetStatistics propiedades = documento.GetWorksheetStatistics();

            int ultimaColumna = propiedades.EndColumnIndex;

            ArrayList cabeza = new ArrayList();
            for (long i = 0; i < (ultimaColumna); i++)
            {
                string num = Letracolumna(i);
                string ca = documento.GetCellValueAsString(num + 1);
                cabeza.Add(ca);


            }
            List<string> re = new List<string>();

            foreach (string sw in cabeza)
            {
                if (!re.Contains(sw))
                {
                    re.Add(sw);
                }
            }

            string cad = re[0];

            for (int i = 1; i < (re.Count); i++)
            {
                cad = cad + "," + re[i];
            }
            Console.WriteLine(cad);
            return cad;


        }
        private static string Letracolumna(long col)
        {
            string columna = "";
            if (col < 26)
            {
                columna = ((char)(col + 65)).ToString();
            }
            else if (col < 52)
            {
                columna = ((char)65).ToString() + ((char)(col + 39)).ToString();
            }
            else if (col < 78)
            {
                columna = ((char)66).ToString() + ((char)(col + 13)).ToString();
            }
            else if (col < 104)
            {
                columna = ((char)67).ToString() + ((char)(col - 13)).ToString();
            }
            return columna;
        }

    }

}