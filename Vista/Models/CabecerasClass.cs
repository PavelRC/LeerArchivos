using System;
using System.Collections.Generic;
using System.IO;

namespace Vista.Models
{
    public class CabecerasClass
    {
        public string[] Arreglo { get; set; }
        public List <string> Lista { get; set; }
        public string direc { get; set; }


        public CabecerasClass()
        {

        }

        public CabecerasClass(string[] Arr, List<string> Li)
        {
            this.Arreglo = Arr;
            this.Lista = Li;
        }


        public CabecerasClass(string dir)
        {
            using (FileStream fs = File.OpenRead(dir))
            {
                var reader = new StreamReader(fs);

                List<String> lista = new List<string>();
                int cont = 0;
                List<string> re = new List<string>();
                while (!reader.EndOfStream)
                {
                    var linea = reader.ReadLine();
                    var valores = linea.Split(';');

                    if (cont == 0)
                    {

                        foreach (string sw in valores)
                        {
                            if (!re.Contains(sw))
                            {
                                re.Add(sw);
                            }
                        }

                        cont = 1;
                    }
                
                }

                string cad = re[0];
                for (int i = 1; i < (re.Count); i++)
                {
                    cad = cad + "," + re[i];
                }
                //return cad;
                this.direc = cad;
            }


        }





    }

    
}
