using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Teco
{
    public class Loja
    {
        public string Nome { get; set; }
        public string Cidade { get; set; }

        public string CoordenadaKml { get; set; }

        /// <summary>
        /// Coordenada Polígono é responsável por criar associação de terrenos e concorrentes por loja
        /// </summary>
        public StringBuilder CoordenadaPoligono { get; set; }
        
        /// <summary>
        /// Coordenada Mapa é responsável por criar a loja no mapa
        /// </summary>
        public StringBuilder CoordenadaMapa { get; set; }

        public bool lojaComProblema { get; set; }


        public void preparaCoordenandaPoligono(string value)
        {
            try
            {
                CoordenadaPoligono = new StringBuilder(""); 
                CoordenadaPoligono = new StringBuilder(value.Replace(",", " ").Replace(" 0 ", ",")); // remove , e substitui o 0 por uma virgula
                CoordenadaPoligono = new StringBuilder( CoordenadaPoligono.ToString().Remove(CoordenadaPoligono.ToString().Length - 1));  // remove o ultimo char que e uma ,

            }
            catch (StackOverflowException e)
            {
                Console.WriteLine($"houve um erro com a loja {Nome}, da cidae {Cidade} da coordenanda {CoordenadaKml}");

            }
        }


        public void preparaCoordenandaMapa(string value)
        {
            try
            {
                CoordenadaMapa = new StringBuilder("");

                string[] coordenadas;
                List<string> mapa = new List<string>();
                string aux = value.Replace("0 ", ""); //.Remove(value.Length - 2, 1);
                string impar = string.Empty;
                string par = string.Empty;

                coordenadas = aux.Split(',');


                foreach (string item in coordenadas)
                {

                    if (coordenadas.ToList().IndexOf(item) % 2 == 0)
                    {
                        impar = item;
                    }
                    else
                    {
                        par = item;
                    }

                    if (!string.IsNullOrEmpty(par) && !string.IsNullOrEmpty(impar))
                    {
                        mapa.Add($"{par}, {impar}");
                        impar = par = string.Empty;
                    }
                }

                validaLojasQueFecham(mapa);

                CoordenadaMapa = new StringBuilder(string.Join("|", mapa));


                validaFechamentoDeLojasAposConversao(CoordenadaMapa, mapa);
                //var x = CoordenadaMapa.ToString().LastOrDefault();
                //CoordenadaMapa = CoordenadaMapa.Remove(CoordenadaPoligono.Length - 1); //.Remove(value.Length - 2, 1);

            }
            catch (StackOverflowException e)
            {
                Console.WriteLine($"houve um erro com a loja {Nome}, da cidae {Cidade} da coordenanda {CoordenadaKml}");

            }
        }

        private void validaFechamentoDeLojasAposConversao(StringBuilder coordenadaMapa, List<string> mapa)
        {
            var first = mapa.FirstOrDefault();

            // Loop through all instances of the string 'text'.
            int count = 0;
            int i = 0;
            while ((i = coordenadaMapa.ToString().IndexOf(first, i)) != -1)
            {
                i += first.Length;
                count++;
            }

           // if (count <= 1)
               // Console.WriteLine($"A conversao de coordenadas mapa da loja:'{Nome.ToUpper()}' está com problema, favor verificar!");

        }

        public void validaLojasQueFecham(List<string> coordenadas)
        {
            bool igual = coordenadas.LastOrDefault().ToString().Contains(coordenadas[0]);
            if (!igual)
            {
                lojaComProblema = true;
                var indexUltimo = coordenadas.IndexOf(coordenadas.LastOrDefault());
                coordenadas.Insert(indexUltimo + 1, coordenadas[0]);
            }
        }

    }


    public class LojaImportacao
    {
        public string Nome { get; set; }
        public string Cidade { get; set; }

        public string CoordenadaKml { get; set; }

        /// <summary>
        /// Coordenada Polígono é responsável por criar associação de terrenos e concorrentes por loja
        /// </summary>
        public string CoordenadaPoligono { get; set; }

        /// <summary>
        /// Coordenada Mapa é responsável por criar a loja no mapa
        /// </summary>
        public string CoordenadaMapa { get; set; }

        public bool lojaComProblema { get; set; }
    }


}
