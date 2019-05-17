using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace Teco
{
    internal class Program
    {
        private static List<Loja> lojas;
        private static List<LojaImportacao> lojasSemFim;
        private static List<LojaImportacao> lojasEmPlanilhaExcel;
        private static List<PlanilhaColunas> itensPlanilha;

        static string arquivoKml = @"C:\Users\thiag\Downloads\930693\930693-14lojas.kml";
        static FileInfo arquivoExcel = new FileInfo(@"C:\Users\thiag\Downloads\930693\lojasAgrupadas.xlsx");

        private static void Main(string[] args)
        {
            lojas = new List<Loja>();
            itensPlanilha = new List<PlanilhaColunas>();
            lojasSemFim = new List<LojaImportacao>();


            Console.WriteLine("Processando planilha \n");

            /*
            lerPlanilhaDeCoordenadasExcel();


            Console.WriteLine("Processamento da planilha com o xml \n");
            imprimePlanilhaExcel(lojasEmPlanilhaExcel, lojasSemFim);

            */
            Console.WriteLine("Processando kml \n");
            lerXml();

            /*
            Console.WriteLine("Kml processado... \n");

            Console.WriteLine("Lendo planilha... \n");

            lerPlanilhaQueContemNomeDasLojas();

            Console.WriteLine("Processamento da planilha com o xml \n");

            processaNovaPlanilha(lojas, itensPlanilha);
            */

            imprimeXml(lojas);

            Console.WriteLine("Processamento da nova planilha... \n");

            

           Console.WriteLine("Aperte uma tecla para finalizar \n");

            Console.ReadLine();


        }

        

        private static void lerPlanilhaDeCoordenadasExcel()
        {
            using (ExcelPackage xlPackage = new ExcelPackage(arquivoExcel))
            {
                ExcelWorksheet myWorksheet = xlPackage.Workbook.Worksheets.LastOrDefault(); //select sheet here
                int totalRows = myWorksheet.Dimension.End.Row;
                int totalColumns = myWorksheet.Dimension.End.Column;

                lojasEmPlanilhaExcel = new List<LojaImportacao>();

                StringBuilder sb = new StringBuilder(); //this is your your data
                for (int rowNum = 2; rowNum <= totalRows; rowNum++) //selet starting row here
                {
                    Loja lojasDaPlanilha = new Loja();
                    LojaImportacao li = new LojaImportacao();
                    IEnumerable<string> row = myWorksheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());

                    string s = string.Join(",", row);
                    //Console.WriteLine(s);
                    li.Nome = row.ToList()[0];
                    li.Cidade = row.ToList()[1];
                    li.CoordenadaKml = row.ToList()[2];

                    lojasDaPlanilha.preparaCoordenandaPoligono(li.CoordenadaKml);
                    lojasDaPlanilha.preparaCoordenandaMapa(li.CoordenadaKml);


                    li.CoordenadaPoligono = lojasDaPlanilha.CoordenadaPoligono.ToString();
                    li.CoordenadaMapa = lojasDaPlanilha.CoordenadaMapa.ToString();
                    li.lojaComProblema = lojasDaPlanilha.lojaComProblema;


                    lojasEmPlanilhaExcel.Add(li);
                }

                lojasSemFim = lojasEmPlanilhaExcel.Where(x => x.lojaComProblema == true).ToList();
            }
        }

        private static void lerXml()
        {
            /*
#pragma warning disable CS0618 // Type or member is obsolete
            XmlDataDocument kmlDoc = new XmlDataDocument();
#pragma warning restore CS0618 // Type or member is obsolete
            FileStream fs = new FileStream(@"C:\Users\thiago.hpereira\Documents\Visual Studio 2017\Projects\Teco\Lojas com menos pontos - V_Final.kml", FileMode.Open, FileAccess.Read);
            kmlDoc.Load(fs);
            */

            //XmlReader reader = XmlReader.Create(@"C:\Users\thiago.hpereira\Documents\Visual Studio 2017\Projects\Teco\4 - Zona Sul - BELO HORIZONTE.kml");
            XmlReader reader = XmlReader.Create(arquivoKml);
            string cidade = string.Empty;

            while (reader.Read())
            {
                reader.Read();
                Loja loja = new Loja();



                while (!(string.Equals(reader.Name.Trim('\t', '\r', '\n'), "Placemark") && reader.NodeType == XmlNodeType.EndElement))
                {
                    if (string.Equals(reader.Name.Trim('\t', '\r', '\n'), String.Empty) && reader.NodeType == XmlNodeType.None)
                    {
                        break;
                    }

                    if (string.Equals(reader.Name.Trim('\t', '\r', '\n'), "Placemark") && reader.NodeType == XmlNodeType.Element)
                    {
                        loja.Cidade = cidade;
                    }

                    if (reader.Name == "Document")
                    {
                        do
                        {
                            reader.Read();
                            if (!string.IsNullOrEmpty(reader.Value.Trim('\t', '\r', '\n')))
                            {
                                cidade = reader.Value;
                                loja.Cidade = reader.Value;
                            }

                            if ((string.Equals(reader.Name.Trim('\t', '\r', '\n'), "kml") && reader.NodeType == XmlNodeType.EndElement))
                            {
                                break;
                            }
                        } while (loja.Cidade == null);
                    }

                    if ((reader.NodeType == XmlNodeType.Element) && (reader.Name == "Placemark"))
                    {
                        do
                        {
                            reader.Read();
                            if (!string.IsNullOrEmpty(reader.Value.Trim('\t', '\r', '\n')))
                            {
                                loja.Nome = reader.Value;
                            }
                        } while (loja.Nome == null);
                    }
                    if ((reader.NodeType == XmlNodeType.Element) && (reader.Name == "coordinates"))
                    {
                        do
                        {
                            reader.Read();
                            if (!string.IsNullOrEmpty(reader.Value.Trim('\t', '\r', '\n')))
                            {
                                loja.CoordenadaKml = reader.Value.Trim('\t', '\r', '\n');
                                loja.preparaCoordenandaPoligono(loja.CoordenadaKml);
                                loja.preparaCoordenandaMapa(loja.CoordenadaKml);
                            }
                        } while (loja.CoordenadaKml == null);
                    }
                    reader.Read();
                }

                if (string.Equals(reader.Name.Trim('\t', '\r', '\n'), String.Empty) && reader.NodeType == XmlNodeType.None)
                {
                    break;
                }

                lojas.Add(loja);

                //Console.WriteLine($"loja:  {lojas.IndexOf(loja) + 1} \n Nome: {loja.Nome} \n Cidade: {loja.Cidade} \n coordenada: {loja.CoordenadaKml.Substring(0, 80)}... \n");
            }

        }

        private static void lerPlanilhaQueContemNomeDasLojas()
        {
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(@"C:\Users\thiago.hpereira\Documents\Visual Studio 2017\Projects\Teco\Adequacaodasmicrorregioes-2019.xlsx")))
            {
                ExcelWorksheet myWorksheet = xlPackage.Workbook.Worksheets.First(); //select sheet here
                int totalRows = myWorksheet.Dimension.End.Row;
                int totalColumns = myWorksheet.Dimension.End.Column;

                StringBuilder sb = new StringBuilder(); //this is your your data

                for (int rowNum = 2; rowNum <= totalRows; rowNum++) //selet starting row here
                {
                    PlanilhaColunas planilha = new PlanilhaColunas();

                    IEnumerable<string> row = myWorksheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());

                    string s = string.Join(",", row);
                    Console.WriteLine(s);
                    planilha.potencial = row.ToList()[0];
                    planilha.loja = row.ToList()[1];
                    planilha.cidade = row.ToList()[2];
                    planilha.regional = row.ToList()[3];
                    planilha.nomereduzido = row.ToList()[4];

                    itensPlanilha.Add(planilha);
                }
            }
        }


        private static void processaNovaPlanilha(List<Loja> lojas, List<PlanilhaColunas> itensPlanilha)
        {

            List<PlanilhaCoordenadas> planilhas = new List<PlanilhaCoordenadas>();

            IEnumerable<PlanilhaCoordenadas> planilhaJoin = from l in lojas
                                                            from p in itensPlanilha
                                                            where p.loja.ToUpper() == l.Nome.ToUpper()
                                                            //&& r.cidade.ToUpper() == p.Cidade.ToUpper()
                                                            select new PlanilhaCoordenadas()
                                                            {
                                                                CidadeLoja = l.Cidade,
                                                                CoordenadaKml = l.CoordenadaKml,
                                                                CoordenadaMapa = l.CoordenadaMapa.ToString(),
                                                                CoordenadaPoligono = l.CoordenadaPoligono.ToString(),
                                                                NomeLoja = l.Nome,
                                                                CidadeNome = p.cidade,
                                                                LojaNome = p.loja,
                                                                Nomereduzido = p.nomereduzido,
                                                                Potencial = p.potencial,
                                                                Regional = p.regional
                                                            } as PlanilhaCoordenadas;

            bool planilhaExcluida = planilhaJoin.Select(x => x.NomeLoja).FirstOrDefault().Contains(lojas.Select(u => u.Nome).FirstOrDefault());
            IEnumerable<Loja> lojasQueNaoForamCruzadasEntreKmlExcel = lojas.Where(p => !planilhaJoin.Any(p2 => p2.NomeLoja == p.Nome));


            ///coorriigir
            //lojasSemFim = lojas.Where(x => x.lojaComProblema == true).ToList();


            List<LojaImportacao> lojasImportacao = new List<LojaImportacao>();
            foreach (Loja item in lojas)
            {
                lojasImportacao.Add(new LojaImportacao()
                {
                    Cidade = item.Cidade,
                    CoordenadaKml = item.CoordenadaKml,
                    CoordenadaMapa = item.CoordenadaMapa.ToString(),
                    CoordenadaPoligono = item.CoordenadaPoligono.ToString(),
                    Nome = item.Nome,
                    lojaComProblema = item.lojaComProblema
                });
            }

            List<LojaImportacao> lojasImportacaoSemFim = new List<LojaImportacao>();

            foreach (LojaImportacao item in lojasSemFim)
            {
                lojasImportacaoSemFim.Add(new LojaImportacao()
                {
                    Nome = item.Nome,
                    Cidade = item.Cidade,
                    CoordenadaKml = item.CoordenadaKml,
                    CoordenadaMapa = item.CoordenadaMapa,
                    CoordenadaPoligono = item.CoordenadaPoligono,
                    lojaComProblema = item.lojaComProblema
                });

            }

            imprimePlanilhaConcatenada(planilhaJoin, lojasImportacaoSemFim, lojasSemFim);

        }

        private static void imprimeXml(List<Loja> lojas)
        {
            List<LojaImportacao> lojasImportacao = new List<LojaImportacao>();
            foreach (Loja item in lojas)
            {
                lojasImportacao.Add(new LojaImportacao()
                {
                    Cidade = item.Cidade,
                    CoordenadaKml = item.CoordenadaKml,
                    CoordenadaMapa = item.CoordenadaMapa.ToString(),
                    CoordenadaPoligono = item.CoordenadaPoligono.ToString(),
                    Nome = item.Nome,
                    lojaComProblema = item.lojaComProblema
                });
            }

            List<LojaImportacao> lojasImportacaoSemFim = new List<LojaImportacao>();

            foreach (LojaImportacao item in lojasSemFim)
            {
                lojasImportacaoSemFim.Add(new LojaImportacao()
                {
                    Nome = item.Nome,
                    Cidade = item.Cidade,
                    CoordenadaKml = item.CoordenadaKml,
                    CoordenadaMapa = item.CoordenadaMapa,
                    CoordenadaPoligono = item.CoordenadaPoligono,
                    lojaComProblema = item.lojaComProblema
                });

            }
            imprimePlanilhaExcel(lojasImportacao, lojasImportacaoSemFim);
        }

        private static void imprimePlanilhaConcatenada(IEnumerable<PlanilhaCoordenadas> planilhaJoin, IEnumerable<LojaImportacao> lojas, IEnumerable<LojaImportacao> lojaSemFim)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Planilha comparada com KML");
                excel.Workbook.Worksheets.Add("Worksheet2");
                excel.Workbook.Worksheets.Add("Lojas Sem Fim no Mapa");

                List<string[]> headerRowPlanilha1 = new List<string[]>()
                                          {
                                            new string[] {
                                                "Loja",
                                                "Cidade",
                                                "coordenada KML",
                                                "coordenanda Poligono",
                                                "coordenada Mapa",
                                                "potencial",
                                                "loja",
                                                "cidade",
                                                "regional",
                                                "nome reduzido"
                                          } };

                List<string[]> headerRowPlanilha2 = new List<string[]>()
                                          {
                                            new string[] {
                                                "Loja",
                                                "Cidade",
                                                "coordenada KML",
                                                "coordenanda Poligono",
                                                "coordenada Mapa"
                                          } };

                // Determine the header range (e.g. A1:D1)
                string headerRange1 = "A1:" + Char.ConvertFromUtf32(headerRowPlanilha1[0].Length + 64) + "1";
                string headerRange2 = "A1:" + Char.ConvertFromUtf32(headerRowPlanilha2[0].Length + 64) + "1";

                // Target a worksheet
                ExcelWorksheet worksheet1 = excel.Workbook.Worksheets["Planilha comparada com KML"];
                ExcelWorksheet worksheet2 = excel.Workbook.Worksheets["Worksheet2"];
                ExcelWorksheet worksheet3 = excel.Workbook.Worksheets["Lojas Sem Fim no Mapa"];

                // Popular header row data
                worksheet1.Cells[headerRange1].LoadFromArrays(headerRowPlanilha1);
                worksheet2.Cells[headerRange2].LoadFromArrays(headerRowPlanilha2);


                worksheet1.Cells[2, 1].LoadFromCollection<PlanilhaCoordenadas>(planilhaJoin);
                worksheet2.Cells[2, 1].LoadFromCollection<LojaImportacao>(lojas);

                if (lojaSemFim.Count() > 0)
                {
                    worksheet3.Cells[headerRange2].LoadFromArrays(headerRowPlanilha2);
                    worksheet3.Cells[2, 1].LoadFromCollection<LojaImportacao>(lojaSemFim);
                }

                FileInfo excelFile = new FileInfo(@"C:\Users\thiago.hpereira\Documents\Visual Studio 2017\Projects\Teco\Lojas18-03-19.xlsx");
                excel.SaveAs(excelFile);
            }
        }

        private static void imprimePlanilhaExcel(IEnumerable<LojaImportacao> planilha, List<LojaImportacao> lojasQueNaoFecham)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("PlanilhaCorrigida");
                excel.Workbook.Worksheets.Add("PlanilhaSemFim");


                List<string[]> headerRowPlanilha1 = new List<string[]>()
                                          {
                                            new string[] {
                                                "Loja",
                                                "Cidade",
                                                "coordenada KML",
                                                "coordenanda Poligono",
                                                "coordenada Mapa",
                                          } };

                // Determine the header range (e.g. A1:D1)
                string headerRange1 = "A1:" + Char.ConvertFromUtf32(headerRowPlanilha1[0].Length + 64) + "1";
                string headerRange2 = "A1:" + Char.ConvertFromUtf32(headerRowPlanilha1[0].Length + 64) + "1";

                // Target a worksheet
                ExcelWorksheet worksheet1 = excel.Workbook.Worksheets["PlanilhaCorrigida"];
                //ExcelWorksheet worksheet2 = excel.Workbook.Worksheets["PlanilhaSemFim"];

                // Popular header row data
                worksheet1.Cells[headerRange1].LoadFromArrays(headerRowPlanilha1);
            //    worksheet2.Cells[headerRange2].LoadFromArrays(headerRowPlanilha1);


                worksheet1.Cells[2, 1].LoadFromCollection<LojaImportacao>(planilha);
                //   worksheet2.Cells[2, 1].LoadFromCollection<LojaImportacao>(lojasQueNaoFecham);


                FileInfo excelFile = arquivoExcel;
                excel.SaveAs(excelFile);
            }
        }


    }
}