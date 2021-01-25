using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TransformXLSX
{
    class Program
    {
        static void Main(string[] args)
        {

#if !DEBUG
            if (args.Length == 0) return;

            if (!File.Exists(args[0]))
            {
                Console.WriteLine("Arquivo não existe ou sem permissão");
                return;
            }
#endif

#if DEBUG
            args = new[] { "C:\\Users\\aasf_\\Downloads\\Meta2-de-2021.xlsx" };
#endif

            var listRows = new List<LinhaPlanilha>();

            using (var file = new FileStream(args[0], FileMode.Open, FileAccess.Read))
            {                
                var xs = new XSSFWorkbook(file);
                var sheet = xs.GetSheetAt(0);
                
                if (sheet != null)
                {                    
                    for (int r = 0; r < sheet.LastRowNum; r++)
                    {
                        var row = sheet.GetRow(r);
                        if (row != null)
                        {
                            var linha = new LinhaPlanilha();

                            linha.Sequencia = row.GetCell(0).GetValue();
                            linha.Grau = row.GetCell(1).GetValue();
                            linha.Entrancia = row.GetCell(2).GetValue();
                            linha.Esfera = row.GetCell(3).GetValue();
                            linha.Jurisdicao = row.GetCell(4).GetValue();
                            linha.OrgaoJulgador = row.GetCell(5).GetValue();
                            linha.OrgaoJulgadorColegiado = row.GetCell(6).GetValue();
                            linha.TipoReparticao = row.GetCell(7).GetValue();
                            linha.Fase = row.GetCell(8).GetValue();
                            linha.NumeroUnico = row.GetCell(9).GetValue();
                            linha.Protocolo = row.GetCell(10).GetValue();
                            linha.AnoProcesso = row.GetCell(11).GetValue();
                            linha.DataProtocolo = row.GetCell(12).GetValue();
                            linha.Classe = row.GetCell(13).GetValue();
                            linha.Sistema = row.GetCell(14).GetValue();
                            linha.EletronicoFisicoHibrido = row.GetCell(15).GetValue();
                            linha.ValorCausa = row.GetCell(16).GetValue();
                            linha.JusticaGratuita = row.GetCell(17).GetValue();
                            linha.EstoqueMeta2 = row.GetCell(18).GetValue();

                            listRows.Add(linha);

                            //Console.WriteLine("");
                        }                        
                    }
                }
            }


            //Escrever novas planilhas...

            var cabecalho = listRows[0];
            listRows.RemoveAt(0);

            var results = listRows.GroupBy(p => p.Jurisdicao + " - " + p.OrgaoJulgador, (key, g) => new { key, Linhas = g.ToList() });
            var dirPlanilhas = Path.Combine(Environment.CurrentDirectory, Process.GetCurrentProcess().ProcessName);
            
            if (!Directory.Exists(dirPlanilhas)) Directory.CreateDirectory(dirPlanilhas);

            foreach (var item in results)
            {
                Console.WriteLine($"key: { item.key}; linhas: {item.Linhas.Count}");

                var xwb = new XSSFWorkbook();

                var sheet = xwb.CreateSheet();
                var rowHead = sheet.CreateRow(0);

                //cabecalho
                rowHead.CreateCell(0).SetCellValue(cabecalho.Sequencia);
                rowHead.CreateCell(1).SetCellValue(cabecalho.Grau);
                rowHead.CreateCell(2).SetCellValue(cabecalho.Entrancia);
                rowHead.CreateCell(3).SetCellValue(cabecalho.Esfera);
                rowHead.CreateCell(4).SetCellValue(cabecalho.Jurisdicao);
                rowHead.CreateCell(5).SetCellValue(cabecalho.OrgaoJulgador);
                rowHead.CreateCell(6).SetCellValue(cabecalho.TipoReparticao);
                rowHead.CreateCell(7).SetCellValue(cabecalho.Fase);
                rowHead.CreateCell(8).SetCellValue(cabecalho.NumeroUnico);
                rowHead.CreateCell(9).SetCellValue(cabecalho.Protocolo);
                rowHead.CreateCell(10).SetCellValue(cabecalho.AnoProcesso);
                rowHead.CreateCell(11).SetCellValue(cabecalho.DataProtocolo);
                rowHead.CreateCell(12).SetCellValue(cabecalho.Classe);
                rowHead.CreateCell(13).SetCellValue(cabecalho.Sistema);
                rowHead.CreateCell(14).SetCellValue(cabecalho.EletronicoFisicoHibrido);
                rowHead.CreateCell(15).SetCellValue(cabecalho.ValorCausa);
                rowHead.CreateCell(16).SetCellValue(cabecalho.JusticaGratuita);


                //linhas
                for (int i = 0; i < item.Linhas.Count; i++)
                {
                    var itemLinha = item.Linhas[i];
                    var row = i + 1;
                    var rowItem = sheet.CreateRow(row);
                    rowItem.CreateCell(0).SetCellValue(itemLinha.Sequencia);
                    rowItem.CreateCell(1).SetCellValue(itemLinha.Grau);
                    rowItem.CreateCell(2).SetCellValue(itemLinha.Entrancia);
                    rowItem.CreateCell(3).SetCellValue(itemLinha.Esfera);
                    rowItem.CreateCell(4).SetCellValue(itemLinha.Jurisdicao);
                    rowItem.CreateCell(5).SetCellValue(itemLinha.OrgaoJulgador);
                    rowItem.CreateCell(6).SetCellValue(itemLinha.TipoReparticao);
                    rowItem.CreateCell(7).SetCellValue(itemLinha.Fase);
                    rowItem.CreateCell(8).SetCellValue(itemLinha.NumeroUnico);
                    rowItem.CreateCell(9).SetCellValue(itemLinha.Protocolo);
                    rowItem.CreateCell(10).SetCellValue(itemLinha.AnoProcesso);
                    rowItem.CreateCell(11).SetCellValue(itemLinha.DataProtocolo);
                    rowItem.CreateCell(12).SetCellValue(itemLinha.Classe);
                    rowItem.CreateCell(13).SetCellValue(itemLinha.Sistema);
                    rowItem.CreateCell(14).SetCellValue(itemLinha.EletronicoFisicoHibrido);
                    rowItem.CreateCell(15).SetCellValue(itemLinha.ValorCausa);
                    rowItem.CreateCell(16).SetCellValue(itemLinha.JusticaGratuita);
                }

                using (MemoryStream stream = new MemoryStream())
                {
                    xwb.Write(stream);

                    var uriPlanilha = Path.Combine(dirPlanilhas, item.key.Replace("//","-").Replace("\\","-").Replace("/", "-").Replace(@"\","-") + ".xlsx");

                    if (File.Exists(uriPlanilha)) File.Delete(uriPlanilha);

                    File.WriteAllBytes(uriPlanilha, stream.ToArray());

                    Console.WriteLine(uriPlanilha);
                }
            }

            Console.ReadKey();
        }
    }
}
