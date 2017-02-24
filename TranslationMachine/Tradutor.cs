using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.Office.Core;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TranslationMachine
{
    class Tradutor
    {
        private float SlideHeight;
        private float SlideWidth;
        private PowerPoint.Presentation presentation;

        //Traduz tudo, caga pra tudo
        private bool rage = false;
        
        public void Run ()
        {
            Console.SetWindowPosition(0, 0);
            Console.OutputEncoding = Encoding.Unicode;

            Console.WriteLine("Console Tradutor de IBM-TextBook para PDF V 1.0.0");
            Console.WriteLine("Go-Horse Edition");

            // Obtém o modo 
            Console.WriteLine("Quer automatizar saporra??? [Y] (Qql outra coisa pra não)");
            if (Console.ReadLine() == "Y") rage = true;
            
            // Cria um objeto de apresentação de PowerPoint
            PowerPoint.Application app = new PowerPoint.Application();
            PowerPoint.Presentations presentations = app.Presentations;

            // Busca os arquivos de PowerPoint e pdf para ser lidos para essa transformação
            string basePath = Environment.CurrentDirectory;
            string inputPath = System.IO.Path.Combine(basePath, "apresentacao.pptx");

            string outFile = "apresentacao.pptx";

            if (rage) outFile = "rage.pptx";

            while (!File.Exists(inputPath))
            {
                Console.WriteLine("Não foi possível encontrar o arquivo especificado (padrão -> apresentacao.pptx). Digite um arquivo local pptx para ser lido");
                inputPath = System.IO.Path.Combine(basePath, Console.ReadLine());
            }

            string input2path = System.IO.Path.Combine(basePath, "consulta.pdf");
            while (!File.Exists(input2path))
            {
                Console.WriteLine("Não foi possível encontrar o arquivo especificado (padrão -> consulta.pdf). Digite um arquivo local pdf para ser lido");
                input2path = System.IO.Path.Combine(basePath, Console.ReadLine());
            }

            Console.WriteLine("Abrindo pptx de: {0}", inputPath);
            Console.WriteLine("Será consultado pdf de: {0}", input2path);

            // Lê o powerpoint
            presentation = presentations.Open(inputPath, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);
            
            // Obtém informações sobre setup dos slides
            SlideHeight = presentation.PageSetup.SlideHeight;
            SlideWidth = presentation.PageSetup.SlideWidth;

            // Obtém quais slides devem ser varridos
            Console.Write("Digite ENTER para varrer todos slides ou um número para começar (1,2,3..)");
            int slideInicial = 1, slideFinal = presentation.Slides.Count;

            string inputSlides = Console.ReadLine();
            if (inputSlides != "" && int.TryParse(inputSlides, out slideInicial))
            {
                LogWithColor(ConsoleColor.DarkCyan, "Digite o slide final: ");

                inputSlides = Console.ReadLine();
                if (inputSlides != "") int.TryParse(inputSlides, out slideFinal);
            }

            // Persistência do número de slides
            if (slideInicial < 1 || slideInicial > presentation.Slides.Count) slideInicial = 1;
            if (slideFinal < 1 || slideFinal > presentation.Slides.Count || slideFinal < slideInicial) slideFinal = presentation.Slides.Count;

            LogWithColor(ConsoleColor.DarkCyan, "Varrendo slides de " + slideInicial + " à " + slideFinal);

            // Varre slide a slide
            for (int i = slideInicial; i < slideFinal + 1; i++)
            {
                PowerPoint.Slide slide = presentation.Slides[i];

                // Varre forma a forma. Basicamente cada "coisa" no PowerPoint é um shape
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    // Se for um textbox que comece com "Pag "
                    if (shape.Type == MsoShapeType.msoTextBox && shape.TextFrame.TextRange.Text.Contains("Pag "))
                    {
                        Console.WriteLine("[DEBUG] Substituição no Slide {0}: {1}", i, shape.TextFrame.TextRange.Text);

                        Regex regex = new Regex(@"[^ \, a-zA-Z]\d+");
                        MatchCollection matches = regex.Matches(shape.TextFrame.TextRange.Text);

                        // Pega as páginas referenciadas naquele textbox
                        List<int> pdfPages = new List<int>();
                        foreach (Match match in matches)
                        {
                            foreach (Capture capture in match.Captures)
                            {
                                pdfPages.Add(int.Parse(capture.Value));
                            }
                        }

                        // Para cada página referenciada do PDF
                        foreach (int page in pdfPages)
                        {
                            // Pega textos da página
                            List<string> textos = GetTextFromPDF(page, input2path);

                            Console.WriteLine("[DEBUG]     Na página {0} do pdf", page);
                            string glue = "";
                            for (int j = 0; j < textos.Count; j++)
                            {
                                Regex regexPasso = new Regex(@"[a-z\d]{1,2}\. ");

                                string resto = "";
                                //__ 1., __a. etc
                                // Se for um texto esperado
                                if (regexPasso.IsMatch(textos[j]))
                                {
                                    resto = regexPasso.Match(textos[j]).Captures[0].Value;
                                    textos[j] = textos[j].Substring(resto.Length);
                                    textos[j] = textos[j].Replace('\n', ' ');
                                    resto = resto.Substring(0, resto.Length - 2);
                                }
                                else
                                {
                                    // Gambiarra para saber que não era esperado
                                    resto = "***";
                                }

                                string traducao = TranslateEnToPtBr(textos[j]);

                                // Se for rage mode
                                //      Traduz tudo não pergunta nada
                                //      Inseri tudo
                                // Caso contrário, espera um comando
                                if (rage)
                                {
                                    glue += traducao + "\n";
                                    // Exibe o original

                                    LogWithColor(ConsoleColor.Green, "[ORIGINAL] " + textos[j]);

                                    // Exibe a tradução obtida do Google Tradutor
                                    Console.Write("           ");
                                    for (int chCount = 0; chCount < traducao.Length; chCount += 10)
                                        Console.Write(chCount + new string(' ', 10 - chCount.ToString().Length));
                                    Console.WriteLine();
                                    //      Vermelho é pra "avisar que não era esperado esse texto provavelmente
                                    if (resto == "***") LogWithColor(ConsoleColor.DarkRed, "[TRADUCAO] " + traducao);
                                    else LogWithColor(ConsoleColor.Magenta, "[TRADUCAO] " + traducao);

                                    if (j == textos.Count - 2)
                                    {
                                        PowerPoint.Shape textBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
                                        15, 50, 500, 300);
                                        ConfigureTextBoxFont(resto, glue, textBox.TextFrame.TextRange);
                                        //Rectangle r = ConfigureTextBoxWithinSlide(textBox, slide);
                                        //textBox.Left = r.Left;
                                        //textBox.Top = r.Top;
                                    }
                                    else if (j == textos.Count - 1)
                                    {
                                        PowerPoint.Shape textBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
                                        15, 50, 500, 300);
                                        ConfigureTextBoxFont(resto, traducao, textBox.TextFrame.TextRange);
                                        //Rectangle r = ConfigureTextBoxWithinSlide(textBox, slide);
                                        //textBox.Left = r.Left;
                                        //textBox.Top = r.Top;
                                    }


                                }
                                // Modo de edição
                                else
                                {
                                    //Para manter historico de edicoes, permitir redo e undo
                                    List<string> passosEdicao = new List<string>();
                                    passosEdicao.Add(traducao);
                                    int passo = 0;

                                    // Exibe o original
                                    Console.Write("           ");
                                    for (int chCount = 0; chCount < textos[j].Length; chCount += 10)
                                        Console.Write(chCount + new string(' ', 10 - chCount.ToString().Length));
                                    Console.WriteLine();
                                    LogWithColor(ConsoleColor.Green, "[ORIGINAL] " + textos[j]);

                                    // Imprime tradução obtida do Google Tradutor
                                    Console.Write("           ");
                                    for (int chCount = 0; chCount < passosEdicao[passo].Length; chCount += 10)
                                        Console.Write(chCount + new string(' ', 10 - chCount.ToString().Length));
                                    Console.WriteLine();
                                    //      Vermelho é pra "avisar que não era esperado esse texto provavelmente
                                    if (resto == "***") LogWithColor(ConsoleColor.DarkRed, "[TRADUCAO] " + passosEdicao[passo]);
                                    else LogWithColor(ConsoleColor.Magenta, "[TRADUCAO] " + passosEdicao[passo]);

                                    //Pega uma entrada do usuário
                                    string choice = Console.ReadLine();

                                    // Enquanto for comando de 1 caractere ou comando para substituir tudo ( > 1 caractere)
                                    // Lista de comandos
                                    // d - DELETE
                                    // u - UNDO
                                    // r - REDO
                                    // s - SUBSTITUTE
                                    // i - INSERT
                                    // c - CLEAR
                                    // a - ABORT
                                    // t - TRANSLATE
                                    // k - SKIP (pula Slide)
                                    // T - TRANSLATE2 (traduz algum texto do usuário e faz nada)
                                    // ENTER - ACCEPT
                                    // Qualquer string com mais de 1 caractere - Substitui a tradução pela string

                                    while (choice.Length > 0)
                                    {
                                        // Mais que 1 caractere troca a tradução pelo input 
                                        if (choice.Length > 1)
                                        {
                                            LogWithColor(ConsoleColor.DarkBlue, "Tradução substituída");
                                            passo++;
                                            if (passo >= passosEdicao.Count) passosEdicao.Add(choice);
                                            else passosEdicao[passo] = choice;
                                        }
                                        // Comando DELETE
                                        else if (choice == "d")
                                        {
                                            LogWithColor(ConsoleColor.DarkBlue, "Comando DELETE");

                                            // Pega posição de inicio
                                            string posIni = "a";
                                            int iPosIni;
                                            while (!int.TryParse(posIni, out iPosIni))
                                            {
                                                LogWithColor(ConsoleColor.DarkBlue, "Digite a pos de inicio: ");
                                                posIni = Console.ReadLine();
                                            }
                                            if (iPosIni < 0) iPosIni = 0;

                                            // Pega posição de fim
                                            string posFim = "a";
                                            int iPosFim;
                                            while (!int.TryParse(posFim, out iPosFim))
                                            {
                                                LogWithColor(ConsoleColor.DarkBlue, "Digite a pos de fim: ");
                                                posFim = Console.ReadLine();
                                            }

                                            // Adiciona a edição
                                            string edicao;
                                            if (iPosFim == -1) edicao = passosEdicao[passo].Substring(0, iPosIni);
                                            else edicao = passosEdicao[passo].Substring(0, iPosIni) + passosEdicao[passo].Substring(iPosFim + 1);

                                            passo++;
                                            if (passo >= passosEdicao.Count) passosEdicao.Add(edicao);
                                            else passosEdicao[passo] = edicao;
                                        }
                                        // Comando UNDO
                                        else if (choice == "u")
                                        {
                                            LogWithColor(ConsoleColor.DarkBlue, "Comando UNDO");
                                            if (passo > 0)
                                                passo--;
                                        }
                                        // Comando REDO
                                        else if (choice == "r")
                                        {
                                            LogWithColor(ConsoleColor.DarkBlue, "Comando REDO");
                                            if (passo < passosEdicao.Count - 1)
                                                passo++;
                                        }
                                        // Comando SUBSTITUTE
                                        else if (choice == "s")
                                        {
                                            LogWithColor(ConsoleColor.DarkBlue, "Comando SUBSTITUTE");

                                            // Pega posição de inicio
                                            string posIni = "a";
                                            int iPosIni;
                                            while (!int.TryParse(posIni, out iPosIni))
                                            {
                                                LogWithColor(ConsoleColor.DarkBlue, "Digite a pos de inicio: ");
                                                posIni = Console.ReadLine();
                                            }
                                            if (iPosIni < 0) iPosIni = 0;

                                            // Pega posição de fim
                                            string posFim = "a";
                                            int iPosFim;
                                            while (!int.TryParse(posFim, out iPosFim))
                                            {
                                                LogWithColor(ConsoleColor.DarkBlue, "Digite a pos de fim: ");
                                                posFim = Console.ReadLine();
                                            }
                                            // Pega a nova string
                                            LogWithColor(ConsoleColor.DarkBlue, "Digite a substituição: ");
                                            string subs = Console.ReadLine();

                                            // Adiciona a edição
                                            string edicao;
                                            if (iPosFim == -1) edicao = passosEdicao[passo].Substring(0, iPosIni) + subs;
                                            else edicao = passosEdicao[passo].Substring(0, iPosIni) + subs + passosEdicao[passo].Substring(iPosFim + 1);

                                            passo++;
                                            if (passo >= passosEdicao.Count) passosEdicao.Add(edicao);
                                            else passosEdicao[passo] = edicao;

                                        }
                                        // Comando INSERT
                                        else if (choice == "i")
                                        {
                                            LogWithColor(ConsoleColor.DarkBlue, "Comando INSERT");

                                            // Pega posição de inicio
                                            string posIni = "a";
                                            int iPosIni;
                                            while (!int.TryParse(posIni, out iPosIni))
                                            {
                                                LogWithColor(ConsoleColor.DarkBlue, "Digite a posição (-1 é no fim): ");
                                                posIni = Console.ReadLine();
                                            }

                                            // Pega a nova string
                                            LogWithColor(ConsoleColor.DarkBlue, "Digite a inserção: ");
                                            string insercao = Console.ReadLine();

                                            // Adiciona a edição
                                            string edicao;
                                            if (iPosIni != -1) edicao = passosEdicao[passo].Substring(0, iPosIni) + insercao + passosEdicao[passo].Substring(iPosIni);
                                            else edicao = passosEdicao[passo] + insercao;

                                            passo++;
                                            if (passo >= passosEdicao.Count) passosEdicao.Add(edicao);
                                            else passosEdicao[passo] = edicao;

                                        }
                                        // Comando CLEAR
                                        else if (choice == "c")
                                        {
                                            LogWithColor(ConsoleColor.DarkBlue, "Comando CLEAR");

                                            // Adiciona a edição
                                            passo++;
                                            if (passo >= passosEdicao.Count) passosEdicao.Add("");
                                            else passosEdicao[passo] = "";
                                        }
                                        // Comando ABORT
                                        // Aborta a varredura mas salva todas traduções feitas
                                        else if (choice == "a")
                                        {
                                            LogWithColor(ConsoleColor.DarkBlue, "Comando ABORT");
                                            goto Abortado;
                                        }
                                        // Comando SKIP
                                        // Pula para o próximo slide
                                        else if (choice == "k")
                                        {
                                            LogWithColor(ConsoleColor.DarkBlue, "Comando SKIP");
                                            goto ProximoSlide;
                                        }
                                        // Comando TRANSLATE
                                        // Seleciona uma parte da string e joga pra tradução 
                                        // É adicionado como próximo item para ser traduzido
                                        // Não pode ser desfeito/refeito
                                        else if (choice == "t")
                                        {
                                            LogWithColor(ConsoleColor.DarkBlue, "Comando TRANSLATE");

                                            // Pega posição de inicio
                                            string posIni = "a";
                                            int iPosIni;
                                            while (!int.TryParse(posIni, out iPosIni))
                                            {
                                                LogWithColor(ConsoleColor.DarkBlue, "Digite a pos de inicio: ");
                                                posIni = Console.ReadLine();
                                            }

                                            // Pega posição de fim
                                            string posFim = "a";
                                            int iPosFim;
                                            while (!int.TryParse(posFim, out iPosFim))
                                            {
                                                LogWithColor(ConsoleColor.DarkBlue, "Digite a pos de fim (-1 é unbounded): ");
                                                posFim = Console.ReadLine();
                                            }

                                            string selecao;
                                            if (iPosFim == -1) selecao = textos[j].Substring(iPosIni);
                                            else selecao = textos[j].Substring(iPosIni, iPosFim - iPosIni);

                                            LogWithColor(ConsoleColor.DarkBlue, "Mandado pra tradução: " + selecao);
                                            textos.Insert(j + 1, selecao);
                                            LogWithColor(ConsoleColor.DarkBlue, "Será traduzido para: " + TranslateEnToPtBr(selecao));

                                        }
                                        //Comando TRANSLATE2
                                        else if (choice == "T")
                                        {
                                            LogWithColor(ConsoleColor.DarkBlue, "Comando TRANSLATE2");

                                            LogWithColor(ConsoleColor.DarkBlue, "Digite a query: ");
                                            string consulta = Console.ReadLine();
                                            Console.WriteLine(TranslateEnToPtBr(consulta));
                                        }
                                        // Exibe o original
                                        Console.Write("           ");
                                        for (int chCount = 0; chCount < textos[j].Length; chCount += 10)
                                            Console.Write(chCount + new string(' ', 10 - chCount.ToString().Length));
                                        Console.WriteLine();
                                        LogWithColor(ConsoleColor.Green, "[ORIGINAL] " + textos[j]);

                                        // Imprime a nova edição
                                        Console.Write("           ");
                                        for (int chCount = 0; chCount < passosEdicao[passo].Length; chCount += 10)
                                            Console.Write(chCount + new string(' ', 10 - chCount.ToString().Length));
                                        Console.WriteLine();
                                        //      Vermelho é pra "avisar que não era esperado esse texto provavelmente
                                        if (resto == "***") LogWithColor(ConsoleColor.DarkRed, "[TRADUCAO] " + passosEdicao[passo]);
                                        else LogWithColor(ConsoleColor.Magenta, "[TRADUCAO] " + passosEdicao[passo]);

                                        // Pega o próximo comando
                                        choice = Console.ReadLine();
                                    }
                                    
                                    // 0 caracteres é ENTER e apenas ACEITA a tradução
                                    if (choice.Length == 0)
                                    {
                                        
                                        if (passosEdicao[passo] == "")
                                        {

                                        }
                                        else
                                        {
                                            PowerPoint.Shape textBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
                                            15, 50, 300, 500);
                                            ConfigureTextBoxFont(resto, passosEdicao[passo], textBox.TextFrame.TextRange);
                                            /*Rectangle r = ConfigureTextBoxWithinSlide(textBox, slide);
                                            textBox.Left = r.Left;
                                            textBox.Top = r.Top;*/
                                        }
                                    }
                                }

                            }
                        }
                        shape.Top = -50;
                        shape.Left = -50;
                        shape.TextFrame.TextRange.Text = "OK" + shape.TextFrame.TextRange.Text.Substring(3);
                    }

                }
                //Apenas para marcar o label
                ProximoSlide:
                int nada;
            }

            Abortado:

            // Tenta Salvar o arquivo
            // Todos comandos feitos antes de abortar são salvos
            Console.WriteLine("Salvando apresentação...");
            string outputPath = System.IO.Path.Combine(basePath, outFile);

            FileInfo outFi = new FileInfo(outFile);

            while (File.Exists(outputPath) && IsFileLocked(outFi))
            {
                Console.WriteLine("Não é possível salvar no arquivo especificado (padrão -> out.pptx). O arquivo está em uso? Digite um novo output:");
                outFile = Console.ReadLine();
                if (outFile != "") outputPath = System.IO.Path.Combine(basePath, outFile);
                try
                {
                    outFi = new FileInfo(outFile);
                }
                catch
                {
                    outFi = new FileInfo("out.pptx");
                }
            }

            try
            {
                presentation.SaveCopyAs(outputPath, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
                Console.WriteLine("Apresentação salva!");
            }
            catch
            {
                Console.WriteLine("Erro ao salvar :(");
            }

            string input = Console.ReadLine();

        }

        private Rectangle ConfigureTextBoxWithinSlide(PowerPoint.Shape newShape, PowerPoint.Slide slide)
        {
            float WIDTH = newShape.Width;
            float HEIGHT = newShape.Height;
            
            // Chute de melhor valor inicial
            float top = 50;
            float left = 15;

            Rectangle newRec = new Rectangle(left, top, WIDTH, HEIGHT);

            // Para todos outros shapes antigos
            List<Rectangle> rectangles = new List<Rectangle>();

            for (int i = 1; i < slide.Shapes.Count + 1; i++)
            {
                rectangles.Add(new Rectangle(slide.Shapes[i].Left, slide.Shapes[i].Top, slide.Shapes[i].Width, slide.Shapes[i].Height));
            }

            // Procura o melhor valor de top e left
            // Se não achar insere em qualquer lugar do slide
            bool overlaps = true;
            while (overlaps)
            {
                overlaps = false;
                foreach (Rectangle r in rectangles)
                {
                    if (RectangleLeftOverlap(newRec, r) || RectangleTopOverlap(newRec, r))
                    {
                        overlaps = true;
                        break;
                    }
                }

                foreach (Rectangle r in rectangles)
                {
                    while (RectangleLeftOverlap(newRec, r))
                    {
                        newRec.Left++;
                        if (newRec.Left > SlideWidth)
                        {
                            newRec.Top += 15;
                            newRec.Left = 15;
                        }
                    }
                    while (RectangleTopOverlap(newRec, r))
                    {
                        newRec.Top++;
                    }
                }
            }
            //newShape.Left = newRec.Left;
            //newShape.Top = newRec.Top;

            return newRec;

        }

        private bool RectangleLeftOverlap(Rectangle newRec, Rectangle oldRec)
        {
            if ((newRec.Left < oldRec.Left + oldRec.Width) && (newRec.Left > oldRec.Left)) return true;
            if ((newRec.Left + newRec.Width < oldRec.Left + oldRec.Width) && (newRec.Left + newRec.Width > oldRec.Left)) return true;
            return false;
        }

        private bool RectangleTopOverlap(Rectangle newRec, Rectangle oldRec)
        {
            if ((newRec.Top < oldRec.Top + oldRec.Height) && (newRec.Top > oldRec.Top)) return true;
            if ((newRec.Top + newRec.Height < oldRec.Top + oldRec.Height) && (newRec.Top + newRec.Height > oldRec.Top)) return true;
            return false;
        }
        // Configura o TextRange para pré-condições estabelecidas
        private void ConfigureTextBoxFont(string resto, string traducao, PowerPoint.TextRange tr)
        {
            int n = 0;
            if (int.TryParse(resto, out n))
            {
                tr.Text = traducao;
                tr.Font.Bold = MsoTriState.msoTrue;
            }
            else
            {
                if (resto.Length == 1)
                {
                    tr.Text = (char.Parse(resto) - 'a' + 1) + "- " + traducao;
                    for (int i = 1; i <= 4; i++)
                    {
                       tr.Characters(1,1).Font.Bold = MsoTriState.msoTrue;
                    }
                }
                else
                {
                    tr.Text = traducao;
                }
            }

            tr.Font.Size = 15;
            tr.Font.Name = "Calibri";


        }

        // Imprime linha colorida no console
        public static void LogWithColor(ConsoleColor color, string msg, ConsoleColor beforeColor = ConsoleColor.White)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(msg);
            Console.ForegroundColor = beforeColor;
        }

        // Pega uma lista de textos de um pdf
        // Cuidado! Não foi tratado se o path é válido dentro dessa função
        public static List<string> GetTextFromPDF(int pageNumber, string path)
        {
            List<string> contents = new List<string>();

            PdfReader reader = new PdfReader(path);

            if (reader.NumberOfPages > pageNumber)
            {

                ITextExtractionStrategy strategy = new LocationTextExtractionStrategy();
                string currentText = PdfTextExtractor.GetTextFromPage(reader, pageNumber, strategy);

                currentText = Encoding.UTF8.GetString(Encoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.UTF8.GetBytes(currentText)));
                contents.AddRange(currentText.Split(new string[] { ".\n", "__ ", ". __", "â€¢", "EXempty"}, StringSplitOptions.RemoveEmptyEntries));
                
                
                while (contents.Contains("EXempty\n"))
                    contents.Remove("EXempty\n");

                // Agora com os textos devidamente separados
                //
                // Apaga os '\n' internos.
                for (int j = 0; j < contents.Count; j++)
                {
                    int i = 0;
                    for (; i < contents[j].Length; i++)
                    {
                        if (contents[j][i] == '\n')
                        {
                            bool hasEspaceNear = false;
                            if (i != 0 && (contents[j][i - 1] == ' ')) hasEspaceNear = true;
                            if (i != contents[j].Length - 1 && contents[j][i + 1] == ' ') hasEspaceNear = true;
                            if (hasEspaceNear)
                            {
                                contents[j] = contents[j].Substring(0, i) + contents[j].Substring(i + 1); ;
                                i--;
                            }
                            else break; 
                        }
                    }
                    // If an '\n' was found and has no spaces near
                    if (i != contents[j].Length)
                    {
                        contents[j] = contents[j].Substring(0, i) + ' ' + contents[j].Substring(i+1);
                        // Rescan the same content for more '\n'
                        j--;
                    }
                   
                }

                // Usado para debug, é a string original da página inteira
                //contents.Add(currentText);
            }

            return contents;
        }

        // Checa se um arquivo está em uso
        // Visto em: http://stackoverflow.com/questions/876473/is-there-a-way-to-check-if-a-file-is-in-use
        private static bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }

        // Traduz de Inglês pra Pt-Br usando API do Google Translate
        public static string TranslateEnToPtBr(string input)
        {
            //Inspirado em https://ctrlq.org/code/19909-google-translate-api

            // No Google, i won't give my money to you
            string[] inputs = input.Split(new String[] { ". " }, StringSplitOptions.RemoveEmptyEntries);
            List<string> outputs = new List<string>();
            foreach (string inp in inputs) {
                string urlAddress = "https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=pt-BR&dt=t&q="
                        + Uri.EscapeDataString(inp);

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlAddress);
                try
                {
                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                    if (response.StatusCode == HttpStatusCode.OK)
                    {
                        Stream receiveStream = response.GetResponseStream();
                        StreamReader readStream = null;

                        readStream = response.CharacterSet == null ?
                            new StreamReader(receiveStream) :
                            new StreamReader(receiveStream, Encoding.GetEncoding(response.CharacterSet));

                        string data = readStream.ReadToEnd();

                        response.Close();
                        readStream.Close();

                        JArray objects = JArray.Parse(data);
                        outputs.Add(objects.First.First.First.ToString());
                    }
                }
                catch (WebException)
                {
                    // Do nothing cause life is short..
                }
            }
            
            return string.Join(". ", outputs);
        }
    }
}
