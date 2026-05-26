namespace LuizdevSolidEdgeST9.Models.Assembly
{
    public class ST9OccurrenceToActiveAssembly
    {
        public static void Add(string pathPart)
        {
            SolidEdgeFramework.Application? application = null;
            SolidEdgeAssembly.AssemblyDocument? assemblyDocument = null;
            SolidEdgeAssembly.Occurrences? occurrences = null;
            SolidEdgeAssembly.Occurrence? occurrence = null;

            try
            {
                SolidEdgeCommunity.OleMessageFilter.Register();

                // Conecta ao Solid Edge que já deve estar aberto
                application = SolidEdgeCommunity.SolidEdgeUtils.Connect(false);

                // Em vez de AddAssemblyDocument(), pegamos o ActiveDocument
                var activeDocument = application.ActiveDocument;

                if (activeDocument == null)
                {
                    throw new Exception("Não há nenhum documento aberto no Solid Edge.");
                }

                // Verifica se o documento ativo é um Assembly (.asm)
                assemblyDocument = activeDocument as SolidEdgeAssembly.AssemblyDocument;

                if (assemblyDocument != null)
                {
                    // Obtém a coleção de ocorrências do documento ativo
                    occurrences = assemblyDocument.Occurrences;

                    // Caminho do arquivo a ser inserido
                    string filename = System.IO.Path.Combine(SolidEdgeCommunity.SolidEdgeUtils.GetTrainingFolderPath(), pathPart);

                    // Adiciona a peça no assembly atual
                    occurrence = occurrences.AddByFilename(filename);

                    Console.WriteLine("Peça adicionada com sucesso ao assembly ativo!");
                }
                else
                {
                    Console.WriteLine("O documento ativo não é um Assembly (.asm).");
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                SolidEdgeCommunity.OleMessageFilter.Unregister();
            }
        }
    }
}