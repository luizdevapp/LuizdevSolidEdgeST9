using SolidEdgeCommunity;
using SolidEdgePart;

namespace LuizdevSolidEdgeST9.Models.Part
{
    public class ST9AddPart
    {
        public static void Add(string filePath)
        {
            SolidEdgeFramework.Application application = null!;
            SolidEdgeFramework.Documents documents = null!;
            PartDocument partDocument = null!;

            try
            {
                OleMessageFilter.Register();
                application = SolidEdgeUtils.Connect(true, true);
                documents = application.Documents;
                partDocument = (PartDocument)documents.Add("SolidEdge.PartDocument");
                partDocument.SaveAs(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro: {ex.Message}");
            }
            finally
            {
                // Libera o filtro de mensagens OLE
                OleMessageFilter.Unregister();
            }
        }

        public static void AddToFolder(string file)
        {
            SolidEdgeFramework.Application application = null!;
            SolidEdgeFramework.Documents documents = null!;
            PartDocument partDocument = null!;

            try
            {
                OleMessageFilter.Register();
                application = SolidEdgeUtils.Connect(true, true);
                application.Visible = true;
                documents = application.Documents;
                partDocument = (PartDocument)documents.Add("SolidEdge.PartDocument");
                if (!File.Exists(file))
                {
                    partDocument.SaveAs(file);
                    //partDocument.Close(false);
                    //application.Quit();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro: {ex.Message}");
            }
            finally
            {
                // Libera o filtro de mensagens OLE
                OleMessageFilter.Unregister();
            }
        }
    }
}