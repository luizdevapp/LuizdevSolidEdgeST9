using SolidEdgeCommunity;
using SolidEdgeCommunity.Extensions;
using SolidEdgePart;

namespace LuizdevSolidEdgeST9.Models.SheetMetal
{
    public class ST9AddSheetMetal
    {
        public static void Add(string file)
        {
            SolidEdgeFramework.Application application = null!;
            SolidEdgeFramework.Documents documents = null!;
            SheetMetalDocument sheetMetalDocument = null!;

            try
            {
                OleMessageFilter.Register();
                application = SolidEdgeUtils.Connect(true, true);
                documents = application.Documents;
                sheetMetalDocument = (SheetMetalDocument)documents.Add("SolidEdge.SheetMetalDocument");
                sheetMetalDocument = (SheetMetalDocument)documents.Add(SolidEdgeSDK.PROGID.SolidEdge_SheetMetalDocument);
                sheetMetalDocument = documents.AddSheetMetalDocument();
                sheetMetalDocument = documents.Add<SheetMetalDocument>();
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
            SheetMetalDocument sheetMetalDocument = null!;

            try
            {
                OleMessageFilter.Register();
                application = SolidEdgeUtils.Connect(true, true);
                documents = application.Documents;
                sheetMetalDocument = (SheetMetalDocument)documents.Add("SolidEdge.SheetMetalDocument");
                sheetMetalDocument = (SheetMetalDocument)documents.Add(SolidEdgeSDK.PROGID.SolidEdge_SheetMetalDocument);
                sheetMetalDocument = documents.AddSheetMetalDocument();
                sheetMetalDocument = documents.Add<SheetMetalDocument>();

                if (!File.Exists(file))
                {
                    sheetMetalDocument.SaveAs(file);
                    sheetMetalDocument.Close(false);
                    application.Quit();
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