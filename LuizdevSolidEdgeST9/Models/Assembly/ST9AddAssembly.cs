namespace LuizdevSolidEdgeST9.Models.Assembly
{
    public class ST9AddAssembly
    {
        public static void Add(string filePath)
        {
            SolidEdgeFramework.Application application;
            SolidEdgeFramework.Documents documents;
            SolidEdgeAssembly.AssemblyDocument assemblyDocument;

            try
            {
#pragma warning disable CA1416 // Validate platform compatibility
                SolidEdgeCommunity.OleMessageFilter.Register();
#pragma warning restore CA1416 // Validate platform compatibility
                application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true);
                documents = application.Documents;
                assemblyDocument = (SolidEdgeAssembly.AssemblyDocument)documents.Add("SolidEdge.AssemblyDocument");
                assemblyDocument.SaveAs($"{filePath}");
            }
            catch (Exception ex)
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