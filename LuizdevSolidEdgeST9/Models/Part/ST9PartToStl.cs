namespace LuizdevSolidEdgeST9.Models.Part
{
    public class ST9PartToStl
    {
        public static void Execute(string parPath, string stlPath)
        {
            SolidEdgeFramework.Application app = null!;
            SolidEdgePart.PartDocument part = null!;

            try
            {
                app = (SolidEdgeFramework.Application)
                    ST9Marshal.GetActiveObject("SolidEdge.Application");

                app.DisplayAlerts = false;

                part = (SolidEdgePart.PartDocument)
                    app.Documents.Open(parPath);

                part.SaveAs(stlPath);

                part.Close(false);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}