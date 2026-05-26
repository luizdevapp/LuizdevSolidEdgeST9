namespace LuizdevSolidEdgeST9.Models.Draft
{
    public class ST9DraftPropertiesProduction3
    {
        public static string NumDrawing = string.Empty;
        public static string NumTool = string.Empty;
        public static string ProductName = string.Empty;
        public static string ProductCode = string.Empty;
        public static string MaterialName = string.Empty;
        public static string CreateBy = string.Empty;
        public static string ReviewedBy = string.Empty;
        public static string CreationDate = string.Empty;

        public ST9DraftPropertiesProduction3()
        {
        }

        public static void Add(string documentFileName)
        {
            SolidEdgeFramework.Application application = null!;
            SolidEdgeFramework.PropertySets propertySets = null!;
            SolidEdgeFramework.Properties properties = null!;
            SolidEdgeFramework.Property property = null!;
            SolidEdgeFramework.SolidEdgeDocument document = null!;
            try
            {
                application = (SolidEdgeFramework.Application)
                ST9Marshal.GetActiveObject("SolidEdge.Application");
                document = (SolidEdgeFramework.SolidEdgeDocument)
                application.ActiveDocument;
                propertySets = (SolidEdgeFramework.PropertySets)document.Properties;

                string documentName = document.Name;
                string onlyFileName = Path.GetFileName(documentFileName);
                if (string.Equals(documentName, onlyFileName, StringComparison.OrdinalIgnoreCase))
                {
                    properties = propertySets.Item("Custom");
                    property = properties.Add("@NumDrawing", "" + NumDrawing + "");
                    property = properties.Add("@NumTool", "" + NumTool + "");
                    property = properties.Add("@ProductName", "" + ProductName + "");
                    property = properties.Add("@ProductCode", "" + ProductCode + "");
                    property = properties.Add("@MaterialName", "" + MaterialName + "");
                    property = properties.Add("@CreateBy", "" + CreateBy + "");
                    property = properties.Add("@ReviewedBy", "" + ReviewedBy + "");
                    property = properties.Add("@CreationDate", "" + CreationDate + "");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}