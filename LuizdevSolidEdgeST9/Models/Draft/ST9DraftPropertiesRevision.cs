namespace LuizdevSolidEdgeST9.Models.Draft
{
    public class ST9DraftPropertiesRevision
    {
        public static string Number { get; set; } = string.Empty;
        public static string RevisionNumber { get; set; } = string.Empty;
        public static string Modification { get; set; } = string.Empty;
        public static string CreateBy { get; set; } = string.Empty;
        public static string CreationDate { get; set; } = string.Empty;
        public static string ReviewedBy { get; set; } = string.Empty;

        public static void Add(string documentFileName)
        {
            SolidEdgeFramework.Application application = null!;
            SolidEdgeFramework.SolidEdgeDocument document = null!;
            SolidEdgeFramework.PropertySets propertySets = null!;
            SolidEdgeFramework.Properties properties = null!;
            SolidEdgeFramework.Property property = null!;
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
                    property = properties.Add("@" + Number + "RevisionNumber", "" + RevisionNumber + "");
                    property = properties.Add("@" + Number + "Modification", "" + Modification + "");
                    property = properties.Add("@" + Number + "CreateBy", "" + CreateBy + "");
                    property = properties.Add("@" + Number + "CreationDate", "" + CreationDate + "");
                    property = properties.Add("@" + Number + "ReviewedBy", "" + ReviewedBy + "");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}