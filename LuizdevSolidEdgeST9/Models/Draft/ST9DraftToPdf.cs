namespace LuizdevSolidEdgeST9.Models.Draft
{
    public class ST9DraftToPdf
    {
        public static void ConvertDraftToPdf(string draftPath, string pathSavePdf)
        {
            SolidEdgeFramework.Application application = null!;
            DraftDocument draftDoc = null!;

            try
            {
                application = (SolidEdgeFramework.Application)ST9Marshal.GetActiveObject("SolidEdge.Application");
            }
            catch (COMException)
            {
                var type = Type.GetTypeFromProgID("SolidEdge.Application")!;
                application = (SolidEdgeFramework.Application)Activator.CreateInstance(type!)!;
                application.Visible = true;
            }

            draftDoc = (SolidEdgeDraft.DraftDocument)application.Documents.Open(draftPath);
            draftDoc.SaveAs(pathSavePdf);
            draftDoc.Close(true);
        }
    }
}