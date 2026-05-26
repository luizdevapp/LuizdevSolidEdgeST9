using SolidEdgeCommunity.Extensions;

namespace LuizdevSolidEdgeST9.Models.Draft
{
    public class ST9AddDraft
    {
        public static void Add(string templateDraft, string partLink, string filePath)
        {
            SolidEdgeFramework.Application application = null!;
            SolidEdgeFramework.Documents documents = null!;
            SolidEdgeDraft.DraftDocument draftDocument = null!;
            SolidEdgeDraft.ModelLinks modelLinks = null!;
            SolidEdgeDraft.ModelLink modelLink = null!;
            SolidEdgeDraft.Sheet sheet = null!;
            SolidEdgeDraft.DrawingViews drawingViews = null!;
            SolidEdgeDraft.DrawingView drawingView = null!;
            string filename = null!;

            try
            {
                SolidEdgeCommunity.OleMessageFilter.Register();
                application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true);
                documents = application.Documents;
                draftDocument = documents.AddDraftDocument(templateDraft);
                modelLinks = draftDocument.ModelLinks;
                filename = Path.Combine(filePath);
                modelLink = modelLinks.Add(partLink);
                sheet = draftDocument.ActiveSheet;
                drawingViews = sheet.DrawingViews;
                drawingView = drawingViews.AddPartView(
                    From: modelLink,
                     Orientation: SolidEdgeDraft.ViewOrientationConstants.igFrontView,
                    Scale: 1,
                    x: 0.1,
                    y: 0.15,
                    ViewType: SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView);
                draftDocument.SaveAs(filePath);
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

        public static void AddToFolder(string templateDraft, string partLink, string filePath)
        {
            SolidEdgeFramework.Application application = null!;
            SolidEdgeFramework.Documents documents = null!;
            SolidEdgeDraft.DraftDocument draftDocument = null!;
            SolidEdgeDraft.ModelLinks modelLinks = null!;
            SolidEdgeDraft.ModelLink modelLink = null!;
            SolidEdgeDraft.Sheet sheet = null!;
            SolidEdgeDraft.DrawingViews drawingViews = null!;
            SolidEdgeDraft.DrawingView drawingView = null!;
            string filename = null!;

            try
            {
                SolidEdgeCommunity.OleMessageFilter.Register();
                application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true);
                application.Visible = true;
                documents = application.Documents;
                draftDocument = documents.AddDraftDocument(templateDraft);
                modelLinks = draftDocument.ModelLinks;
                filename = Path.Combine(filePath);
                modelLink = modelLinks.Add(partLink);
                sheet = draftDocument.ActiveSheet;
                drawingViews = sheet.DrawingViews;
                drawingView = drawingViews.AddPartView(
                    From: modelLink,
                     Orientation: SolidEdgeDraft.ViewOrientationConstants.igFrontView,
                    Scale: 1,
                    x: 0.1,
                    y: 0.15,
                    ViewType: SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView);
                if (!File.Exists(filePath))
                {
                    draftDocument.SaveAs(filePath);
                    //partDocument.Close(false);
                    //application.Quit();
                }
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