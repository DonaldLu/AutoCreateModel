using Autodesk.Revit.UI;

namespace AutoCreateModel
{
    public class RevitDocument
    {
        private UIDocument m_revitDoc;
        private Autodesk.Revit.Creation.Application m_appCreator;
        public UIDocument RevitDoc
        {
            get
            {
                return m_revitDoc;
            }
        }
        public RevitDocument(UIApplication app)
        {
            m_revitDoc = app.ActiveUIDocument;
            m_appCreator = app.Application.Create;
        }
    }
}
