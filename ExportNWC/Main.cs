using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportNWC
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            View view = doc.ActiveView;

            using (var ts = new Transaction(doc, "export NWC"))
            {
                ts.Start();
                var nwcOption = new NavisworksExportOptions
                {
                    ExportElementIds =true,                    
                    Parameters=NavisworksParameters.All,
                    ExportLinks=false,
                    ExportRoomAsAttribute=true,
                    Coordinates= NavisworksCoordinates.Shared,
                    ExportScope=NavisworksExportScope.Model,
                    ExportUrls=true,
                    ExportRoomGeometry=true,
                    ConvertElementProperties=false,
                    ExportParts=false,
                    FindMissingMaterials= true,
                    FacetingFactor=1,
                    DivideFileIntoLevels=true,
                    ConvertLights=false,
                    ConvertLinkedCADFormats=true                    
                };
                doc.Export(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "export.nwc", nwcOption);
                ts.Commit();
            }
            return Result.Succeeded;
        }
    }
}
