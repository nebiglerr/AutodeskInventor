using Inventor;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Autodesk_Inventor
{
    internal class Program
    {
        static string GenerateName()
        {
            return System.IO.Path.GetRandomFileName();
        }

        static string SavePath()
        {
            var projectPath = "C:\\Temp\\";
            var partFileName = GenerateName() + ".ipt";
            var fullFilePath = projectPath + partFileName;

            return fullFilePath;
        }

        static void SketchCircles(Application inventorApp, PartDocument partDoc)
        {

            var sketch = partDoc.ComponentDefinition.Sketches.Add(partDoc.ComponentDefinition.WorkPlanes[3]);

            var sketchCircle = sketch.SketchCircles.AddByCenterRadius(inventorApp.TransientGeometry.CreatePoint2d(0, 0), 10.0);

            PartComponentDefinition compDef = partDoc.ComponentDefinition;
            PlanarSketch sketchs = compDef.Sketches.Add(compDef.WorkPlanes[3]);

            SketchCircle circle = sketchs.SketchCircles.AddByCenterRadius(inventorApp.TransientGeometry.CreatePoint2d(0, 0), 5);

            partDoc.SaveAs(SavePath(), false);
        }

        static void SketchRectangle(Application inventorApp, PartDocument partDoc)
        {
            partDoc = (PartDocument)inventorApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, inventorApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject));
            PartComponentDefinition partComponentDefinition = partDoc.ComponentDefinition;
            PlanarSketch planarSketch = partComponentDefinition.Sketches.Add(partComponentDefinition.WorkPlanes[3]);
            TransientGeometry transientGeometry = inventorApp.TransientGeometry;
            SketchEntitiesEnumerator sketchEnttitesEnum = planarSketch.SketchLines.AddAsTwoPointRectangle(transientGeometry.CreatePoint2d(0, 0), transientGeometry.CreatePoint2d(4, 3));
            Profile profile = planarSketch.Profiles.AddForSolid();

            ExtrudeDefinition extrudeDefinition = partComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(profile, PartFeatureOperationEnum.kJoinOperation);
            extrudeDefinition.SetDistanceExtent(4, PartFeatureExtentDirectionEnum.kNegativeExtentDirection);

            partComponentDefinition.Features.ExtrudeFeatures.Add(extrudeDefinition);

            partDoc.SaveAs(SavePath(), false);
        }

        static void CreateExtrusion(Application inventorApp, PartDocument partDoc)
        {
            partDoc = inventorApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject,
                 inventorApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject,
                                                  SystemOfMeasureEnum.kDefaultSystemOfMeasure,
                                                  DraftingStandardEnum.kDefault_DraftingStandard, null),
                                                  true) as PartDocument;

            WorkPlane oWorkPlane = partDoc.ComponentDefinition.WorkPlanes[2];

            PlanarSketch oSketch = partDoc.ComponentDefinition.Sketches.Add(oWorkPlane, false);

            TransientGeometry oTG = inventorApp.TransientGeometry;

            Point2d[] oPoints = new Point2d[5];

            oPoints[0] = oTG.CreatePoint2d(0, 0);
            oPoints[1] = oTG.CreatePoint2d(-10, 0);
            oPoints[2] = oTG.CreatePoint2d(-10, -10);
            oPoints[3] = oTG.CreatePoint2d(5, -10);
            oPoints[4] = oTG.CreatePoint2d(5, -5);

            SketchLine[] oLines = new SketchLine[5];

            oLines[0] = oSketch.SketchLines.AddByTwoPoints(oPoints[0], oPoints[1]);
            oLines[1] = oSketch.SketchLines.AddByTwoPoints(oLines[0].EndSketchPoint, oPoints[2]);
            oLines[2] = oSketch.SketchLines.AddByTwoPoints(oLines[1].EndSketchPoint, oPoints[3]);
            oLines[3] = oSketch.SketchLines.AddByTwoPoints(oLines[2].EndSketchPoint, oPoints[4]);

            oSketch.SketchArcs.AddByCenterStartEndPoint(oTG.CreatePoint2d(0, -5), oLines[3].EndSketchPoint, oLines[0].StartSketchPoint, true);

            Profile oProfile = oSketch.Profiles.AddForSolid(true, null, null);



            PartComponentDefinition oPartDocDef = partDoc.ComponentDefinition;

            ExtrudeFeatures extrudes = oPartDocDef.Features.ExtrudeFeatures;

            ExtrudeDefinition extrudeDef = extrudes.CreateExtrudeDefinition(oProfile, PartFeatureOperationEnum.kNewBodyOperation);

            extrudeDef.SetDistanceExtent(8, PartFeatureExtentDirectionEnum.kPositiveExtentDirection);
            extrudeDef.SetDistanceExtentTwo(20);
            extrudeDef.TaperAngle = "-2 deg";
            extrudeDef.TaperAngleTwo = "-10 deg";

            ExtrudeFeature extrude = extrudes.Add(extrudeDef);

            Camera oCamera = inventorApp.ActiveView.Camera;

            oCamera.ViewOrientationType = ViewOrientationTypeEnum.kIsoTopRightViewOrientation;
            oCamera.Apply();

            inventorApp.ActiveView.Fit(true);
            partDoc.SaveAs(SavePath(), false);
        }

        static void CreateBrep(Application inventorApp, PartDocument partDoc)
        {
            partDoc = inventorApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject,
               inventorApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject,
                                                SystemOfMeasureEnum.kDefaultSystemOfMeasure,
                                                DraftingStandardEnum.kDefault_DraftingStandard, null),
                                                true) as PartDocument;


            WorkPlane oWorkPlanes = partDoc.ComponentDefinition.WorkPlanes[2];

            PlanarSketch oSketchs = partDoc.ComponentDefinition.Sketches.Add(oWorkPlanes, false);

            TransientGeometry oTGS = inventorApp.TransientGeometry;


            Point2d[] oPointss = new Point2d[5];

            oPointss[0] = oTGS.CreatePoint2d(0, 0);
            oPointss[1] = oTGS.CreatePoint2d(-10, 0);
            oPointss[2] = oTGS.CreatePoint2d(-10, -10);
            oPointss[3] = oTGS.CreatePoint2d(5, -10);
            oPointss[4] = oTGS.CreatePoint2d(5, -5);


            SketchLine[] oLiness = new SketchLine[5];

            oLiness[0] = oSketchs.SketchLines.AddByTwoPoints(oPointss[0], oPointss[1]);
            oLiness[1] = oSketchs.SketchLines.AddByTwoPoints(oLiness[0].EndSketchPoint, oPointss[2]);
            oLiness[2] = oSketchs.SketchLines.AddByTwoPoints(oLiness[1].EndSketchPoint, oPointss[3]);
            oLiness[3] = oSketchs.SketchLines.AddByTwoPoints(oLiness[2].EndSketchPoint, oPointss[4]);

            oSketchs.SketchArcs.AddByCenterStartEndPoint(oTGS.CreatePoint2d(0, -5), oLiness[3].EndSketchPoint, oLiness[0].StartSketchPoint, true);


            Profile oProfiles = oSketchs.Profiles.AddForSolid(true, null, null);


            PartComponentDefinition oPartDocDefs = partDoc.ComponentDefinition;


            ExtrudeFeatures extrudess = oPartDocDefs.Features.ExtrudeFeatures;


            ExtrudeDefinition extrudeDefs = extrudess.CreateExtrudeDefinition(oProfiles, PartFeatureOperationEnum.kJoinOperation);


            extrudeDefs.SetDistanceExtent(1, PartFeatureExtentDirectionEnum.kPositiveExtentDirection);

            ExtrudeFeature extruder = extrudess.Add(extrudeDefs);

            partDoc.SaveAs(SavePath(), false);
        }
        static void Main(string[] args)
        {

            Console.WriteLine("Uygulama Başlatılıyor...");
            var inventorApp = Activator.CreateInstance(Type.GetTypeFromProgID("Inventor.Application")) as Inventor.Application;
            inventorApp.Visible = true;

            Console.WriteLine("Model Oluşturuluyor...");
            var partDoc = inventorApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, inventorApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject)) as PartDocument;

            SketchCircles(inventorApp, partDoc);

            Console.WriteLine("Model Kaydedildi.");
            Console.WriteLine("Model Oluşturuluyor...");
            SketchRectangle(inventorApp, partDoc);

            Console.WriteLine("Model Kaydedildi.");
            Console.WriteLine("Model Oluşturuluyor...");

            CreateExtrusion(inventorApp,partDoc);

            Console.WriteLine("Model Kaydedildi.");
            Console.WriteLine("Model Oluşturuluyor...");

            CreateBrep(inventorApp,partDoc);

            Console.WriteLine("Model Kaydedildi.");
            Console.WriteLine("Uygulamayı kapatmak için herhangi bir tuşa basınız");
            Console.ReadKey();
            inventorApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(inventorApp);
            inventorApp = null;

        }

       

    }
}
