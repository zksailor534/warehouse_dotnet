using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Windows;

[assembly: CommandClass(typeof(ASI_DOTNET.Column))]

namespace ASI_DOTNET
{
    class Column
    {

        [CommandMethod("CreateColumn")]
        public static void CreateColumn()
        {
            // Get the current document and database, and start a transaction
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table record for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                             OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;

                // Prompt for the column height
                PromptDoubleResult cHeightRes;
                PromptDistanceOptions cHeightOpts = new PromptDistanceOptions("");
                cHeightOpts.Message = "\nEnter the total column height: ";
                cHeightRes = acDoc.Editor.GetDistance(cHeightOpts);
                double cHeight = cHeightRes.Value;

                // Exit if the user presses ESC or cancels the command
                if (cHeightRes.Status == PromptStatus.Cancel) return;

                // Prompt for the column width
                PromptDoubleResult cWidthRes;
                PromptDistanceOptions cWidthOpts = new PromptDistanceOptions("");
                cWidthOpts.Message = "\nEnter the column width: ";
                cWidthRes = acDoc.Editor.GetDistance(cWidthOpts);
                double cWidth = cWidthRes.Value;

                // Exit if the user presses ESC or cancels the command
                if (cWidthRes.Status == PromptStatus.Cancel) return;

                // Prompt for the base width
                PromptDoubleResult bWidthRes;
                PromptDistanceOptions bWidthOpts = new PromptDistanceOptions("");
                bWidthOpts.Message = "\nEnter the baseplate width: ";
                bWidthRes = acDoc.Editor.GetDistance(bWidthOpts);
                double bWidth = bWidthRes.Value;

                // Exit if the user presses ESC or cancels the command
                if (bWidthRes.Status == PromptStatus.Cancel) return;

                // Prompt for the base height
                PromptDoubleResult bHeightRes;
                PromptDistanceOptions bHeightOpts = new PromptDistanceOptions("");
                bHeightOpts.Message = "\nEnter the baseplate height: ";
                bHeightOpts.DefaultValue = 0.75;
                bHeightRes = acDoc.Editor.GetDistance(bHeightOpts);
                double bHeight = bHeightRes.Value;

                // Exit if the user presses ESC or cancels the command
                if (bHeightRes.Status == PromptStatus.Cancel) return;

                // Prompt for the column orientation on the baseplate
                PromptKeywordOptions cPlaceOpts = new PromptKeywordOptions("");
                cPlaceOpts.Message = "\nEnter baseplate offset: ";
                cPlaceOpts.Keywords.Add("Center");
                cPlaceOpts.Keywords.Add("Side");
                cPlaceOpts.Keywords.Add("Corner");
                cPlaceOpts.Keywords.Default = "Center";
                cPlaceOpts.AllowArbitraryInput = false;
                PromptResult cPlaceRes = acDoc.Editor.GetKeywords(cPlaceOpts);

                // Exit if the user presses ESC or cancels the command
                if (cPlaceRes.Status == PromptStatus.Cancel) return;

                // Calculate baseplate locations
                Point3d bMid = new Point3d(bWidth / 2, bWidth / 2, bHeight / 2);
                Vector3d bVec = Point3d.Origin.GetVectorTo(bMid);
                var cBottom = cHeight / 2 + bHeight;
                Vector3d cVec;
                Point3d cMid;

                // Calculate column location relative to baseplate
                if (cPlaceRes.StringResult == "Corner")
                {
                    cMid = new Point3d(cWidth / 2, cWidth / 2, cBottom);
                }
                else if (cPlaceRes.StringResult == "Side")
                {
                    cMid = new Point3d(bWidth / 2, cWidth / 2, cBottom);
                }
                else
                {
                    cMid = new Point3d(bWidth / 2, bWidth / 2, cBottom);
                }

                cVec = Point3d.Origin.GetVectorTo(cMid);
                
                // Create the 3D solid baseplate
                Solid3d b1 = new Solid3d();
                b1.SetDatabaseDefaults();
                b1.CreateBox(bWidth, bWidth, bHeight);

                // Position the baseplate 
                b1.TransformBy(Matrix3d.Displacement(bVec));

                // Create the 3D solid column
                Solid3d c1 = new Solid3d();
                c1.SetDatabaseDefaults();
                c1.CreateBox(cWidth, cWidth, cHeight);

                // Position the column 
                c1.TransformBy(Matrix3d.Displacement(cVec));

                // Add the new objects to the block table record and the transaction
                acBlkTblRec.AppendEntity(b1);
                acTrans.AddNewlyCreatedDBObject(b1, true);
                acBlkTblRec.AppendEntity(c1);
                acTrans.AddNewlyCreatedDBObject(c1, true);

                // Open the active viewport
                ViewportTableRecord acVportTblRec;
                acVportTblRec = acTrans.GetObject(acDoc.Editor.ActiveViewportId,
                                                  OpenMode.ForWrite) as ViewportTableRecord;

                // Save the new objects to the database
                acTrans.Commit();
            }
        }
    }
    
}
