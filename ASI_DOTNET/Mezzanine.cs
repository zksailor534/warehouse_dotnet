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
        public static void CreateColumn(Database acCurDb,
            double cHeight,
            double cDiameter,
            double bDiameter)
        {
            // Frame block name
            string cName = "Column - " + cHeight + " - " + cDiameter + "x" + cDiameter;

            // Baseplate default height (could be argument)
            double bHeight = 0.75;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table record for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                             OpenMode.ForRead) as BlockTable;

                // Blank objectid reference
                ObjectId fId = ObjectId.Null;

                if (acBlkTbl.Has(cName))
                {
                    // Retrieve object id
                    fId = acBlkTbl[cName];
                }
                else
                {
                    using (BlockTableRecord acBlkTblRec = new BlockTableRecord())
                    {
                        acBlkTblRec.Name = cName;

                        // Set the insertion point for the block
                        acBlkTblRec.Origin = new Point3d(0, 0, 0);
                        
                        // Calculate baseplate locations
                        Point3d bMid = new Point3d(bDiameter / 2, bDiameter / 2, bHeight / 2);
                        Vector3d bVec = Point3d.Origin.GetVectorTo(bMid);
                        var cBottom = cHeight / 2 + bHeight;
                        Vector3d cVec;
                        Point3d cMid;

                        //// Calculate column location relative to baseplate
                        //if (cPlaceRes.StringResult == "Corner")
                        //{
                        //    cMid = new Point3d(cWidth / 2, cWidth / 2, cBottom);
                        //}
                        //else if (cPlaceRes.StringResult == "Side")
                        //{
                        //    cMid = new Point3d(bWidth / 2, cWidth / 2, cBottom);
                        //}
                        //else
                        //{
                            cMid = new Point3d(bDiameter / 2, bDiameter / 2, cBottom);
                        //}

                        cVec = Point3d.Origin.GetVectorTo(cMid);
                
                        // Create the 3D solid baseplate
                        Solid3d bPlate = new Solid3d();
                        bPlate.SetDatabaseDefaults();
                        bPlate.CreateBox(bDiameter, bDiameter, bHeight);

                        // Position the baseplate 
                        bPlate.TransformBy(Matrix3d.Displacement(bVec));

                        // Add baseplate to block
                        acBlkTblRec.AppendEntity(bPlate);

                        // Create the 3D solid column
                        Solid3d column = new Solid3d();
                        column.SetDatabaseDefaults();
                        column.CreateBox(cDiameter, cDiameter, cHeight - bHeight);

                        // Position the column 
                        column.TransformBy(Matrix3d.Displacement(cVec));

                        // Add column to block
                        acBlkTblRec.AppendEntity(column);

                        /// Add Block to Block Table and close Transaction
                        acBlkTbl.UpgradeOpen();
                        acBlkTbl.Add(acBlkTblRec);
                        acTrans.AddNewlyCreatedDBObject(acBlkTblRec, true);
                    }

                }

                // Save the new object to the database
                acTrans.Commit();

            }

        }


        [CommandMethod("ColumnPrompt")]
        public static void ColumnPrompt()
        {
            // Frame default characteristics
            double cHeight = 96;
            double cDiameter = 6;
            double bDiameter = 16;

            // Get the current document and database, and start a transaction
            // !!! Move this outside function
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Prompt for the column height
            PromptDoubleResult cHeightRes;
            PromptDistanceOptions cHeightOpts = new PromptDistanceOptions("");
            cHeightOpts.Message = "\nEnter the total column height: ";
            cHeightRes = acDoc.Editor.GetDistance(cHeightOpts);
            cHeight = cHeightRes.Value;

            // Exit if the user presses ESC or cancels the command
            if (cHeightRes.Status == PromptStatus.Cancel) return;

            // Prompt for the column width
            PromptDoubleResult cDiameterRes;
            PromptDistanceOptions cDiameterOpts = new PromptDistanceOptions("");
            cDiameterOpts.Message = "\nEnter the column diameter: ";
            cDiameterRes = acDoc.Editor.GetDistance(cDiameterOpts);
            cDiameter = cDiameterRes.Value;

            // Exit if the user presses ESC or cancels the command
            if (cDiameterRes.Status == PromptStatus.Cancel) return;

            // Prompt for the base width
            PromptDoubleResult bDiameterRes;
            PromptDistanceOptions bDiameterOpts = new PromptDistanceOptions("");
            bDiameterOpts.Message = "\nEnter the baseplate diameter: ";
            bDiameterRes = acDoc.Editor.GetDistance(bDiameterOpts);
            bDiameter = bDiameterRes.Value;

            // Exit if the user presses ESC or cancels the command
            if (bDiameterRes.Status == PromptStatus.Cancel) return;

            //// Prompt for the base height
            //PromptDoubleResult bHeightRes;
            //PromptDistanceOptions bHeightOpts = new PromptDistanceOptions("");
            //bHeightOpts.Message = "\nEnter the baseplate height: ";
            //bHeightOpts.DefaultValue = 0.75;
            //bHeightRes = acDoc.Editor.GetDistance(bHeightOpts);
            //double bHeight = bHeightRes.Value;

            //// Exit if the user presses ESC or cancels the command
            //if (bHeightRes.Status == PromptStatus.Cancel) return;

            //// Prompt for the column orientation on the baseplate
            //PromptKeywordOptions cPlaceOpts = new PromptKeywordOptions("");
            //cPlaceOpts.Message = "\nEnter baseplate offset: ";
            //cPlaceOpts.Keywords.Add("Center");
            //cPlaceOpts.Keywords.Add("Side");
            //cPlaceOpts.Keywords.Add("Corner");
            //cPlaceOpts.Keywords.Default = "Center";
            //cPlaceOpts.AllowArbitraryInput = false;
            //PromptResult cPlaceRes = acDoc.Editor.GetKeywords(cPlaceOpts);

            //// Exit if the user presses ESC or cancels the command
            //if (cPlaceRes.Status == PromptStatus.Cancel) return;

            Column.CreateColumn(acCurDb, cHeight, cDiameter, bDiameter);

            //// Open the active viewport
            //ViewportTableRecord acVportTblRec;
            //acVportTblRec = acTrans.GetObject(acDoc.Editor.ActiveViewportId,
            //                                  OpenMode.ForWrite) as ViewportTableRecord;

            //// Save the new objects to the database
            //acTrans.Commit();

        }

    }
    
}
