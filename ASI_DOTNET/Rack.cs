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

[assembly: CommandClass(typeof(ASI_DOTNET.Rack))]

namespace ASI_DOTNET
{
    class Rack
    {

        public static void CreateFrame(Database acCurDb,
            double fHeight,
            double fWidth,
            double fDiameter)
        {
            // Frame block name
            string fName = "Rack Frame - " + fHeight + "x" + fWidth;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                // Blank objectid reference
                ObjectId fId = ObjectId.Null;

                if (acBlkTbl.Has(fName))
                {
                    // Retrieve object id
                    fId = acBlkTbl[fName];
                }
                else
                {
                    using (BlockTableRecord acBlkTblRec = new BlockTableRecord())
                    {
                        acBlkTblRec.Name = fName;

                        // Set the insertion point for the block
                        acBlkTblRec.Origin = new Point3d(0, 0, 0);

                        // Add frame posts to the block table record
                        for ( double dist = fDiameter; dist <= fWidth; dist += fWidth - fDiameter) {

                            Solid3d Post = new Solid3d();
                            Post.SetDatabaseDefaults();

                            // Create the 3D solid frame post
                            Post.CreateBox(fDiameter, fDiameter, fHeight);

                            // Find location for post
                            Point3d pMid = new Point3d(fDiameter / 2, dist - fDiameter / 2, fHeight / 2);
                            Vector3d pVec = Point3d.Origin.GetVectorTo(pMid);

                            // Position the post
                            Post.TransformBy(Matrix3d.Displacement(pVec));

                            acBlkTblRec.AppendEntity(Post);

                        }

                        // Calculate cross-bracing
                        int bNum;
                        double bSpace = 0;
                        double bAngle, bLeftOver;
                        double bSize = 1;
                        for (bNum = Convert.ToInt32(fHeight / 12); bNum >= 1; bNum--)
                        {
                            bSpace = Math.Floor((fHeight - (bNum + 1) * bSize) / bNum);
                            bAngle = Math.Atan2(bSpace - bSize, fWidth - (2 * fDiameter));
                            if ((bAngle < Math.PI / 3) && (bSpace >= fWidth))
                            {
                                break;
                            }
                        }
                        bLeftOver = fHeight - (bNum + 1) * bSize - bNum * bSpace;

                        /// Add horizontal cross-braces to the block
                        for (int i = 1; i <= bNum + 1; i++)
                        {
                            Solid3d hBrace = new Solid3d();
                            hBrace.SetDatabaseDefaults();

                            // Create the 3D solid horizontal cross brace
                            hBrace.CreateBox(bSize, fWidth - 2 * fDiameter, bSize);

                            // Find location for brace
                            Point3d bMid = new Point3d(fDiameter / 2, fWidth / 2,
                                bLeftOver / 2 + ((bSpace + bSize) * (i - 1)) + bSize / 2);
                            Vector3d bVec = Point3d.Origin.GetVectorTo(bMid);

                            // Position the brace
                            hBrace.TransformBy(Matrix3d.Displacement(bVec));

                            acBlkTblRec.AppendEntity(hBrace);
                        }

                        /// Add angled cross-braces to the block
                        /// Create polyline of brace profile and rotate to correct orienation
                        Polyline bPoly = new Polyline();
                        bPoly.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);
                        bPoly.AddVertexAt(1, new Point2d(bSize, 0), 0, 0, 0);
                        bPoly.AddVertexAt(2, new Point2d(bSize, bSize), 0, 0, 0);
                        bPoly.AddVertexAt(3, new Point2d(0, bSize), 0, 0, 0);
                        bPoly.Closed = true;
                        bPoly.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.XAxis, Point3d.Origin));

                        /// Create sweep path
                        Line bPath = new Line(Point3d.Origin, new Point3d(0, fWidth - 2 * fDiameter, bSpace - bSize));

                        /// Create swept cross braces
                        for (int i = 1; i <= bNum; i++)
                        {
                            // Create swept brace at origin
                            Solid3d aBrace = Utils.SweepPolylineOverLine(bPoly, bPath);

                            // Calculate location
                            double aBxLoc = (fDiameter / 2) - (bSize / 2);
                            double aByLoc = fDiameter;
                            double aBzLoc = bLeftOver / 2 + bSize * i + bSpace * (i - 1);

                            /// Even braces get rotated 180 deg and different x and y coordinates
                            if (i % 2 == 0)
                            {
                                aBrace.TransformBy(Matrix3d.Rotation(Math.PI, Vector3d.ZAxis, Point3d.Origin));
                                aBxLoc = (fDiameter / 2) + (bSize / 2);
                                aByLoc = fWidth - fDiameter;
                            }

                            // Create location for brace
                            Point3d aMid = new Point3d(aBxLoc, aByLoc, aBzLoc);
                            Vector3d aVec = Point3d.Origin.GetVectorTo(aMid);

                            // Position the brace
                            aBrace.TransformBy(Matrix3d.Displacement(aVec));

                            // Add brace to block
                            acBlkTblRec.AppendEntity(aBrace);
                        }

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


        [CommandMethod("RackPrompt")]
        public static void RackPrompt()
        {
            // Get the current document and database, and start a transaction
            // !!! Move this outside function
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Prompt for the frame height
            PromptDoubleResult heightRes;
            PromptDistanceOptions heightOpts = new PromptDistanceOptions("");
            heightOpts.Message = "\nEnter the frame height: ";
            heightOpts.DefaultValue = 96;
            heightRes = acDoc.Editor.GetDistance(heightOpts);
            double height = heightRes.Value;

            // Exit if the user presses ESC or cancels the command
            if (heightRes.Status != PromptStatus.OK) return;

            // Prompt for the frame width
            PromptDoubleResult widthRes;
            PromptDistanceOptions widthOpts = new PromptDistanceOptions("");
            widthOpts.Message = "\nEnter the frame width: ";
            widthOpts.DefaultValue = 36;
            widthRes = acDoc.Editor.GetDistance(widthOpts);
            double width = widthRes.Value;

            // Exit if the user presses ESC or cancels the command
            if (widthRes.Status != PromptStatus.OK) return;

            // Prompt for the frame diameter
            PromptDoubleResult diameterRes;
            PromptDistanceOptions diameterOpts = new PromptDistanceOptions("");
            diameterOpts.Message = "\nEnter the frame width: ";
            diameterOpts.DefaultValue = 3.0;
            diameterRes = acDoc.Editor.GetDistance(diameterOpts);
            double diameter = diameterRes.Value;

            // Exit if the user presses ESC or cancels the command
            if (diameterRes.Status != PromptStatus.OK) return;

            // Create Frame block
            Rack.CreateFrame(acCurDb, height, width, diameter);
        }

    }

}
