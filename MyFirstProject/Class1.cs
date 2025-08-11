using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

[assembly: CommandClass(typeof(MyFirstProject.Class1))]

namespace MyFirstProject
{
    public class Class1
    {
        [CommandMethod("ShowForm")]
        public static void ShowForm()
        {
            TestForm frmTest = new();
            frmTest.Show();

        }

            [CommandMethod("AdskGreeting")]
        public static void AdskGreeting()
        {
            // Get the current document and database, and start a transaction
            Document? acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ?? throw new InvalidOperationException("No active document found.");
            Database? acCurDb = acDoc.Database ?? throw new InvalidOperationException("No database found for the active document.");

            // Starts a new transaction with the Transaction Manager
            using Transaction acTrans = acCurDb.TransactionManager.StartTransaction();
            // Open the Block table record for read
            BlockTable acBlkTbl = (BlockTable)acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) ?? throw new InvalidOperationException("BlockTable could not be retrieved.");

            // Open the Block table record Model space for write
            BlockTableRecord acBlkTblRec = (BlockTableRecord)acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) ?? throw new InvalidOperationException("BlockTableRecord could not be retrieved.");

            /* Creates a new MText object and assigns it a location,
            text value and text style */
            using (MText objText = new())
            {
                // Specify the insertion point of the MText object
                objText.Location = new Autodesk.AutoCAD.Geometry.Point3d(2, 2, 0);

                // Set the text string for the MText object
                objText.Contents = "Greetings, Welcome to AutoCAD .NET";

                // Set the text style for the MText object
                objText.TextStyleId = acCurDb.Textstyle;

                // Appends the new MText object to model space
                acBlkTblRec.AppendEntity(objText);

                // Appends to new MText object to the active transaction
                acTrans.AddNewlyCreatedDBObject(objText, true);
            }

            // Saves the changes to the database and closes the transaction
            acTrans.Commit();
        }
    }
}