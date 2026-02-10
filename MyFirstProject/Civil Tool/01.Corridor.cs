using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Acad = Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

using Civil = Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using Autodesk.Civil.Runtime;
using Autodesk.Civil.Settings;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Civil.ApplicationServices;
using CivSurface = Autodesk.Civil.DatabaseServices.Surface;
using MyFirstProject.Extensions;
// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(MyFirstProject.Corridors))]

namespace MyFirstProject
{
    class Corridors
    {
        [CommandMethod("CTC_AddAllSection")]
        public static void CVC_AddAllSection()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                //start here
                UserInput uI = new();

                // get corridor
                ObjectId corridorId = UserInput.GCorridorId("\n Chọn mô hình corridor để add section:\n");
                Corridor? corridor = tr.GetObject(corridorId, OpenMode.ForWrite) as Corridor;
                int soTT = 0;

#pragma warning disable CS8602 // Dereference of a possibly null reference.
                foreach (Baseline baseline in corridor.Baselines)
                {
                    Alignment? alignment = tr.GetObject(baseline.AlignmentId, OpenMode.ForRead) as Alignment;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    ObjectIdCollection sampleLineGroupIds = alignment.GetSampleLineGroupIds();
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    A.Ok("Số lượng nhóm cọc có trên tuyến là " + sampleLineGroupIds.Count.ToString());


                    // lập danh sách 
                    for (int i = 0; i < sampleLineGroupIds.Count; i++)
                    {
                        ObjectId sampleLineGroupId1 = alignment.GetSampleLineGroupIds()[i];
                        SampleLineGroup? sampleLineGroup1 = tr.GetObject(sampleLineGroupId1, OpenMode.ForRead) as SampleLineGroup;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        A.Ok("Nhóm sampleline thứ " + i.ToString() + " có tên là " + sampleLineGroup1.Name + "/n");
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    }
                    if (sampleLineGroupIds.Count != 1)
                    {
                        soTT = UserInput.GInt("Nhập số thự nhóm cọc muốn add section");
                    }
                    //get station of sampleline
                    int j = 0;
                    ObjectId sampleLineGroupId = alignment.GetSampleLineGroupIds()[soTT];
                    SampleLineGroup? sampleLineGroup = tr.GetObject(sampleLineGroupId, OpenMode.ForRead) as SampleLineGroup;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    foreach (ObjectId samplelineId in sampleLineGroup.GetSampleLineIds())
                    {
                        j++;
                    }
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    double[] station = new double[j];
                    int k = 0;
                    foreach (ObjectId samplelineId in sampleLineGroup.GetSampleLineIds())
                    {
                        SampleLine? sampleLine = tr.GetObject(samplelineId, OpenMode.ForRead) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        station[k] = sampleLine.Station;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                        k++;
                    }

                    //add section to corridor
                    foreach (BaselineRegion baselineRegion in baseline.BaselineRegions)
                    {
                        baselineRegion.ClearAdditionalStations();
                        double[] stationRegular = baselineRegion.SortedStations();
                        foreach (double i in station)
                        {
                            if ((i > baselineRegion.StartStation) && i < baselineRegion.EndStation)
                            {
                                foreach (double stationInRegion in stationRegular)
                                {
                                    if (i == stationInRegion)
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        baselineRegion.AddStation(i, "AddSection");
                                    }
                                }
                            }
                        }
                    }
                }
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                corridor.Rebuild();
                tr.Commit();
            }

            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }
        /*
        [CommandMethod("CVC_CreateCurbReturn_CorridorRegion")]
        public static void CVC_CreateCorridorRegion()
        {
            // start transantion
            using (Transaction tr = a.db.TransactionManager.StartTransaction())
            {
                try
                {
                    UserInput UI = new UserInput();
                    UtilitiesCAD CAD = new UtilitiesCAD();
                    UtilitiesC3D C3D = new UtilitiesC3D();
                    SectionViewGroupCreationPlacementOptions sectionViewGroupCreationPlacementOptions = new();
                    string templatePath = UtilitiesC3D.GetSafeTemplatePath("C:/Windows/LAYOUT CIVIL 3D.dwt");
                    if (string.IsNullOrEmpty(templatePath)) return;
                    sectionViewGroupCreationPlacementOptions.UseProductionPlacement(templatePath, "A3-TN-1-200");

                    //start here
                    //get corridor
                    ObjectId corridorId = UI.G_corridorId("\n Chọn mô hình corridor để tạo mô hình rẽ phải:\n");
                    Corridor corridor = tr.GetObject(corridorId, OpenMode.ForRead) as Corridor;



                    ObjectId alignmentId = UI.G_alignmentId("\n Chọn tim đường " + "cho việc tạo corridor /n");
                    Alignment alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;

                    //get station from alignment
                    double[] station = new double[alignment.Entities.Count];
                    for (int i = 0; i < alignment.Entities.Count; i++)
                    {
                        AlignmentEntity alignmentEntity = alignment.Entities.GetEntityByOrder(i);
                        switch (alignmentEntity.EntityType)
                        {
                            case AlignmentEntityType.Line:
                                {
                                    AlignmentLine alignmentLine = alignmentEntity as AlignmentLine;
                                    station[i] = alignmentLine.Length;
                                    break;
                                }
                            case AlignmentEntityType.Arc:
                                {
                                    AlignmentArc alignmentArc = alignmentEntity as AlignmentArc;
                                    station[i] = alignmentArc.Length;
                                    break;
                                }
                        }
                    }
                    double startStation = station[0];
                    double endStation = station[0] + station[1];

                    // set start and end point for corridor region
                    ObjectId profileId = alignment.GetProfileIds()[0];
                    Profile profile = tr.GetObject(profileId, OpenMode.ForRead) as Profile;

                    //check baseline exist
                    string baselineName = "BL-" + alignment.Name + "-" + profile.Name;
                    foreach (Baseline BL in corridor.Baselines)
                    {
                        if (BL.Name == baselineName)
                        {
                            corridor.Baselines.Remove(corridor.Baselines[baselineName]);
                        }
                    }

                    // then add it again
                    Baseline baselineAdd = corridor.Baselines.Add("BL-" + alignment.Name + "-" + profile.Name, alignmentId, profileId);

                    // create corridor region
                    string regionName = "RG-" + alignment.Name + "-" + startStation.ToString() + "-" + endStation.ToString();
                    ObjectId assemblyId = a.cdoc.AssemblyCollection["Curb Return Fillets"];
                    Assembly assembly = tr.GetObject(assemblyId, OpenMode.ForRead) as Assembly;
                    BaselineRegion baselineRegion = baselineAdd.BaselineRegions.Add(regionName, assemblyId, startStation, endStation);

                    //set frequency for assembly
                    C3D.SetFrequencySection(baselineRegion, 2);

                    // seclect object for target
                    ObjectId TargetId_1 = UI.G_alignmentId(" Chọn tim đường cho target 1");
                    Alignment alignment1 = tr.GetObject(TargetId_1, OpenMode.ForRead) as Alignment;
                    ObjectId profileId_1 = alignment1.GetProfileIds()[0];

                    ObjectId TargetId_2 = UI.G_alignmentId(" Chọn tim đường cho target 2");
                    Alignment alignment2 = tr.GetObject(TargetId_2, OpenMode.ForRead) as Alignment;
                    ObjectId profileId_2 = alignment2.GetProfileIds()[0];

                    ObjectId TargetId_3 = UI.G_Polyline("/n Chọn 1 polyline " + " cho target 3");

                    // set target 0
                    ObjectIdCollection TagetIds_0 = new ObjectIdCollection();
                    TagetIds_0.Add(TargetId_1);
                    TagetIds_0.Add(TargetId_2);

                    // set target 1
                    ObjectIdCollection TagetIds_1 = new ObjectIdCollection();
                    TagetIds_1.Add(profileId_1);
                    TagetIds_1.Add(profileId_2);

                    // set target 1
                    ObjectIdCollection TagetIds_3 = new ObjectIdCollection();
                    TagetIds_3.Add(TargetId_3);


                    //set target for subassemble
                    SubassemblyTargetInfoCollection subassemblyTargetInfoCollection = baselineRegion.GetTargets();

                    //show name subassembly
                    int o = 0;
                    foreach (var item in subassemblyTargetInfoCollection)
                    {
                        a.ed.WriteMessage("\n"+o+item.TargetType.ToString());
                        o++;
                    }
                    /*
                    int j = 1;
                    subassemblyTargetInfoCollection[j].TargetIds = TagetIds_0;
                    int k = 0;
                    subassemblyTargetInfoCollection[k].TargetIds = TagetIds_1;
                    int m = 3;
                    subassemblyTargetInfoCollection[m].TargetIds = TagetIds_3;

                    //set nearest option
                    subassemblyTargetInfoCollection[1].TargetToOption = SubassemblyTargetToOption.Nearest;
                    subassemblyTargetInfoCollection[0].TargetToOption = SubassemblyTargetToOption.Nearest;

                    //set target
                    baselineRegion.SetTargets(subassemblyTargetInfoCollection);
                    

                    tr.Commit();
                }
                catch (Autodesk.AutoCAD.Runtime.Exception e)
                {
                    a.ed.WriteMessage(e.Message);
                }
            }
        }
        */

        [CommandMethod("CTC_TaoCooridor_DuongDoThi_RePhai")]
        public static void CTC_TaoCooridor_DuongDoThi_RePhai()
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();


                //start here
                int alignmentNumber = UserInput.GInt("Khai báo số corridor rẽ phải?");
                //get corridor
                ObjectId corridorId = UserInput.GCorridorId("\n Chọn mô hình corridor để tạo mô hình rẽ phải:\n");
                Corridor? corridor = tr.GetObject(corridorId, OpenMode.ForRead) as Corridor;

                // seclect object for target
                ObjectId TargetId_1 = UserInput.GAlignmentId(" Chọn tim đường cho target 1:");
                Alignment? alignment1 = tr.GetObject(TargetId_1, OpenMode.ForRead) as Alignment;
                ObjectId TargetId_2 = UserInput.GAlignmentId(" Chọn tim đường cho target 2:");
                Alignment? alignment2 = tr.GetObject(TargetId_2, OpenMode.ForRead) as Alignment;

                for (int i = 0; i < alignmentNumber; i++)
                {
                    ObjectId alignmentId = UserInput.GAlignmentId("\n Chọn tim đường " + " cho việc tạo corridor:");
                    Alignment? alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
                    ObjectId polylineId = UserInput.GPolyline("/n Chọn 1 polyline " + " cho target 3:");
                    Polyline? polyline = tr.GetObject(polylineId, OpenMode.ForRead) as Polyline;
#pragma warning disable CS8604 // Possible null reference argument.
#pragma warning disable CS8604 // Possible null reference argument.
#pragma warning disable CS8604 // Possible null reference argument.
#pragma warning disable CS8604 // Possible null reference argument.
#pragma warning disable CS8604 // Possible null reference argument.
                    UtilitiesC3D.TaoCooridorDuongDoThi(alignment, corridor, alignment1, alignment2, polyline);
#pragma warning restore CS8604 // Possible null reference argument.
#pragma warning restore CS8604 // Possible null reference argument.
#pragma warning restore CS8604 // Possible null reference argument.
#pragma warning restore CS8604 // Possible null reference argument.
#pragma warning restore CS8604 // Possible null reference argument.
                }
                ;

                tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                A.Ed.WriteMessage(e.Message);
            }
        }
    }
}

    