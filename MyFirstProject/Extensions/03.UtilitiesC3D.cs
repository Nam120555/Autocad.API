using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using CivSurface = Autodesk.Civil.DatabaseServices.Surface;
using Autodesk.Civil.Runtime;
using Autodesk.Civil.Settings;
// (C) Copyright 2016 by  
//
using System;
using System.Windows.Markup;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace MyFirstProject.Extensions
{
    public class UtilitiesC3D
    {
        public static void CreateCogoPointFromPoint3D(Point3d point3D, string derciption)
        {
            // create cogo point from location
            CogoPointCollection cogoPointColl = A.Cdoc.CogoPoints;
            _ = cogoPointColl.Add(point3D, derciption, true);

        }
        public static void GStationsFromSamplelineGroup(out double[] station, out SampleLine sampleLineOut)
        {
            Transaction tr = A.Db.TransactionManager.StartTransaction();
            {
                _ = new UserInput();

                //select sampleline
                ObjectId samplelineId = UserInput.GSampleLineId("\n Chọn 1 nhóm cọc: \n");
                SampleLine? sampleLine = tr.GetObject(samplelineId, OpenMode.ForRead) as SampleLine;
#pragma warning disable CS8601 // Possible null reference assignment.
                sampleLineOut = sampleLine;
#pragma warning restore CS8601 // Possible null reference assignment.
                              //get samplelineGroup
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                ObjectId sampleLineGroupId = sampleLine.GroupId;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                SampleLineGroup? sampleLineGroup = tr.GetObject(sampleLineGroupId, OpenMode.ForRead) as SampleLineGroup;

                //get number of station in samplelineGroup
                int i = 0;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                foreach (ObjectId sampleLineId in sampleLineGroup.GetSampleLineIds())
                {
                    i++;
                }
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                // get station
                station = new double[i];
                int j = 0;
                foreach (ObjectId sampleLineId in sampleLineGroup.GetSampleLineIds())
                {
                    SampleLine? sampleline = tr.GetObject(sampleLineId, OpenMode.ForRead) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    station[j] = sampleline.Station;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    j++;
                }
                tr.Commit();
            }
        }
        public static void SetFrequencySection(BaselineRegion baselineRegion, int frequencyDistance)
        {
            baselineRegion.AppliedAssemblySetting.FrequencyAlongCurves = frequencyDistance;
            baselineRegion.AppliedAssemblySetting.FrequencyAlongProfileCurves = frequencyDistance;
            baselineRegion.AppliedAssemblySetting.FrequencyAlongSpirals = frequencyDistance;
            baselineRegion.AppliedAssemblySetting.FrequencyAlongTangents = frequencyDistance;
            baselineRegion.AppliedAssemblySetting.FrequencyAlongTargetCurves = frequencyDistance;
            baselineRegion.AppliedAssemblySetting.AppliedAdjacentToOffsetTargetStartEnd = true;
            baselineRegion.AppliedAssemblySetting.AppliedAtHorizontalGeometryPoints = true;
            baselineRegion.AppliedAssemblySetting.AppliedAtOffsetTargetGeometryPoints = true;
            baselineRegion.AppliedAssemblySetting.AppliedAtProfileGeometryPoints = true;
            baselineRegion.AppliedAssemblySetting.AppliedAtProfileHighLowPoints = true;
            baselineRegion.AppliedAssemblySetting.AppliedAtSuperelevationCriticalPoints = true;
        }
        public static void MergelectionSet()
        {
            //UserInput UI = new UserInput();
            // Request for objects to be selected in the drawing area
            PromptSelectionResult acSSPrompt;
            acSSPrompt = A.Ed.GetSelection();

            SelectionSet acSSet1;
            ObjectIdCollection acObjIdColl = [];

            // If the prompt status is OK, objects were selected
            if (acSSPrompt.Status == PromptStatus.OK)
            {
                // Get the selected objects
                acSSet1 = acSSPrompt.Value;

                // Append the selected objects to the ObjectIdCollection
                acObjIdColl = [.. acSSet1.GetObjectIds()];
            }

            // Request for objects to be selected in the drawing area
            acSSPrompt = A.Ed.GetSelection();

            SelectionSet acSSet2;

            // If the prompt status is OK, objects were selected
            if (acSSPrompt.Status == PromptStatus.OK)
            {
                acSSet2 = acSSPrompt.Value;

                // Check the size of the ObjectIdCollection, if zero, then initialize it
                if (acObjIdColl.Count == 0)
                {
                    acObjIdColl = [.. acSSet2.GetObjectIds()];
                }
                else
                {
                    // Step through the second selection set
                    foreach (ObjectId acObjId in acSSet2.GetObjectIds())
                    {
                        // Add each object id to the ObjectIdCollection
                        acObjIdColl.Add(acObjId);
                    }
                }
            }

            Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowAlertDialog("Number of objects selected: " + acObjIdColl.Count.ToString());
        }
        public static void CheckUDPforCogoPoint(string ClassificationName)
        {
            // Check if classification already exists. If not, create it. 
            UDPClassification? sampleClassification;
            if (A.Cdoc.PointUDPClassifications.Contains(ClassificationName))
            {
                sampleClassification = A.Cdoc.PointUDPClassifications[ClassificationName];
                // sampleClassification = _civildoc.PointUDPClassifications["Inexistent"];  // Throws exception.
            }
            else
            {
                sampleClassification = A.Cdoc.PointUDPClassifications.Add(ClassificationName);
                // sampleClassification = _civildoc.PointUDPClassifications.Add("Existent"); // Throws exception.
            }

            // Create new UDP// 
            AttributeTypeInfoInt typeInfoInt = new("Sample_Int_UDP")
            {
                Description = "Sample integer User Defined Property",
                DefaultValue = 15,
                LowerBoundValue = 0,
                UpperBoundValue = 100,
                LowerBoundInclusive = true,
                UpperBoundInclusive = false
            };
            _ = sampleClassification.CreateUDP(typeInfoInt);
        }

        public static void TaoCooridorDuongDoThi(Alignment alignment, Corridor corridor, Alignment alignment1, Alignment alignment2, Polyline polyline)
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();

                //start here


                //ObjectId alignmentId = UI.G_alignmentId("for create corridor");
                //Alignment alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;

                //get station from alignment
                double[] station = new double[alignment.Entities.Count];
                for (int i = 0; i < alignment.Entities.Count; i++)
                {
                    AlignmentEntity alignmentEntity = alignment.Entities.GetEntityByOrder(i);
                    switch (alignmentEntity.EntityType)
                    {
                        case AlignmentEntityType.Line:
                            {
                                AlignmentLine? alignmentLine = alignmentEntity as AlignmentLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                                station[i] = alignmentLine.Length;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                                break;
                            }
                        case AlignmentEntityType.Arc:
                            {
                                AlignmentArc? alignmentArc = alignmentEntity as AlignmentArc;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                                station[i] = alignmentArc.Length;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                                break;
                            }
                    }
                }
                double startStation = station[0];
                double endStation = station[0] + station[1];

                // set start and end point for corridor region
                ObjectId profileId = alignment.GetProfileIds()[0];
                Profile? profile = tr.GetObject(profileId, OpenMode.ForRead) as Profile;

                //check baseline exist
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                string baselineName = "BL-" + alignment.Name + "-" + profile.Name;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                foreach (Baseline BL in corridor.Baselines)
                {
                    if (BL.Name == baselineName)
                    {
                        corridor.Baselines.Remove(corridor.Baselines[baselineName]);
                    }
                }

                // then add it again
                Baseline baselineAdd = corridor.Baselines.Add("BL-" + alignment.Name + "-" + profile.Name, alignment.Id, profileId);

                // create corridor region
                string regionName = "RG-" + alignment.Name + "-" + startStation.ToString() + "-" + endStation.ToString();
                ObjectId assemblyId = A.Cdoc.AssemblyCollection["Curb Return Fillets"];
                Assembly? assembly = tr.GetObject(assemblyId, OpenMode.ForRead) as Assembly;
                BaselineRegion baselineRegion = baselineAdd.BaselineRegions.Add(regionName, assemblyId, startStation, endStation);

                //set frequency for assembly
                SetFrequencySection(baselineRegion, 2);
                ObjectId profileId_1 = alignment1.GetProfileIds()[0];
                ObjectId profileId_2 = alignment2.GetProfileIds()[0];

                // set target 0
                ObjectIdCollection TagetIds_0 = [alignment1.Id, alignment2.Id];

                // set target 1
                ObjectIdCollection TagetIds_1 = [profileId_1, profileId_2];

                // set target 1
                ObjectIdCollection TagetIds_3 = [polyline.Id];


                //set target for subassemble
                SubassemblyTargetInfoCollection subassemblyTargetInfoCollection = baselineRegion.GetTargets();
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
                A.Ed.WriteMessage(e.Message);
            }
        }

        public static void TaoCooridorDuongGiaoThong(Alignment alignment, Corridor corridor, Alignment alignment1, Alignment alignment2, CivSurface polyline)
        {
            // start transantion
            using Transaction tr = A.Db.TransactionManager.StartTransaction();
            try
            {
                UserInput UI = new();
                UtilitiesCAD CAD = new();
                UtilitiesC3D C3D = new();

                //start here


                //ObjectId alignmentId = UI.G_alignmentId("for create corridor");
                //Alignment alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;

                //get station from alignment
                double[] station = new double[alignment.Entities.Count];
                for (int i = 0; i < alignment.Entities.Count; i++)
                {
                    AlignmentEntity alignmentEntity = alignment.Entities.GetEntityByOrder(i);
                    switch (alignmentEntity.EntityType)
                    {
                        case AlignmentEntityType.Line:
                            {
                                AlignmentLine? alignmentLine = alignmentEntity as AlignmentLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                                station[i] = alignmentLine.Length;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                                break;
                            }
                        case AlignmentEntityType.Arc:
                            {
                                AlignmentArc? alignmentArc = alignmentEntity as AlignmentArc;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                                station[i] = alignmentArc.Length;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                                break;
                            }
                    }
                }
                double startStation = station[0];
                double endStation = station[0] + station[1];

                // set start and end point for corridor region
                ObjectId profileId = alignment.GetProfileIds()[0];
                Profile? profile = tr.GetObject(profileId, OpenMode.ForRead) as Profile;

                //check baseline exist
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                string baselineName = "BL-" + alignment.Name + "-" + profile.Name;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                foreach (Baseline BL in corridor.Baselines)
                {
                    if (BL.Name == baselineName)
                    {
                        corridor.Baselines.Remove(corridor.Baselines[baselineName]);
                    }
                }

                // then add it again
                Baseline baselineAdd = corridor.Baselines.Add("BL-" + alignment.Name + "-" + profile.Name, alignment.Id, profileId);

                // create corridor region
                string regionName = "RG-" + alignment.Name + "-" + startStation.ToString() + "-" + endStation.ToString();
                ObjectId assemblyId = A.Cdoc.AssemblyCollection["Curb Return Fillets"];
                Assembly? assembly = tr.GetObject(assemblyId, OpenMode.ForRead) as Assembly;
                BaselineRegion baselineRegion = baselineAdd.BaselineRegions.Add(regionName, assemblyId, startStation, endStation);

                //set frequency for assembly
                SetFrequencySection(baselineRegion, 2);
                ObjectId profileId_1 = alignment1.GetProfileIds()[0];
                ObjectId profileId_2 = alignment2.GetProfileIds()[0];

                // set target 0
                ObjectIdCollection TagetIds_0 = [alignment1.Id, alignment2.Id];

                // set target 1
                ObjectIdCollection TagetIds_1 = [profileId_1, profileId_2];

                // set target 1
                ObjectIdCollection TagetIds_3 = [polyline.Id];


                //set target for subassemble
                SubassemblyTargetInfoCollection subassemblyTargetInfoCollection = baselineRegion.GetTargets();
                int j = 1;
                subassemblyTargetInfoCollection[j].TargetIds = TagetIds_0;
                int k = 0;
                subassemblyTargetInfoCollection[k].TargetIds = TagetIds_1;
                int m = 4;
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
                A.Ed.WriteMessage(e.Message);
            }
        }

        public static void GCoordinatePointFromAlignment(Alignment alignment, int sampleLineGroupIndex, out string[] samplelineName, out string[] eastings, out string[] northings, out string[] stations)
        {
            Transaction tr = A.Db.TransactionManager.StartTransaction();

            //get samplelineGroup
            ObjectId sampleLineGroupId = alignment.GetSampleLineGroupIds()[sampleLineGroupIndex];
            SampleLineGroup? sampleLineGroup = tr.GetObject(sampleLineGroupId, OpenMode.ForRead) as SampleLineGroup;

            // get coordinate and name of point
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            int i = sampleLineGroup.GetSampleLineIds().Count;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
            stations = new string[i];
            _ = new string[i];
            samplelineName = new string[i];
            eastings = new string[i];
            northings = new string[i];
            int SamplelineIndex = 0;
            foreach (ObjectId sampleLineId in sampleLineGroup.GetSampleLineIds())
            {
                SampleLine? sampleline = tr.GetObject(sampleLineId, OpenMode.ForWrite) as SampleLine;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                double station = sampleline.Station;
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                double easting = 0;
                double northing = 0;
                alignment.PointLocation(station, 0, ref easting, ref northing);
                samplelineName[SamplelineIndex] = sampleline.Name;
                eastings[SamplelineIndex] = Convert.ToString(Math.Round(easting, 2));
                northings[SamplelineIndex] = Convert.ToString(Math.Round(northing, 2));
                stations[SamplelineIndex] = Convert.ToString(Math.Round(station, 2));
                SamplelineIndex++;
            }
            tr.Commit();
        }

        public static void SetDefaultPointSetting(string stylePoint, string stylelabel)
        {
            Transaction tr = A.Db.TransactionManager.StartTransaction();
            {
                _ = new UserInput();
                _ = new UtilitiesCAD();
                _ = new UtilitiesC3D();
                //start here

                // get the SettingsPoint object
                SettingsPoint pointSettings = A.Cdoc.Settings.GetSettings<SettingsPoint>() as SettingsPoint;

                // now set the value for default Style and Label Style
                // make sure the values exists in DWG file before you try to set them
                pointSettings.Styles.Point.Value = stylePoint;
                pointSettings.Styles.PointLabel.Value = stylelabel;

                tr.Commit();
            }
        }

        public static void AddSectionBand(SectionView sectionView, string SectionViewSectionDataBandStyles, int orderBand, ObjectId Section1Id, ObjectId Section2Id, double Gap, double Weeding)
        {
            SectionViewBandItemCollection bottomBandItems = sectionView.Bands.GetBottomBandItems();
            ObjectId bandStyleId = A.Cdoc.Styles.BandStyles.SectionViewSectionDataBandStyles[SectionViewSectionDataBandStyles];
            bottomBandItems.Add(bandStyleId);
            bottomBandItems[orderBand].Gap = Gap;
            bottomBandItems[orderBand].Section1Id = Section1Id;
            bottomBandItems[orderBand].Section2Id = Section2Id;
            bottomBandItems[orderBand].Weeding = Weeding;
            sectionView.Bands.SetBottomBandItems(bottomBandItems);
        }

        public static void RemoveSectionBand(SectionView sectionView, string SectionViewSectionDataBandStyles)
        {
            SectionViewBandItemCollection bottomBandItems = sectionView.Bands.GetBottomBandItems();
            ObjectId bandStyleId = A.Cdoc.Styles.BandStyles.SectionViewSectionDataBandStyles[SectionViewSectionDataBandStyles];
            for (int i = 0; i < bottomBandItems.Count; i++)
            {
                if (bottomBandItems[i].BandStyleId == bandStyleId)
                {
                    bottomBandItems.RemoveAt(i);
                }
            }
            sectionView.Bands.SetBottomBandItems(bottomBandItems);
        }

        public static void DrawSectionViewGroup(ObjectId alignmentId, Point3d point3D)
        {
            Transaction tr = A.Db.TransactionManager.StartTransaction();
            {
                _ = new UserInput();
                _ = new UtilitiesCAD();
                _ = new UtilitiesC3D();
                //start here
                Alignment? alignment = tr.GetObject(alignmentId, OpenMode.ForRead) as Alignment;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                double startStation = alignment.StartingStation;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                double endstation = alignment.EndingStation;
                ObjectId sampleLineGroupId = alignment.GetSampleLineGroupIds()[0];
                SampleLineGroup? sampleLineGroup = tr.GetObject(sampleLineGroupId, OpenMode.ForWrite) as SampleLineGroup;


                SectionViewGroupCreationRangeOptions sectionViewGroupCreationRangeOptions = new(sampleLineGroupId);

                ObjectIdCollection sectionSourceIdColl = [];
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                SectionSourceCollection sectionSources = sampleLineGroup.GetSectionSources();
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                //Add_SectionSource
                foreach (SectionSource sectionsource in sectionSources)
                {


                    if (sectionsource.SourceType == SectionSourceType.TinSurface)
                    {
                        TinSurface? type = tr.GetObject(sectionsource.SourceId, OpenMode.ForRead) as TinSurface;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        if (type.Name.Contains("TN"))
                        {
                            sectionsource.IsSampled = true;
                            sectionSourceIdColl.Add(sectionsource.SourceId);
                            sectionsource.StyleId = A.Cdoc.Styles.SectionStyles["1.TN Ground"];

                        }
                        else sectionsource.IsSampled = false;
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                    }

                    if (sectionsource.SourceType == SectionSourceType.Corridor)
                    {
                        Corridor? type = tr.GetObject(sectionsource.SourceId, OpenMode.ForRead) as Corridor;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        if (type.Name.Contains("MoHinh"))
                        {
                            sectionsource.IsSampled = true;
                            //sectionSourceIdColl.Add(sectionsource.SourceId);
                            sectionsource.StyleId = A.Cdoc.Styles.CodeSetStyles["1. All Codes 1-1000"];
                        }
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                    }

                    if (sectionsource.SourceType == SectionSourceType.CorridorSurface)
                    {
                        TinSurface? type = tr.GetObject(sectionsource.SourceId, OpenMode.ForRead) as TinSurface;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        if (type.Name.Contains("top"))
                        {
                            sectionsource.IsSampled = true;
                            sectionSourceIdColl.Add(sectionsource.SourceId);
                            sectionsource.StyleId = A.Cdoc.Styles.SectionStyles["2.Top Ground"];
                        }
                        else if (type.Name.Contains("datum"))
                        {
                            sectionsource.IsSampled = true;
                            sectionSourceIdColl.Add(sectionsource.SourceId);
                            sectionsource.StyleId = A.Cdoc.Styles.SectionStyles["3.Datum Ground"];
                        }
                        else sectionsource.IsSampled = false;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    }

                    if (sectionsource.SourceType == SectionSourceType.Material)
                    {

                    }

                }


                // ve trắc ngang
                SectionViewGroupCreationPlacementOptions sectionViewGroupCreationPlacementOptions = new();
                string templatePath = GetSafeTemplatePath("C:/Windows/LAYOUT CIVIL 3D.dwt");
                if (string.IsNullOrEmpty(templatePath)) return;
                sectionViewGroupCreationPlacementOptions.UseProductionPlacement(templatePath, "A3-TN-1-200");
            
            SectionViewGroupCollection sectionViewGroupCollection = sampleLineGroup.SectionViewGroups;
                sectionViewGroupCollection.Add(point3D, startStation, endstation, sectionViewGroupCreationRangeOptions, sectionViewGroupCreationPlacementOptions);


                SectionViewGroup sectionViewGroup0 = sectionViewGroupCollection[0];

                ObjectIdCollection sectionViewIdColl = sectionViewGroup0.GetSectionViewIds();

                foreach (ObjectId sectionviewId in sectionViewIdColl)
                {
                    SectionView? sectionView = tr.GetObject(sectionviewId, OpenMode.ForWrite) as SectionView;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS0618 // Type or member is obsolete
                    var sampleLineId = sectionView.ParentEntityId;
#pragma warning restore CS0618 // Type or member is obsolete
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    SampleLine? sampleLine = tr.GetObject(sampleLineId, OpenMode.ForWrite) as SampleLine;


                    // add section label

                    foreach (ObjectId sectionsourcceId in sectionSourceIdColl)
                    {
                        TinSurface? tinSurface = tr.GetObject(sectionsourcceId, OpenMode.ForWrite) as TinSurface;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        if (tinSurface.Name.Contains("TN"))
                        {
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                            ObjectId section_TN_Id = sampleLine.GetSectionId(sectionsourcceId);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                            ObjectId section_TN_labelId = SectionGradeBreakLabelGroup.Create(sectionviewId, section_TN_Id, A.Cdoc.Styles.LabelStyles.SectionLabelStyles.GradeBreakLabelStyles["Duong giong (EG)"]);
                            SectionGradeBreakLabelGroup? sectionGradeBreakLabelGroup_TN = tr.GetObject(section_TN_labelId, OpenMode.ForWrite) as SectionGradeBreakLabelGroup;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS0618 // Type or member is obsolete
                            sectionGradeBreakLabelGroup_TN.DefaultDimensionAnchorOption = Autodesk.Civil.DimensionAnchorOptionType.ViewBottom;
#pragma warning restore CS0618 // Type or member is obsolete
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                            sectionGradeBreakLabelGroup_TN.Weeding = 1;

                        }
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                        if (tinSurface.Name.Contains("top"))
                        {
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                            ObjectId section_TN_Id = sampleLine.GetSectionId(sectionsourcceId);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                            ObjectId section_TN_labelId = SectionGradeBreakLabelGroup.Create(sectionviewId, section_TN_Id, A.Cdoc.Styles.LabelStyles.SectionLabelStyles.GradeBreakLabelStyles["Duong giong"]);
                            SectionGradeBreakLabelGroup? sectionGradeBreakLabelGroup_TN = tr.GetObject(section_TN_labelId, OpenMode.ForWrite) as SectionGradeBreakLabelGroup;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS0618 // Type or member is obsolete
                            sectionGradeBreakLabelGroup_TN.DefaultDimensionAnchorOption = Autodesk.Civil.DimensionAnchorOptionType.ViewBottom;
#pragma warning restore CS0618 // Type or member is obsolete
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                            sectionGradeBreakLabelGroup_TN.Weeding = 0.7;
                        }
                    }


                    //sectionId
                    ObjectId sectionTnId = new();
                    ObjectId sectionTopId = new();
                    _ = new ObjectId();

                    foreach (ObjectId sectionsourcceId in sectionSourceIdColl)
                    {
                        TinSurface? tinSurface = tr.GetObject(sectionsourcceId, OpenMode.ForWrite) as TinSurface;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                        if (tinSurface.Name.Contains("TN"))
                        {
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                            sectionTnId = sampleLine.GetSectionId(sectionsourcceId);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                        }
#pragma warning restore CS8602 // Dereference of a possibly null reference.

                        if (tinSurface.Name.Contains("top"))
                        {
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                            sectionTopId = sampleLine.GetSectionId(sectionsourcceId);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                        }

                        if (tinSurface.Name.Contains("datum"))
                        {
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                            _ = sampleLine.GetSectionId(sectionsourcceId);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                        }

                    }

                    //add band
                    AddSectionBand(sectionView, "Cao do thiet ke 1-1000", 0, sectionTopId, sectionTnId, 0, 0.7);
                    AddSectionBand(sectionView, "Khoang cach mia TK 1-1000", 1, sectionTopId, sectionTnId, 0, 0.7);
                    AddSectionBand(sectionView, "Cao do tu nhien 1-1000", 2, sectionTnId, sectionTnId, 0, 1);
                    AddSectionBand(sectionView, "Khoang cach mia TN 1-1000", 3, sectionTnId, sectionTnId, 0, 1);

                    tr.Commit();
                }
            }
        }

        public static ObjectIdCollection GPointIdsFromPointGroup(ObjectId pointGroupId)
        {
            PointGroup? pointGroup = pointGroupId.GetObject(OpenMode.ForWrite) as PointGroup;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            uint[] numberPoint = pointGroup.GetPointNumbers();
#pragma warning restore CS8602 // Dereference of a possibly null reference.
            CogoPointCollection cogoPointCollection = A.Cdoc.CogoPoints;
            ObjectIdCollection pointIds = [];
            for (int i = 0; i < numberPoint.Length; i++)
            {
                pointIds.Add(cogoPointCollection.GetPointByPointNumber(numberPoint[i]));
            }
            return pointIds;
        }

        public static ObjectIdCollection GPointOnAlignment(ObjectIdCollection pointIds, ObjectId alignmentId, double delta)
        {
            ObjectIdCollection pointOnAlignmentId = [];
            double station = new();
            double offset = new();
            foreach (ObjectId pointId in pointIds)
            {
                CogoPoint? cogoPoint = pointId.GetObject(OpenMode.ForWrite) as CogoPoint;
                //check point on alignment
                Alignment? alignment = alignmentId.GetObject(OpenMode.ForWrite) as Alignment;
                try
                {
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    alignment.StationOffset(cogoPoint.Easting, cogoPoint.Northing, ref station, ref offset);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                    if (Math.Abs(offset) < delta)
                    {
                        pointOnAlignmentId.Add(pointId);
                        A.Ed.WriteMessage(offset.ToString() + "\n");
                    }
                    

                }
                catch
                {

                }

            }
            return pointOnAlignmentId;
        }

        public static void UpdateAllPointGroup()
        {
            PointGroupCollection objectIds = A.Cdoc.PointGroups;
            objectIds.UpdateAllPointGroups();
        }

        public static ObjectId GSectionSurfaceId(SectionSourceCollection sectionSources, string maContain)
        {
            Transaction tr = A.Db.TransactionManager.StartTransaction();
            //start here
            ObjectId objectId = new();
            foreach (SectionSource sectionsource in sectionSources)
            {
                if (sectionsource.SourceType == SectionSourceType.TinSurface & sectionsource.IsSampled == true)
                {
                    TinSurface? type = tr.GetObject(sectionsource.SourceId, OpenMode.ForRead) as TinSurface;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    if (type.Name.Contains(maContain))
                    {

                        objectId = sectionsource.SourceId;

                    }
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                }
            }
            return objectId;
        }

        public static double FindY(SectionPointCollection sectionPoints, double x, double X, double Y, double Z)
        {

            double y = 0;
            for (int i = 0; i < sectionPoints.Count - 1; i++)
            {
                double x1 = sectionPoints[i].Location.X + X;
                double x2 = sectionPoints[i + 1].Location.X + X;
                double y1 = sectionPoints[i].Location.Y + Y - Z;
                double y2 = sectionPoints[i + 1].Location.Y + Y - Z;
                    
                if (x1 <= x & x <= x2)
                {
                    double at = (y2 - y1) / (x2 - x1);
                    double b = -x1 * (y2 - y1) / (x2 - x1) + y1;
                    y = at * x + b;
                }
            }
            return y;
        }

        public static PointGroup CPointGroupWithDecription(string nameGroup, string Decription)
        {
            Transaction tr = A.Db.TransactionManager.StartTransaction();
            {
                _ = new UserInput();
                _ = new UtilitiesCAD();
                _ = new UtilitiesC3D();
                //start here

                //create point group Name
                PointGroup? pointGroup_Out = null;
                PointGroupCollection pointGroupColl = A.Cdoc.PointGroups;
                if (!pointGroupColl.Contains(nameGroup))
                {
                    pointGroupColl.Add(nameGroup);
                }

                //set query
                StandardPointGroupQuery query = new()
                {
                    IncludeRawDescriptions = Decription
                };

                //set decription for point group
                foreach (ObjectId pointGroupId in pointGroupColl)
                {
                    PointGroup? pointGroup = tr.GetObject(pointGroupId, OpenMode.ForWrite) as PointGroup;
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    if (pointGroup.Name == nameGroup)
                    {
                        pointGroup.SetQuery(query);
                        pointGroup_Out = pointGroup;
                    }
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                }

                tr.Commit();
#pragma warning disable CS8603 // Possible null reference return.
                return pointGroup_Out;
#pragma warning restore CS8603 // Possible null reference return.
            }

        }


        public static ObjectId CreateSampleline(string sampleLineName, ObjectId sampleLineGroupId, Alignment alignment, double station)
        {
            Transaction tr = A.Db.TransactionManager.StartTransaction();
            {
                _ = new UserInput();
                _ = new UtilitiesCAD();
                _ = new UtilitiesC3D();
                //start here

                //lấy khoảng dịch 2 bên
                Point2dCollection point2Ds = [];
                double easting = new();
                double northing = new();
                alignment.PointLocation(station, -10, ref easting, ref northing);
                Point2d point2D = new(easting, northing);
                point2Ds.Add(point2D);
                alignment.PointLocation(station, 10, ref easting, ref northing);
                Point2d point2D1 = new(easting, northing);
                point2Ds.Add(point2D1);

                //tao sampleline
                ObjectId sampleLineId = SampleLine.Create(sampleLineName, sampleLineGroupId, point2Ds);

                tr.Commit();
                return sampleLineId;
            }
        }

















        public static string GetSafeTemplatePath(string preferredPath)
        {
            if (System.IO.File.Exists(preferredPath)) return preferredPath;

            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\n[Cảnh báo] Không tìm thấy template tại: {preferredPath}");
            
            PromptOpenFileOptions pfo = new("Chọn template trắc ngang (.dwt)")
            {
                Filter = "Template files (*.dwt)|*.dwt|Drawing files (*.dwg)|*.dwg"
            };
            PromptFileNameResult pfr = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.GetFileNameForOpen(pfo);
            
            if (pfr.Status == PromptStatus.OK) return pfr.StringResult;
            
            return ""; // Trả về trống sẽ gây lỗi API Civil 3D, nên có thể cân nhắc throw hoặc dùng temp mặc định
        }

    }
}
