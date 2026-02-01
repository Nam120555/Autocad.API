using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using System;
using System.Linq;

namespace Civil3DCsharp
{
    /// <summary>
    /// Extension methods for Civil 3D objects
    /// </summary>
    public static class Civil3DExtensions
    {
        /// <summary>
        /// Gets the parent alignment ID for a sample line
        /// </summary>
        /// <param name="sampleLine">The sample line</param>
        /// <returns>ObjectId of the parent alignment</returns>
        public static ObjectId GetParentAlignmentId(this SampleLine sampleLine)
        {
            ArgumentNullException.ThrowIfNull(sampleLine);
            
            foreach (ObjectId alignmentId in CivilApplication.ActiveDocument.GetAlignmentIds())
            {
                using var alignment = alignmentId.GetObject(OpenMode.ForRead) as Alignment;
                if (alignment == null) continue;
                
                foreach (ObjectId sampleLineGroupId in alignment.GetSampleLineGroupIds())
                {
                    if (sampleLineGroupId == sampleLine.GroupId)
                    {
                        return alignmentId;
                    }
                }
            }
            return ObjectId.Null;
        }

        /// <summary>
        /// Gets the section view group object for a section view
        /// </summary>
        /// <param name="sectionView">The section view</param>
        /// <returns>The section view group object</returns>
        public static SectionViewGroup? GetSectionViewGroupObject(this SectionView sectionView)
        {
            var db = sectionView.Database;
            using var tr = db.TransactionManager.StartTransaction();
            
            ObjectId sampleLineId = sectionView.SampleLineId;
            var sampleLineGroup = tr.GetObject(sampleLineId, OpenMode.ForRead) as SampleLineGroup;
            if (sampleLineGroup == null) return null;
            
            foreach (SectionViewGroup sectionViewGroup in sampleLineGroup.SectionViewGroups)
            {
                if (sectionViewGroup.GetSectionViewIds().Contains(sectionView.ObjectId))
                {
                    return sectionViewGroup;
                }
            }
            return null;
        }

        /// <summary>
        /// Legacy extension method for backward compatibility
        /// </summary>
        /// <param name="sl">The sample line</param>
        /// <returns>ObjectId of the parent alignment</returns>
        public static ObjectId ParentAlignmentId(this SampleLine sl)
        {
            return sl.GetParentAlignmentId();
        }

        /// <summary>
        /// Legacy extension method for backward compatibility
        /// </summary>
        /// <param name="sv">The section view</param>
        /// <returns>The section view group object</returns>
        public static SectionViewGroup SectionViewGroupObject(this SectionView sv)
        {
#pragma warning disable CS8603 // Possible null reference return.
            return sv.GetSectionViewGroupObject();
#pragma warning restore CS8603 // Possible null reference return.
        }
    }
}