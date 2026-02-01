using Autodesk.Aec.PropertyData.DatabaseServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using MyFirstProject.Extensions;

namespace MyFirstProject.Extensions
{
    public class PropertySetUtils
    {
        /// <summary>
        /// Thi?t l?p Property Set cho 3D Solid v?i các thu?c tính ???c tính toán
        /// </summary>
        /// <param name="tr">Transaction</param>
        /// <param name="solid">3D Solid object</param>
        public void SetupSolidWithCalculatedProperties(Transaction tr, Solid3d solid)
        {
            if (solid == null) return;

            try
            {
                // L?y thông tin c? b?n c?a solid
                string layerName = solid.Layer;
                double volume = solid.MassProperties.Volume;
                Point3d centroid = solid.MassProperties.Centroid;

                // S? d?ng tên Property Set c? ??nh
                string propertySetName = "IFC ???NG GIAO THÔNG2";

                // Ki?m tra và t?o Property Set Definition n?u ch?a có
                ObjectId propertySetDefId = GetOrCreatePropertySetDefinition(tr, propertySetName);

                if (propertySetDefId.IsNull) return;

                // Attach Property Set vào solid
                AttachPropertySetToObject(tr, solid, propertySetDefId);

                // Set các giá tr? properties
                SetPropertyValues(tr, solid, propertySetDefId, new Dictionary<string, object>
                {
                    { "Layer", layerName },
                    { "Volume", Math.Round(volume, 3) },
                    { "CentroidX", Math.Round(centroid.X, 3) },
                    { "CentroidY", Math.Round(centroid.Y, 3) },
                    { "CentroidZ", Math.Round(centroid.Z, 3) },
                    { "CreatedDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") }
                });
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\nL?i khi thi?t l?p Property Set: {ex.Message}");
            }
        }

        /// <summary>
        /// L?y ho?c t?o Property Set Definition
        /// </summary>
        private ObjectId GetOrCreatePropertySetDefinition(Transaction tr, string propertySetName)
        {
            try
            {
                Database db = A.Db;
                DictionaryPropertySetDefinitions propSetDefs = new(db);

                // Ki?m tra xem Property Set Definition ?ã t?n t?i ch?a
                if (propSetDefs.Has(propertySetName, tr))
                {
                    return propSetDefs.GetAt(propertySetName);
                }

                // T?o m?i Property Set Definition
                PropertySetDefinition propSetDef = new();
                propSetDef.SubSetDatabaseDefaults(db);
                propSetDef.Description = "Property Set cho ???ng giao thông IFC";

                // Thêm các Property Definitions
                AddPropertyDefinitions(propSetDef);

                // Thi?t l?p Applies To
                StringCollection appliesToFilter = new();
                appliesToFilter.Add("AcDb3dSolid");
                propSetDef.SetAppliesToFilter(appliesToFilter, false);

                propSetDefs.AddNewRecord(propertySetName, propSetDef);
                tr.AddNewlyCreatedDBObject(propSetDef, true);

                return propSetDef.ObjectId;
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\nL?i khi t?o Property Set Definition: {ex.Message}");
                return ObjectId.Null;
            }
        }

        /// <summary>
        /// Thêm các Property Definitions vào Property Set Definition
        /// </summary>
        private void AddPropertyDefinitions(PropertySetDefinition propSetDef)
        {
            // Property: C?u ki?n
            PropertyDefinition cauKienProp = new();
            cauKienProp.Name = "C?u ki?n";
            cauKienProp.Description = "Lo?i c?u ki?n";
            cauKienProp.DataType = Autodesk.Aec.PropertyData.DataType.Text;
            cauKienProp.DefaultData = "";
            propSetDef.Definitions.Add(cauKienProp);

            // Property: V?t li?u
            PropertyDefinition vatLieuProp = new();
            vatLieuProp.Name = "V?t li?u";
            vatLieuProp.Description = "Lo?i v?t li?u";
            vatLieuProp.DataType = Autodesk.Aec.PropertyData.DataType.Text;
            vatLieuProp.DefaultData = "";
            propSetDef.Definitions.Add(vatLieuProp);

            // Property: Th? tích (m3)
            PropertyDefinition theTichProp = new();
            theTichProp.Name = "Th? tích (m3)";
            theTichProp.Description = "Th? tích c?a ??i t??ng";
            theTichProp.DataType = Autodesk.Aec.PropertyData.DataType.Real;
            theTichProp.DefaultData = 0.0;
            propSetDef.Definitions.Add(theTichProp);

            // Property: CentroidX
            PropertyDefinition centroidXProp = new();
            centroidXProp.Name = "CentroidX";
            centroidXProp.Description = "T?a ?? X tr?ng tâm";
            centroidXProp.DataType = Autodesk.Aec.PropertyData.DataType.Real;
            centroidXProp.DefaultData = 0.0;
            propSetDef.Definitions.Add(centroidXProp);

            // Property: CentroidY
            PropertyDefinition centroidYProp = new();
            centroidYProp.Name = "CentroidY";
            centroidYProp.Description = "T?a ?? Y tr?ng tâm";
            centroidYProp.DataType = Autodesk.Aec.PropertyData.DataType.Real;
            centroidYProp.DefaultData = 0.0;
            propSetDef.Definitions.Add(centroidYProp);

            // Property: CentroidZ
            PropertyDefinition centroidZProp = new();
            centroidZProp.Name = "CentroidZ";
            centroidZProp.Description = "T?a ?? Z tr?ng tâm";
            centroidZProp.DataType = Autodesk.Aec.PropertyData.DataType.Real;
            centroidZProp.DefaultData = 0.0;
            propSetDef.Definitions.Add(centroidZProp);

            // Property: CreatedDate
            PropertyDefinition dateProp = new();
            dateProp.Name = "CreatedDate";
            dateProp.Description = "Ngày t?o";
            dateProp.DataType = Autodesk.Aec.PropertyData.DataType.Text;
            dateProp.DefaultData = "";
            propSetDef.Definitions.Add(dateProp);

            // Property: Layer
            PropertyDefinition layerProp = new();
            layerProp.Name = "Layer";
            layerProp.Description = "Layer c?a ??i t??ng";
            layerProp.DataType = Autodesk.Aec.PropertyData.DataType.Text;
            layerProp.DefaultData = "";
            propSetDef.Definitions.Add(layerProp);

            // Property: Volume
            PropertyDefinition volumeProp = new();
            volumeProp.Name = "Volume";
            volumeProp.Description = "Th? tích (m³)";
            volumeProp.DataType = Autodesk.Aec.PropertyData.DataType.Real;
            volumeProp.DefaultData = 0.0;
            propSetDef.Definitions.Add(volumeProp);
        }

        /// <summary>
        /// Attach Property Set vào object
        /// </summary>
        private void AttachPropertySetToObject(Transaction tr, DBObject dbObject, ObjectId propertySetDefId)
        {
            try
            {
                PropertyDataServices.AddPropertySet(dbObject, propertySetDefId);
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\nL?i khi attach Property Set: {ex.Message}");
            }
        }

        /// <summary>
        /// Set giá tr? cho các properties
        /// </summary>
        private void SetPropertyValues(Transaction tr, DBObject dbObject, ObjectId propertySetDefId, Dictionary<string, object> values)
        {
            try
            {
                ObjectIdCollection propertySetIds = PropertyDataServices.GetPropertySets(dbObject);
                
                foreach (ObjectId propSetId in propertySetIds)
                {
                    PropertySet propSet = tr.GetObject(propSetId, OpenMode.ForWrite) as PropertySet;
                    if (propSet?.PropertySetDefinition == propertySetDefId)
                    {
                        PropertySetDefinition propSetDef = tr.GetObject(propertySetDefId, OpenMode.ForRead) as PropertySetDefinition;
                        
                        foreach (var kvp in values)
                        {
                            try
                            {
                                // Tìm index c?a property theo tên
                                int propertyIndex = -1;
                                for (int i = 0; i < propSetDef.Definitions.Count; i++)
                                {
                                    if (propSetDef.Definitions[i].Name == kvp.Key)
                                    {
                                        propertyIndex = i;
                                        break;
                                    }
                                }
                                
                                if (propertyIndex >= 0)
                                {
                                    propSet.SetAt(propertyIndex, kvp.Value);
                                }
                            }
                            catch (System.Exception ex)
                            {
                                A.Ed.WriteMessage($"\nL?i khi set property {kvp.Key}: {ex.Message}");
                            }
                        }

                        // Set giá tr? cho "Th? tích (m3)" t? Volume
                        if (values.ContainsKey("Volume"))
                        {
                            try
                            {
                                int theTichIndex = -1;
                                for (int i = 0; i < propSetDef.Definitions.Count; i++)
                                {
                                    if (propSetDef.Definitions[i].Name == "Th? tích (m3)")
                                    {
                                        theTichIndex = i;
                                        break;
                                    }
                                }
                                
                                if (theTichIndex >= 0)
                                {
                                    propSet.SetAt(theTichIndex, values["Volume"]);
                                }
                            }
                            catch (System.Exception ex)
                            {
                                A.Ed.WriteMessage($"\nL?i khi set property Th? tích (m3): {ex.Message}");
                            }
                        }
                        break;
                    }
                }
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\nL?i khi set property values: {ex.Message}");
            }
        }

        /// <summary>
        /// Hi?n th? thông tin Property Set c?a 3D Solid
        /// </summary>
        public void ShowSolidPropertySetInfo(Transaction tr, Solid3d solid)
        {
            try
            {
                A.Ed.WriteMessage($"\n=== THÔNG TIN 3D SOLID ===");
                A.Ed.WriteMessage($"\nLayer: {solid.Layer}");
                A.Ed.WriteMessage($"\nHandle: {solid.Handle}");
                
                var massProps = solid.MassProperties;
                A.Ed.WriteMessage($"\nTh? tích: {massProps.Volume:F3} m³");
                A.Ed.WriteMessage($"\nTr?ng tâm: X={massProps.Centroid.X:F3}, Y={massProps.Centroid.Y:F3}, Z={massProps.Centroid.Z:F3}");

                // Hi?n th? Property Sets
                ObjectIdCollection propertySetIds = PropertyDataServices.GetPropertySets(solid);
                if (propertySetIds.Count > 0)
                {
                    A.Ed.WriteMessage($"\n\n=== PROPERTY SETS ({propertySetIds.Count}) ===");
                    
                    foreach (ObjectId propSetId in propertySetIds)
                    {
                        PropertySet propSet = tr.GetObject(propSetId, OpenMode.ForRead) as PropertySet;
                        if (propSet != null)
                        {
                            PropertySetDefinition propSetDef = tr.GetObject(propSet.PropertySetDefinition, OpenMode.ForRead) as PropertySetDefinition;
                            A.Ed.WriteMessage($"\n\nProperty Set: {propSetDef?.Name}");
                            A.Ed.WriteMessage($"Mô t?: {propSetDef?.Description}");
                            
                            if (propSetDef != null)
                            {
                                for (int i = 0; i < propSetDef.Definitions.Count; i++)
                                {
                                    PropertyDefinition propDef = propSetDef.Definitions[i];
                                    try
                                    {
                                        object value = propSet.GetAt(i);
                                        A.Ed.WriteMessage($"\n  - {propDef.Name}: {value} ({propDef.Description})");
                                    }
                                    catch
                                    {
                                        A.Ed.WriteMessage($"\n  - {propDef.Name}: <không có giá tr?> ({propDef.Description})");
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    A.Ed.WriteMessage($"\n\nKhông có Property Set nào ???c g?n vào ??i t??ng này.");
                }
                
                A.Ed.WriteMessage($"\n================================\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\nL?i khi hi?n th? thông tin: {ex.Message}");
            }
        }
    }
}