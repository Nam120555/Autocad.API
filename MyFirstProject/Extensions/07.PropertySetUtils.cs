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
        /// Thi?t l?p Property Set cho 3D Solid v?i c�c thu?c t�nh ???c t�nh to�n
        /// </summary>
        /// <param name="tr">Transaction</param>
        /// <param name="solid">3D Solid object</param>
        public void SetupSolidWithCalculatedProperties(Transaction tr, Solid3d solid)
        {
            if (solid == null) return;

            try
            {
                // L?y th�ng tin c? b?n c?a solid
                string layerName = solid.Layer;
                double volume = solid.MassProperties.Volume;
                Point3d centroid = solid.MassProperties.Centroid;

                // S? d?ng t�n Property Set c? ??nh
                string propertySetName = "IFC ???NG GIAO TH�NG2";

                // Ki?m tra v� t?o Property Set Definition n?u ch?a c�
                ObjectId propertySetDefId = GetOrCreatePropertySetDefinition(tr, propertySetName);

                if (propertySetDefId.IsNull) return;

                // Attach Property Set v�o solid
                AttachPropertySetToObject(tr, solid, propertySetDefId);

                // Set c�c gi� tr? properties
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

                // Ki?m tra xem Property Set Definition ?� t?n t?i ch?a
                if (propSetDefs.Has(propertySetName, tr))
                {
                    return propSetDefs.GetAt(propertySetName);
                }

                // T?o m?i Property Set Definition
                PropertySetDefinition propSetDef = new();
                propSetDef.SubSetDatabaseDefaults(db);
                propSetDef.Description = "Property Set cho ???ng giao th�ng IFC";

                // Th�m c�c Property Definitions
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
        /// Th�m c�c Property Definitions v�o Property Set Definition
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

            // Property: Th? t�ch (m3)
            PropertyDefinition theTichProp = new();
            theTichProp.Name = "Th? t�ch (m3)";
            theTichProp.Description = "Th? t�ch c?a ??i t??ng";
            theTichProp.DataType = Autodesk.Aec.PropertyData.DataType.Real;
            theTichProp.DefaultData = 0.0;
            propSetDef.Definitions.Add(theTichProp);

            // Property: CentroidX
            PropertyDefinition centroidXProp = new();
            centroidXProp.Name = "CentroidX";
            centroidXProp.Description = "T?a ?? X tr?ng t�m";
            centroidXProp.DataType = Autodesk.Aec.PropertyData.DataType.Real;
            centroidXProp.DefaultData = 0.0;
            propSetDef.Definitions.Add(centroidXProp);

            // Property: CentroidY
            PropertyDefinition centroidYProp = new();
            centroidYProp.Name = "CentroidY";
            centroidYProp.Description = "T?a ?? Y tr?ng t�m";
            centroidYProp.DataType = Autodesk.Aec.PropertyData.DataType.Real;
            centroidYProp.DefaultData = 0.0;
            propSetDef.Definitions.Add(centroidYProp);

            // Property: CentroidZ
            PropertyDefinition centroidZProp = new();
            centroidZProp.Name = "CentroidZ";
            centroidZProp.Description = "T?a ?? Z tr?ng t�m";
            centroidZProp.DataType = Autodesk.Aec.PropertyData.DataType.Real;
            centroidZProp.DefaultData = 0.0;
            propSetDef.Definitions.Add(centroidZProp);

            // Property: CreatedDate
            PropertyDefinition dateProp = new();
            dateProp.Name = "CreatedDate";
            dateProp.Description = "Ng�y t?o";
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
            volumeProp.Description = "Th? t�ch (m�)";
            volumeProp.DataType = Autodesk.Aec.PropertyData.DataType.Real;
            volumeProp.DefaultData = 0.0;
            propSetDef.Definitions.Add(volumeProp);
        }

        /// <summary>
        /// Attach Property Set v�o object
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
        /// Set gi� tr? cho c�c properties
        /// </summary>
        private void SetPropertyValues(Transaction tr, DBObject dbObject, ObjectId propertySetDefId, Dictionary<string, object> values)
        {
            try
            {
                ObjectIdCollection propertySetIds = PropertyDataServices.GetPropertySets(dbObject);
                
                foreach (ObjectId propSetId in propertySetIds)
                {
                    PropertySet? propSet = tr.GetObject(propSetId, OpenMode.ForWrite) as PropertySet;
                    if (propSet?.PropertySetDefinition == propertySetDefId)
                    {
                        PropertySetDefinition? propSetDef = tr.GetObject(propertySetDefId, OpenMode.ForRead) as PropertySetDefinition;
                        
                        if (propSetDef == null) break;
                        foreach (var kvp in values)
                        {
                            try
                            {
                                // T�m index c?a property theo t�n
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

                        // Set gi� tr? cho "Th? t�ch (m3)" t? Volume
                        if (values.ContainsKey("Volume"))
                        {
                            try
                            {
                                int theTichIndex = -1;
                                for (int i = 0; i < propSetDef.Definitions.Count; i++)
                                {
                                    if (propSetDef.Definitions[i].Name == "Th? t�ch (m3)")
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
                                A.Ed.WriteMessage($"\nL?i khi set property Th? t�ch (m3): {ex.Message}");
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
        /// Hi?n th? th�ng tin Property Set c?a 3D Solid
        /// </summary>
        public void ShowSolidPropertySetInfo(Transaction tr, Solid3d solid)
        {
            try
            {
                A.Ed.WriteMessage($"\n=== TH�NG TIN 3D SOLID ===");
                A.Ed.WriteMessage($"\nLayer: {solid.Layer}");
                A.Ed.WriteMessage($"\nHandle: {solid.Handle}");
                
                var massProps = solid.MassProperties;
                A.Ed.WriteMessage($"\nTh? t�ch: {massProps.Volume:F3} m�");
                A.Ed.WriteMessage($"\nTr?ng t�m: X={massProps.Centroid.X:F3}, Y={massProps.Centroid.Y:F3}, Z={massProps.Centroid.Z:F3}");

                // Hi?n th? Property Sets
                ObjectIdCollection propertySetIds = PropertyDataServices.GetPropertySets(solid);
                if (propertySetIds.Count > 0)
                {
                    A.Ed.WriteMessage($"\n\n=== PROPERTY SETS ({propertySetIds.Count}) ===");
                    
                    foreach (ObjectId propSetId in propertySetIds)
                    {
                        PropertySet? propSet = tr.GetObject(propSetId, OpenMode.ForRead) as PropertySet;
                        if (propSet != null)
                        {
                            PropertySetDefinition? propSetDef = tr.GetObject(propSet.PropertySetDefinition, OpenMode.ForRead) as PropertySetDefinition;
                            A.Ed.WriteMessage($"\n\nProperty Set: {propSetDef?.Name}");
                            A.Ed.WriteMessage($"M� t?: {propSetDef?.Description}");
                            
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
                                        A.Ed.WriteMessage($"\n  - {propDef.Name}: <kh�ng c� gi� tr?> ({propDef.Description})");
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    A.Ed.WriteMessage($"\n\nKh�ng c� Property Set n�o ???c g?n v�o ??i t??ng n�y.");
                }
                
                A.Ed.WriteMessage($"\n================================\n");
            }
            catch (System.Exception ex)
            {
                A.Ed.WriteMessage($"\nL?i khi hi?n th? th�ng tin: {ex.Message}");
            }
        }
    }
}