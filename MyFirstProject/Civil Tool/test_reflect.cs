using System;
using System.Reflection;
using System.Linq;

class Program {
    static void Main() {
        try {
            var asm = Assembly.LoadFrom(@"C:\Program Files\Autodesk\AutoCAD 2024\C3D\AecBase.mgd.dll");
            var asm2 = Assembly.LoadFrom(@"C:\Program Files\Autodesk\AutoCAD 2024\C3D\AeccDbMgd.dll");
            Console.WriteLine("Loaded");
            var type = asm2.GetTypes().FirstOrDefault(t => t.Name == "AlignmentSCSGroup");
            if (type != null) {
                foreach (var p in type.GetProperties()) {
                    Console.WriteLine(p.Name + " : " + p.PropertyType.Name);
                }
            } else {
                Console.WriteLine("Type not found");
            }
        } catch (Exception ex) {
            Console.WriteLine(ex.Message);
        }
    }
}
