using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyFirstProject.Extensions
{
    public class Csharp
    {
        public static double[] CSMangSapXepTangDan(double[] arraySort)
        {
            double tg;
            for (int i = 0; i < arraySort.Length - 1; i++)
            {
                for (int j = i + 1; j < arraySort.Length; j++)
                {
                    if (arraySort[i] > arraySort[j])
                    {
                        // Hoan vi 2 so a[i] va a[j]
                        tg = arraySort[i];
                        arraySort[i] = arraySort[j];
                        arraySort[j] = tg;
                    }
                }
            }
            return arraySort;

        }
    }






}
