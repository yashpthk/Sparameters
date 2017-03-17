using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CapacitanceExtractor
{
    interface iInterpolation
    {
        double[][] interpolate(double[] xData, double[][] yData, double[] xDataPoints);
    }
}
