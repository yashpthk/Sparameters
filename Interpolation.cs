using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CapacitanceExtractor
{
    class Interpolation:iInterpolation
    {
        /// <summary>
        /// Interpolates the specified x data.
        /// </summary>
        /// <param name="xData">The x data.</param>
        /// <param name="yData">The y data.</param>
        /// <param name="xDataPoints">The x data points.</param>
        /// <returns></returns>
        public double[][] interpolate(double[] xData, double[][] yData, double[] xDataPoints)
        {
            double[][] yDataValues = new double[yData.Length][];
            for (int k = 0; k < yData.Length; k++)
            {
                yDataValues[k] = new double[xDataPoints.Length];
                for (int i = 0; i < xDataPoints.Length; i++)
                {
                    for (int j = 0; j < xData.Length; j++)
                    {
                        double x0 = xData[j];
                        double y0 = yData[k][j];

                        if (xDataPoints[i] > x0)
                        { continue; }
                        if (xDataPoints[i] == x0)
                        {
                            yDataValues[k][i] = yData[k][j];
                            break;
                        }
                        else
                        {
                            yDataValues[k][i] = y0 + (xDataPoints[i] - xData[j - 1]) * (y0 - yData[k][j - 1]) / (x0 - xData[j - 1]);
                            break;
                        }
                    }
                } 
            }
            return yDataValues;
        }
    }
}
