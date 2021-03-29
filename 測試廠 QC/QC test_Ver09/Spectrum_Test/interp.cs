using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Spectrum_Test
{
    class interp
    {
        public List<double> interp1(List<double> x, List<double> y, List<double> x_new )
        {
            List<double> y_new = new List<double>();

            List<double> dx, dy, slope, intercept;
            dx = new List<double>();
            dy = new List<double>();
            slope = new List<double>();
            intercept = new List<double>();

            for (int i = 0; i < x.Count(); ++i)
            {
                if (i < x.Count() - 1)
                {
                    dx.Add(x[i + 1] - x[i]);
                    dy.Add(y[i + 1] - y[i]);
                    slope.Add(dy[i] / dx[i]);
                    intercept.Add(y[i] - x[i] * slope[i]);
                }
                else
                {
                    dx.Add(dx[i - 1]);
                    dy.Add(dy[i - 1]);
                    slope.Add(slope[i - 1]);
                    intercept.Add(intercept[i - 1]);
                }
            }

            for (int i = 0; i < x_new.Count(); ++i)
            {
                int idx = findNearestNeighbourIndex(x_new[i], x);
                y_new.Add(slope[idx] * x_new[i] + intercept[idx]);

            }

            return y_new;
        }

        int findNearestNeighbourIndex(double value, List<double> x)
        {
            double dist = double.MaxValue;
            int idx = -1;
            for (int i = 0; i < x.Count(); ++i)
            {
                double newDist = value - x[i];
                if (newDist >= 0 && newDist <= dist)
                {
                    dist = newDist;
                    idx = i;
                }
            }

            return idx;
        }
    }
}
