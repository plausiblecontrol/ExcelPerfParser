using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PerfParse {
  class process {
    public int pid;
    public int total;
    public string command;
    public List<double> resset = new List<double>();
    public List<double> cpuPercent = new List<double>();
    public process(string xPID, string xResSet, string xCPUpercent, string xCommand) {
      command = xCommand;
      total = 1;
      pid = Convert.ToInt32(xPID);
      resset.Add(Convert.ToDouble(xResSet));
      cpuPercent.Add(Convert.ToDouble(xCPUpercent));
    }
    public void add(string xResSet, string xCPUpercent) {
      resset.Add(Convert.ToDouble(xResSet));
      cpuPercent.Add(Convert.ToDouble(xCPUpercent));
      total++;
    }
    public double CPUMax() {
      return cpuPercent.Max();
    }
    public double isHigh() {
      return cpuPercent.Average();
    }


    static double Variance(List<double> list) {
      double mean = list.Average();
      double result = list.Sum(number => Math.Pow(number - mean, 2.0));
      return result / list.Count();
    }
    static double StandardDev2(List<double> list) {
      return Math.Sqrt(Variance(list));
    }
    static double StandardDev(List<double> valueList) {
      double M = 0.0;
      double S = 0.0;
      int k = 1;
      foreach (double value in valueList) {
        double tmpM = M;
        M += (value - tmpM) / k;
        S += (value - tmpM) * (value - M);
        k++;
      }
      return Math.Sqrt(S / (k - 1));
    }
    static double[] LeastSquares(List<double> valueList) {
      int numPoints = 0;
      double sumX = 0;
      double sumY = 0;
      double sumXX = 0;
      double sumXY = 0;
      //foreach (var tuple in valueList) {
      for (int i = 0; i < valueList.Count(); i++) {
        numPoints++;
        sumX += i;// tuple.Item1;
        sumY += valueList[i];// tuple.Item2;
        sumXX += i * i;//tuple.Item1 * tuple.Item1;
        sumXY += i * valueList[i];//tuple.Item1 * tuple.Item2;
      }

      double b = (-sumX * sumXY + sumXX * sumY) / (numPoints * sumXX - sumX * sumX);
      double m = (-sumX * sumY + numPoints * sumXY) / (numPoints * sumXX - sumX * sumX);

      return new double[] { m, b };
    }

    static List<double> sigma(List<double> valueList, double StdDev, double[] equation) {
      List<double> memList = new List<double>();
      for (int i = 1; i <= valueList.Count; i++) {
        double kI = equation[0] * i + equation[1];
        if (valueList[i - 1] <= (kI + StdDev) && valueList[i - 1] >= (kI - StdDev)) {
          memList.Add(valueList[i - 1]);
        } else if (valueList[i - 1] > (kI + StdDev)) {
          memList.Add(kI + StdDev);
        } else {
          memList.Add(kI - StdDev);
        }
      }
      return memList;
    }
    public double isLeaking() {
      double[] eqMB = new double[2];
      double stdD1 = StandardDev(resset);
      eqMB = LeastSquares(resset);
      List<double> resset2 = sigma(resset, stdD1, eqMB);
      double stdD2 = StandardDev(resset2);
      eqMB = LeastSquares(resset2);
      List<double> resset3 = sigma(resset2, stdD2, eqMB);
      eqMB = LeastSquares(resset3);

      return eqMB[0];
      //eqMB = LeastSquares(memList);
      //    double stdD1 = StandardDev(memList);
      //    List<double> memList2 = sigma(memList, stdD1, eqMB);
      //    cX = memList2.Count();
      //    eqMB = LeastSquares(memList2);
      //    double stdD2 = StandardDev(memList2)*1.2;
      //    List<double> memList3 = sigma(memList2, stdD2, eqMB);
      //    eqMB = LeastSquares(memList3);

      //    if (eqMB[0] < 0) {
      //      leaking = true;

      //double SumY = 0.0;
      //double SumX = 0.0;
      //double SumXY = 0.0;
      //double SumXX = 0.0;
      //for (int i = 0; i < total; i++) {
      //  SumXY += i * resset[i];
      //  SumXX += i * i;
      //  SumY += resset[i];
      //  SumX += i;
      //}
      //return ((SumXY - SumX * SumY / total) / (SumXX - SumX * SumX / total));
    }
  }
}
