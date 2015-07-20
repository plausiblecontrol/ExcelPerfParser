using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PerfParse {
  class UARG {
    public double PID;
    public string CMD;
    public UARG(double processID, string Command) {
      PID = processID;
      CMD = Command;
    }
  }
}
