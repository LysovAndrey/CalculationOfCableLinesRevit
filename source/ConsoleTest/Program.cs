using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleTest {


  public class BDCable {
    public string Source { get; set; }
    public string Name { get; set; }
  }


  public class Test : HashSet<BDCable> {
    class BDCableImp {
      public BDCable Get(string source, string name) {
        return new BDCable() { Source = source, Name = name };
      }
    }

    public Test() {
      int length = 10;
      for (int i = 0; i < length; i++) {
        BDCableImp c = new BDCableImp();
        
        base.Add(c.Get($"source_{i}",$"name_{i}"));
      }
    }

    public List<BDCable> GetCable() {
      return this.ToList();
    }

  }




  internal class Program {
    static void Main(string[] args) {

      Test test = new Test();
      List<BDCable> bDCables = test.GetCable();
      Console.WriteLine($"Количество - {test.Count}");

      foreach (var item in bDCables) {
        Console.WriteLine($"{item.Name} - {item.Source}");
      }

      Console.ReadLine();
    }
  }
}
