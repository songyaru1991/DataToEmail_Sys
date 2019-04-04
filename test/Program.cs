using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace test
{
    class Program
    {
        static void Main(string[] args)
        {
            StringBuilder sb = new StringBuilder();
            DateTime sysTime = DateTime.Now.AddDays(-2);
            string[] strs = sysTime.ToString("yyyy-MM-dd").Split('/');

            for (int i = 0; i < strs.Length; i++)
            {
                sb.Append(strs[i]);
            }

          string newTime=  sysTime.ToString("yyyy-MM-dd");
            Console.WriteLine(newTime);

            Console.ReadKey();
        }
    }
}
