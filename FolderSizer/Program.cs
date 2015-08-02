using System;
using System.Linq;

namespace FolderSizer
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Folder sizes");

            var window = FolderData.GetOpenWindows().First();

            foreach (var d in window.SubFolders.OrderByDescending(x=>x.size))
            {               
                Console.WriteLine(d.FormattedSize.PadRight(10,' ') + " " + d.Name);
            }

            Console.ReadKey();
        }

    }
}
