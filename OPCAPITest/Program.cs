using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OPC;
using OPC.DA;
using OPC.Enum;

namespace OPCAPITest
{
    class Program
    {
        static void Main(string[] args)
        {
            Server s;

            using (Enumerator enumerator = new Enumerator())
            {
                IList<Server> servers = enumerator.GetAvailableServers(Specification.OPC_DA30);

                s = servers.Where(x => x.ProgId.Contains("Kepware")).First();
            }

            using (DAServer server = new DAServer())
            {
                server.Connect(s);

                var items = server.Browse(String.Empty, BrowseType.Branch);

                items = server.Browse(items[3].ItemID, BrowseType.Branch);

                items = server.Browse(items[2].ItemID, BrowseType.Leaf);

                Subscription subscription = server.AddSubscription("Subscription 1", 30, 0.0f, true);

                subscription.DataChanged += Subscription_DataChanged;

                subscription.AddItems(
                    items);

                Console.ReadLine();
            }

        }

        private static void Subscription_DataChanged(object sender, DataChangedEventArgs e)
        {
            foreach(DAReadResult r in e.Results)
            {
                Console.WriteLine($"{r.ItemID}\t{r.Value.Value}");
            }
        }
    }
}
