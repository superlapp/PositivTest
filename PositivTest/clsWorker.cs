using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack;

namespace PositivTest
{
    class clsWorker
    {
        string page = @"http://www.moyareklama.by/%D0%93%D0%BE%D0%BC%D0%B5%D0%BB%D1%8C/%D0%BA%D0%B2%D0%B0%D1%80%D1%82%D0%B8%D1%80%D1%8B_%D0%BF%D1%80%D0%BE%D0%B4%D0%B0%D0%B6%D0%B0/%D0%B2%D1%81%D0%B5/8/1/";

        public List<string> Parse()
        {
            List<string> rls = new List<string>();
            //
            HtmlDocument doc = new HtmlWeb().Load(page);
            HtmlNodeCollection nc1 = doc.DocumentNode.SelectNodes("//div[@class='sa_header ']");
            //
            if (nc1 != null)
            {
                foreach (HtmlNode hn in nc1)
                {
                    HtmlNodeCollection ff = hn.ChildNodes;
                    string r = "";
                    foreach (HtmlNode h in ff)
                    {
                        if (h.NodeType == HtmlNodeType.Element)
                        {
                            string j = h.InnerText.Remove(0, 1);
                            string k = j.Trim();
                            r = r + k + ";";
                        }
                    }
                    if (r.Length != 0)
                    {
                        rls.Add(r);
                    }
                }
            }
            //
            return rls;
        }
    }
}
