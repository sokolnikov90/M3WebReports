using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Web;
using System.IO;

namespace M3IPClient
{
    public class SignIn
    {
        public struct ViewAllowed
        {
            public string name;
            public List<ViewItem> viewItems;
        }

        public struct ViewItem
        {
            public string name;
        }

        public struct Info : IDataInfo
        {
            public int isError { get; set; }
            public int userId;
            public int roleId;
            public List<ViewAllowed> viewsAllowed;
        }

        public Info info = new Info();

        public void ParseMessage(XmlNode messageNode)
        {
            info.isError = Convert.ToInt32(messageNode.SelectSingleNode("Request/Result").InnerText);

            if (info.isError == 0)
            {
                info.userId = Convert.ToInt32(messageNode.SelectSingleNode("Request/UserId").InnerText);
                info.roleId = Convert.ToInt32(messageNode.SelectSingleNode("Request/RoleId").InnerText);
                //info.role.name = messageNode.SelectSingleNode("Request/Role/Name").InnerText.Trim();
                //info.role.description = messageNode.SelectSingleNode("Request/Role/Description").InnerText.Trim();

                XmlNodeList viewNodeList = messageNode.SelectNodes("Request/Role/WebViewsAllowed/View");

                if (viewNodeList != null)
                {
                    info.viewsAllowed = new List<ViewAllowed>();

                    for (int i = 0; i < viewNodeList.Count; i++)
                    {
                        ViewAllowed viewAllowed = new ViewAllowed()
                        {
                            name = viewNodeList[i].SelectSingleNode("./@Name").InnerText.Trim(),
                            viewItems = new List<ViewItem>()
                        };

                        XmlNodeList viewItemsNodeList = viewNodeList[i].SelectNodes("./ViewItem");

                        if (viewItemsNodeList != null)
                        {
                            for (int j = 0; j < viewItemsNodeList.Count; j++)
                            {
                                viewAllowed.viewItems.Add(new ViewItem()
                                {
                                    name = viewItemsNodeList[j].InnerText.Trim()
                                });
                            }
                        }

                        info.viewsAllowed.Add(viewAllowed);
                    }
                }
            }
        }
    }
}