using System;
using System.Collections.Generic;
using System.Text;
using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using System.IO;

namespace ParameterExport
{
    class ParamTree
    {
        public List<ParamTree> lstChildren = new List<ParamTree>();
        public string Name { get; set; }
        public string Value { get; set; }
        public string levelKey { get; set; }

        /// <summary>
        /// 添加子节点
        /// </summary>
        /// <param name="child"></param>
        public void AddChildNodes(ParamTree child)
        {
            this.lstChildren.Add(child);
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="name"></param>
        /// <param name="levelKey"></param>
        public ParamTree(string levelKey, string name, string value = "")
       {
           this.Name = name;
           this.levelKey = levelKey;
           this.Value = value;
       }

        /// <summary>
        /// 获取字节点的个数
        /// </summary>
        /// <returns></returns>
        public int GetChildNodesCount()
        {
            return lstChildren.Count;
        }

        /// <summary>
        /// 根据节点名称获取节点
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public ParamTree GetChildNodesByName(string name)
        {
            ParamTree paramTree = null;
             List<ParamTree> listpt = this.GetChildNodes();
            foreach (ParamTree pt in listpt)
            {
                if (pt.Name == name)
                {
                    paramTree = pt;
                    break;
                }
                else
                {
                    paramTree = null;
                }
            }
            return paramTree;
        }

        /// <summary>
        /// 获取所有的子节点
        /// </summary>
        /// <returns></returns>
        public List<ParamTree> GetChildNodes()
        {
            return lstChildren;
        }
    }
}
