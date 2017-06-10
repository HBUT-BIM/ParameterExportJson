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
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;

namespace ParameterExport
{

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.ReadOnly)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public class Command : IExternalCommand
    {
        #region IExternalCommand Members
       
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message,
            ElementSet elements)
        {
            // set out default result to failure.
            Autodesk.Revit.UI.Result retRes = Autodesk.Revit.UI.Result.Succeeded;
            Autodesk.Revit.UI.UIApplication app = commandData.Application;

            // get the elements selected
            // The current selection can be retrieved from the active 
            // document via the selection object
            ElementSet seletion = new ElementSet();
            foreach (ElementId elementId in app.ActiveUIDocument.Selection.GetElementIds())
            {
               seletion.Insert(app.ActiveUIDocument.Document.GetElement(elementId));
            }

            List<string> parametergroup = new List<string>();
            List<ParamTree> All_Elem = new List<ParamTree>();
            int n = seletion.Size;

            //迭代器
            ElementSetIterator it = seletion.ForwardIterator();

            //遍历所有选中的元素;
            for (int i = 0; i < n; i++)
            {
                //每次迭代下一个元素;
                it.MoveNext();
                Element element = it.Current as Element;

                ParameterSet parameters = element.Parameters;//获取单个元素的全部属性;
                ParameterSet parameters2 = GetParamSet(element);
                


                foreach (Parameter param in parameters)
                {
                    if (param == null) continue; //如果参数没有，继续下一个  
                    parametergroup = GetParameterGroup(parametergroup, param);    
                }

                foreach (Parameter param in parameters2)
                {
                    if (param == null) continue; //如果参数没有，继续下一个 
                    parametergroup = GetParameterGroup(parametergroup, param);
                }
                ParamTree tree = MakeParamTree(element, parametergroup);
                All_Elem.Add(tree);
            }

           
            CreateJsonFile(CreateJson(All_Elem));
            retRes = Autodesk.Revit.UI.Result.Succeeded;
            return retRes;
        }


        /// <summary>
        /// 获取单个构件的属性集
        /// </summary>
        /// <param name="parametergroup"></param>
        /// <param name="param"></param>
        /// <returns>返回一个列表，里面包含所有的大分类</returns>
        private List<string> GetParameterGroup(List<string> parametergroup, Parameter param)
        {
            if (parametergroup.Contains(LabelUtils.GetLabelFor(param.Definition.ParameterGroup)))
            {
            }
            else
            {
                parametergroup.Add(LabelUtils.GetLabelFor(param.Definition.ParameterGroup));
            }
            return parametergroup;
        }

        /// <summary>
        /// 对构件类型进行分类，判断是否为特定的内建类型
        /// </summary>
        /// <param name="elem"></param>
        /// <returns>返回该构件的类型属性</returns>
        private ParameterSet GetParamSet(Element elem)
        {
            FamilyInstance instance = elem as FamilyInstance;
            //管道
            Pipe pipe = elem as Pipe;
            //风管
            Duct duct = elem as Duct;
            //墙
            Wall wall = elem as Wall;
            //楼板
            Floor floor = elem as Floor;
            //轴网
            Grid grid = elem as Grid;

            ParameterSet parameters = new ParameterSet();
            if (instance != null)
            {
                parameters = instance.Symbol.Parameters;
            }
            else if (pipe != null)
            {
                parameters = pipe.PipeType.Parameters;
            }
            else if (duct != null)
            {
                parameters = duct.DuctType.Parameters;
            }
            else if (wall != null)
            {
                parameters = wall.WallType.Parameters;
            }
            else if (floor != null)
            {
                parameters = floor.FloorType.Parameters;
            }
            else if (grid != null)
            {
                parameters = grid.Parameters;
            }

            return parameters;
        }



        /// <summary>
        /// 制作单个构件的属性树
        /// </summary>
        /// <param name="parametergroup"></param>
        private ParamTree MakeParamTree(Element elem, List<string> parametergroup )
        {
            //第一层根节点的名字构件的Id
            ParamTree Tree = new ParamTree("1",elem.UniqueId.ToString());
            //第二层
            foreach (string str in parametergroup)
            {
                Tree.AddChildNodes(new ParamTree("2",str));
            }
            //得到每个属性的属性类别，然后查找到第二层的节点
            ParameterSet parameters = elem.Parameters;//获取单个元素的全部属性;

            //获取类型属性:
            ParameterSet parameters2 = GetParamSet(elem);
            

            InsertThrParam(elem, Tree, parameters);
            InsertThrParam(elem, Tree, parameters2);

            return Tree;
        }


        /// <summary>
        /// 向属性树的第三层插放数据
        /// </summary>
        /// <param name="parameterset"></param>
        private void InsertThrParam(Element elem, ParamTree Tree, ParameterSet parameter)
        {
            foreach (Parameter param in parameter)
            {
                if (param == null) continue; //如果参数没有，继续下一个
                //定位第二层
                ParamTree tmptree = Tree.GetChildNodesByName(LabelUtils.GetLabelFor(param.Definition.ParameterGroup));

                string sb = "";

                switch (param.StorageType)
                {
                    case Autodesk.Revit.DB.StorageType.Double:
                        sb += param.AsValueString();
                        break;

                    case Autodesk.Revit.DB.StorageType.Integer:
                        sb += param.AsInteger().ToString();
                        break;

                    case Autodesk.Revit.DB.StorageType.ElementId:
                        sb += (elem != null ? param.AsValueString() : "Not set");
                        break;

                    case Autodesk.Revit.DB.StorageType.String:
                        sb += param.AsString();
                        break;

                    case Autodesk.Revit.DB.StorageType.None:
                        sb += param.AsValueString();
                        break;
                    default:
                        break;
                }
                //循环添加第三层数据
                tmptree.AddChildNodes(new ParamTree("3", param.Definition.Name, sb));
            }
        }
        #endregion



        /// <summary>
        /// 生成JSON对象
        /// </summary>
        /// <param name="lstResult"></param>
        /// <returns></returns>
        private string CreateJson(List<ParamTree> lstResult)
        {
            StringBuilder strBuild = new StringBuilder("[");

            int index = 0;

            foreach (ParamTree item in lstResult)
            {
                if (index == 0)
                {
                    strBuild.Append(this.CreateElement(item));
                    strBuild.Append("\r\n");
                }
                else
                {
                    strBuild.Append("," + this.CreateElement(item));
                    strBuild.AppendLine("\r\n");
                }
                index++;

            }
            strBuild.AppendLine("]");
            return strBuild.ToString();
        }


        /// <summary>
    /// 创建一个对象
    /// </summary>
    /// <param name="item"></param>
    /// <returns></returns>
    private string CreateElement(ParamTree item)
    {
        StringBuilder strBuild = new StringBuilder("");

        if (item.GetChildNodesCount() != 0)
        {
            strBuild.AppendLine("\n\t{");
            strBuild.AppendLine("\t\t\"Name\":" + "\"" + item.Name + "\",");
            //strBuild.AppendLine("\t\t\"Value\":" + "\"" + item.Value + "\",");
            strBuild.AppendLine("\t\t\"children\":[");
            bool isFirst = true;

            foreach (ParamTree element in item.GetChildNodes())
            {
                if (isFirst)
                {
                    strBuild.Append(this.CreateElement(element));
                    isFirst = false;
                }
                else
                {
                    strBuild.AppendLine("," + this.CreateElement(element));
                }
            }
            strBuild.AppendLine("]}");
        }
        else
        {
            strBuild.AppendLine("\n\t{");
            strBuild.AppendLine("\t\t\"Name\":" + "\"" + item.Name + "\",");
            strBuild.AppendLine("\t\t\"Value\":" + "\"" + item.Value + "\"");
            strBuild.AppendLine("\n\t}");
            return strBuild.ToString();
        }

        return strBuild.ToString();
    }



        #region 创建JSON文件
        /// <summary>
        /// 创建JSON文件 
        /// </summary>
        /// <param name="fileData"></param>
        private void CreateJsonFile(String fileData)
        {
            string filePath = "C:\\Users\\yuanyizhou\\Documents" + "\\ParamTree_data.json";

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            using (StreamWriter sw = new StreamWriter(new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.Write), System.Text.Encoding.UTF8))
            {
                sw.WriteLine(fileData);
            }
        }
        #endregion

    }
}
