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

            //������
            ElementSetIterator it = seletion.ForwardIterator();

            //��������ѡ�е�Ԫ��;
            for (int i = 0; i < n; i++)
            {
                //ÿ�ε�����һ��Ԫ��;
                it.MoveNext();
                Element element = it.Current as Element;

                ParameterSet parameters = element.Parameters;//��ȡ����Ԫ�ص�ȫ������;
                ParameterSet parameters2 = GetParamSet(element);
                


                foreach (Parameter param in parameters)
                {
                    if (param == null) continue; //�������û�У�������һ��  
                    parametergroup = GetParameterGroup(parametergroup, param);    
                }

                foreach (Parameter param in parameters2)
                {
                    if (param == null) continue; //�������û�У�������һ�� 
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
        /// ��ȡ�������������Լ�
        /// </summary>
        /// <param name="parametergroup"></param>
        /// <param name="param"></param>
        /// <returns>����һ���б�����������еĴ����</returns>
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
        /// �Թ������ͽ��з��࣬�ж��Ƿ�Ϊ�ض����ڽ�����
        /// </summary>
        /// <param name="elem"></param>
        /// <returns>���ظù�������������</returns>
        private ParameterSet GetParamSet(Element elem)
        {
            FamilyInstance instance = elem as FamilyInstance;
            //�ܵ�
            Pipe pipe = elem as Pipe;
            //���
            Duct duct = elem as Duct;
            //ǽ
            Wall wall = elem as Wall;
            //¥��
            Floor floor = elem as Floor;
            //����
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
        /// ��������������������
        /// </summary>
        /// <param name="parametergroup"></param>
        private ParamTree MakeParamTree(Element elem, List<string> parametergroup )
        {
            //��һ����ڵ�����ֹ�����Id
            ParamTree Tree = new ParamTree("1",elem.UniqueId.ToString());
            //�ڶ���
            foreach (string str in parametergroup)
            {
                Tree.AddChildNodes(new ParamTree("2",str));
            }
            //�õ�ÿ�����Ե��������Ȼ����ҵ��ڶ���Ľڵ�
            ParameterSet parameters = elem.Parameters;//��ȡ����Ԫ�ص�ȫ������;

            //��ȡ��������:
            ParameterSet parameters2 = GetParamSet(elem);
            

            InsertThrParam(elem, Tree, parameters);
            InsertThrParam(elem, Tree, parameters2);

            return Tree;
        }


        /// <summary>
        /// ���������ĵ�����������
        /// </summary>
        /// <param name="parameterset"></param>
        private void InsertThrParam(Element elem, ParamTree Tree, ParameterSet parameter)
        {
            foreach (Parameter param in parameter)
            {
                if (param == null) continue; //�������û�У�������һ��
                //��λ�ڶ���
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
                //ѭ����ӵ���������
                tmptree.AddChildNodes(new ParamTree("3", param.Definition.Name, sb));
            }
        }
        #endregion



        /// <summary>
        /// ����JSON����
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
    /// ����һ������
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



        #region ����JSON�ļ�
        /// <summary>
        /// ����JSON�ļ� 
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
