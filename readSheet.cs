using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Collections;
using System.Drawing;
using System.Data;
using System.Windows.Forms;


namespace DataClient
{
    public class readSheet
    {
        public class FieldMapCls
        {
            public bool bFinish = false;
            public int nDstPointTableIdx = -1;     //定义当前字段对应的管点目标索引
            public int nDstLineTableIdx = -1;      //定义当前字段对应的管线目标索引
            public string strDstFieldName;  //目标字段名
            public string strSrcFieldName;  //源字段名
            public FieldType emFieldType;

            public FieldMapCls()
            {
                bFinish = false;
                nDstPointTableIdx = -1;
                nDstLineTableIdx = -1;
                strDstFieldName = "";
                strSrcFieldName = "";
            }
        }

        public readSheet()
        {

            fillPipeKindMap();
        }

        private void fillPipeKindMap(){
            m_mapPipeKind.Clear();
            m_mapPipeKind.Add("饮用水", "JS");
            m_mapPipeKind.Add("天然气", "TQ");
            m_mapPipeKind.Add("供电", "GD");
            m_mapPipeKind.Add("交通信号", "JD");
            m_mapPipeKind.Add("中国电信", "DX");
            m_mapPipeKind.Add("中国移动", "YX");
            m_mapPipeKind.Add("中国联通", "LX");
            m_mapPipeKind.Add("军用", "AX");
            m_mapPipeKind.Add("监控信号", "KX");
            m_mapPipeKind.Add("电视", "TV");
            m_mapPipeKind.Add("电力通信", "EX");
            m_mapPipeKind.Add("路灯", "LD");
            m_mapPipeKind.Add("排水", "PS");
            m_mapPipeKind.Add("雨水", "YS");
            m_mapPipeKind.Add("非饮用水", "FS");
            m_mapPipeKind.Add("中国铁通", "TX");
            m_mapPipeKind.Add("中国网通", "WX");
        }

        private MdbHelper m_mdbHelper = null;
        public MdbHelper MdbHelper
        {
            set
            {
                m_mdbHelper = value;
            }
            get
            {
                return m_mdbHelper;
            }
        }

        private string m_strProjectName = "";
        public string ProjectName
        {
            set
            {
                m_strProjectName = value;
            }
            get
            {
                return m_strProjectName;
            }
        }

        private string m_pointNoPrefix = "";
        private string m_sheetName;
        private string m_strPipeKind = "";
        private string m_strOwnerDep = "";
        private ArrayList m_alFields;
        private System.Windows.Forms.ListBox m_listBoxCtrl;
        public string PointNoPrefix
        {
            set
            {
                m_pointNoPrefix = value;
            }
            get
            {
                return m_pointNoPrefix;
            }
        }

        public System.Windows.Forms.ListBox Output
        {
            set
            {
                m_listBoxCtrl = value;
            }
            get
            {
                return m_listBoxCtrl;
            }
        }

        public void OutPut(string strVal){
            m_listBoxCtrl.Items.Add(strVal);
            m_listBoxCtrl.SetSelected(m_listBoxCtrl.Items.Count - 1, true);
        }

        private Hashtable m_mapPipeKind = new Hashtable(); 

        private int m_nPointNoRowIdx = -1;      //点号对应的行号
        private int m_nPointNoColIdx = -1;      //点号对应的列号

        private int m_nDataRowIdx = -1;         //数据开始对应的行号
        private int m_nDataColIdx = -1;         //数据开始对应的行号

        public string m_const_pipeKind = "管线种类";
        public string m_const_ownerDep = "权属单位";

        private PipeLineFieldsMgr m_plFieldMgr = new PipeLineFieldsMgr();
        private PipePointFieldsMgr m_ptFieldMgr = new PipePointFieldsMgr();

        private System.Data.DataTable m_ptDataTable = null;     //点对应的datatable
        private System.Data.DataTable m_lnDataTable = null;     //线对应的datatable

        private FieldMapCls[] m_FieldInfos = null;  //放字段映射信息
        private bool FindFieldFinish()
        {
            for (int i = 0; i < m_FieldInfos.Length; i++)
            {
                FieldMapCls fmc = m_FieldInfos[i];
                if (fmc == null)
                {
                    return false;
                }

                if (!fmc.bFinish)
                {
                    return false;
                }
            }
            return true;
        }

        public void DealSheet(Worksheet worksheet)
        {
            readPipeKindAndDep(worksheet);
            readFieldInfos(worksheet);
            CreatePointDataTable();
            CreateLineDataTable();
            readDataToDataTable(worksheet);
        }

        private void CreatePointDataTable(){
            m_ptDataTable = new System.Data.DataTable(m_sheetName + "点");
            String msg = String.Format("正在创建表 [{0}]!", m_sheetName + "点");
            OutPut(msg);

            for (int i = 0; i < m_ptFieldMgr.m_pointFields.Length; i++)
            {
                FieldStruct fs = m_ptFieldMgr.m_pointFields[i];

                DataColumn dc = new DataColumn();
                dc.Caption = fs.name;
                dc.ColumnName = fs.name;
                switch (fs.fieldType){
                    case FieldType.ft_String: 
                        dc.DataType = System.Type.GetType("System.String");
                        dc.MaxLength = fs.length;
                        break;
                    case FieldType.ft_Integer: 
                        dc.DataType = System.Type.GetType("System.Int32");
                        break;
                    case FieldType.ft_Date:
                        dc.DataType = System.Type.GetType("System.DateTime");
                        break;
                    case FieldType.ft_Decimal:
                        dc.DataType = System.Type.GetType("System.Decimal");
                        break;
                    default:
                        dc.DataType = System.Type.GetType("System.String");
                        break;
                }
                m_ptDataTable.Columns.Add(dc);
            }
        }

        private void CreateLineDataTable(){
            m_lnDataTable = new System.Data.DataTable(m_sheetName + "线");
            String msg = String.Format("正在创建表[{0}]!", m_sheetName + "线");
            OutPut(msg);

            for (int i = 0; i < m_plFieldMgr.m_lineFields.Length; i++)
            {
                FieldStruct fs = m_plFieldMgr.m_lineFields[i];

                DataColumn dc = new DataColumn();
                dc.Caption = fs.name;
                dc.ColumnName = fs.name;
                switch (fs.fieldType)
                {
                    case FieldType.ft_String:
                        dc.DataType = System.Type.GetType("System.String");
                        dc.MaxLength = fs.length;
                        break;
                    case FieldType.ft_Integer:
                        dc.DataType = System.Type.GetType("System.Int32");
                        break;

                    case FieldType.ft_Date:
                        dc.DataType = System.Type.GetType("System.DateTime");
                        break;
                    case FieldType.ft_Decimal:
                        dc.DataType = System.Type.GetType("System.Decimal");
                        break;
                    default:
                        dc.DataType = System.Type.GetType("System.String");
                        break;
                }
                m_lnDataTable.Columns.Add(dc);
            }
        }

        public string SheetName
        {
            set
            {
                m_sheetName = value;
            }
            get
            {
                return m_sheetName;
            }
        }

        public string PipeKind{
            set{
                m_strPipeKind = value;
            }
            get{
                return m_strPipeKind;
            }
        }

        public string OwnerDep
        {
            set
            {
                m_strOwnerDep = value;
            }
            get
            {
                return m_strOwnerDep;
            }
        }

        public ArrayList Fields
        {
            set
            {
                m_alFields = value;
            }
            get
            {
                return m_alFields;
            }
        }

        public System.Data.DataTable PipePointTable
        {
            get
            {
                return m_ptDataTable;
            }
        }

        public System.Data.DataTable PipeLineTable
        {
            get
            {
                return m_lnDataTable;
            }
        }

        //先计算字段数量
        private int CalcateFieldCount(Worksheet worksheet)
        {
            int nRowCount = worksheet.UsedRange.Rows.Count;
            int nColCount = worksheet.UsedRange.Columns.Count;

            int nFieldCount = 0;
            bool bStartCount = false;

            for (int rowIdx = 1; rowIdx <= nRowCount; rowIdx++)
            {
                for (int colIdx = 1; colIdx <= nColCount; colIdx++)
                {
                    Range rgCur = ((Range)worksheet.Cells[rowIdx, colIdx]);

                    //假如不再有值了跳出
                    if (rgCur.Value == null)
                    {
                        if (bStartCount)
                        {
                            nFieldCount++;
                        }
                    }
                    else
                    {
                        string strVal = rgCur.Value as string;
                        strVal.Trim();
                        strVal = strVal.Replace(" ", string.Empty);
                        strVal = strVal.Replace(":", string.Empty);
                        strVal = strVal.Replace("：", string.Empty);

                        if (strVal == "点号")
                        {
                            bStartCount = true;
                        }

                        if (bStartCount)
                        {
                            nFieldCount++;
                        }
                        
                        if (strVal == "备注")
                        {
                            return nFieldCount;
                        }
                    }
                }
            }

            return -1;
        }

        public void readPipeKindAndDep(Worksheet worksheet){
            worksheet.Activate();
            m_sheetName = worksheet.Name;
            bool bFindPipeKind = false;
            bool bFindOwnerDep = false;

            int nRowCount = worksheet.UsedRange.Rows.Count;
            int nColCount = worksheet.UsedRange.Columns.Count;
            for (int rowIdx = 1; rowIdx <= nRowCount; rowIdx++)
            {
                for (int colIdx = 1; colIdx <= nColCount; colIdx++)
                {
                    if (m_strOwnerDep.Length * m_strPipeKind.Length > 0)
                    {
                        return;
                    }
                    Range rgCur = ((Range)worksheet.Cells[rowIdx, colIdx]);
                    rgCur.Select();

                    //假如不再有值了跳出
                    if (rgCur.Value == null)
                    {
                        Console.Write("rgn[" + rowIdx + "," + colIdx + "] = " + "空" + "\t");
                    }
                    else
                    {
                        string strVal = rgCur.Value as string;
                        strVal.Trim();
                        strVal = strVal.Replace(" ", string.Empty);
                        strVal = strVal.Replace(":", string.Empty);
                        strVal = strVal.Replace("：", string.Empty);
                        if (strVal.Contains(m_const_pipeKind))
                        {
                            bFindPipeKind = true;
                        }
                        else
                        {
                            if (bFindPipeKind)
                            {
                                m_strPipeKind = m_mapPipeKind[strVal].ToString();
                                Console.WriteLine(m_const_pipeKind + m_strPipeKind);
                                bFindPipeKind = false;
                                rgCur.Interior.Color = (ColorTranslator.ToOle(Color.Red));
                                continue;
                            }
                        }

                        if (strVal.Contains(m_const_ownerDep))
                        {
                            bFindOwnerDep = true;
                        }
                        else
                        {
                            if (bFindOwnerDep)
                            {
                                m_strOwnerDep = strVal;
                                Console.WriteLine(m_const_ownerDep + m_strOwnerDep);
                                bFindOwnerDep = false;
                                rgCur.Interior.Color = (ColorTranslator.ToOle(Color.Green));
                                continue;
                            }

                        }

                        rgCur.Interior.Color = (ColorTranslator.ToOle(Color.Orange));
                    }
                }
                Console.WriteLine();
            }
        }

        public void readFieldInfos(Worksheet worksheet)
        {
            m_sheetName = worksheet.Name;
            String msg = String.Format("开始读取表 [{0}] 字段信息!", m_sheetName);
            OutPut(msg);
           
            int nRowCount = worksheet.UsedRange.Rows.Count;
            int nColCount = worksheet.UsedRange.Columns.Count;
            bool bFindFirstField = false;
            int nRealFieldCount = CalcateFieldCount(worksheet);
            m_FieldInfos = new FieldMapCls[nRealFieldCount];  //初始化字段存储数组
            for (int rowIdx = 1; rowIdx <= nRowCount; rowIdx++)
            {
                for (int colIdx = 1; colIdx <= nColCount; colIdx++)
                {
                    Range rgCur = ((Range)worksheet.Cells[rowIdx, colIdx]);
                    rgCur.Select();

                    //假如不再有值了跳出
                    if (rgCur.Value == null)
                    {
                        Console.Write("rgn[" + rowIdx + "," + colIdx + "] = " + "空" + "\t");
                    }
                    else
                    {
                        string strVal = rgCur.Value as string;
                        strVal.Trim();
                        strVal = strVal.Replace(" ", string.Empty);
                        strVal = strVal.Replace(":", string.Empty);
                        strVal = strVal.Replace("：", string.Empty);
                        strVal = strVal.Replace("\n", string.Empty);

                        if (strVal == "点号")
                        {
                            m_nPointNoRowIdx = rowIdx;
                            m_nPointNoColIdx = colIdx;

                            m_nDataRowIdx = rowIdx + 3;
                            m_nDataColIdx = colIdx;

                            bFindFirstField = true;
                            m_FieldInfos[0] = new FieldMapCls();
                            m_FieldInfos[0].strSrcFieldName = strVal;

                            int nLindeDstIdx = m_plFieldMgr.getIndex(strVal);
                            int nPointDstIdx = m_ptFieldMgr.getIndex(strVal);

                            m_FieldInfos[0].nDstLineTableIdx = nLindeDstIdx;
                            m_FieldInfos[0].nDstPointTableIdx = nPointDstIdx;
                            m_FieldInfos[0].bFinish = true;
                            rgCur.Interior.Color = (ColorTranslator.ToOle(Color.Purple));
                        }
                        else
                        {
                            //当未找到第一个字段时
                            if (!bFindFirstField)
                            {
                                continue;
                            }
                            //当前目前索引已经有字段了continue
                            if (m_FieldInfos[colIdx - 1] != null)
                            {
                                continue;
                            }

                            Range rgCurNext = (Range)worksheet.Cells[rowIdx, colIdx + 1];
                            if (rgCurNext.Value != null || strVal == "备注" || strVal == "管线高程")
                            {
                                m_FieldInfos[colIdx - 1] = new FieldMapCls();
                                m_FieldInfos[colIdx - 1].strSrcFieldName = strVal;

                                int nLindeDstIdx = m_plFieldMgr.getIndex(strVal);
                                int nPointDstIdx = m_ptFieldMgr.getIndex(strVal);

                                m_FieldInfos[colIdx - 1].nDstLineTableIdx = nLindeDstIdx;
                                m_FieldInfos[colIdx - 1].nDstPointTableIdx = nPointDstIdx;
                                m_FieldInfos[colIdx - 1].bFinish = true;

                                rgCur.Interior.Color = (ColorTranslator.ToOle(Color.Purple));
                                //假如已完成所有字段查找，退出
                                if (FindFieldFinish())
                                {
                                    return;
                                }
                            }
                        }
                    }
                }
                Console.WriteLine();
            }
        }

        public void readDataToDataTable(Worksheet worksheet)
        {
            m_sheetName = worksheet.Name;
            String msgRead = String.Format("开始读取[{0}]表数据!", m_sheetName);
            OutPut(msgRead);

            int nRowCount = worksheet.UsedRange.Rows.Count;
            int nColCount = worksheet.UsedRange.Columns.Count;
            try{
                int nWriteCountPer100 = 0;
                 for (int rowIdx = m_nDataRowIdx; rowIdx <= nRowCount; rowIdx++){
                    System.Data.DataRow drPt = m_ptDataTable.NewRow();
                    System.Data.DataRow drLn = m_lnDataTable.NewRow();
                    for (int colIdx = m_nDataColIdx; colIdx <= m_FieldInfos.Length; colIdx++)
                    {
                        int nDstIdxOfPt = m_FieldInfos[colIdx - m_nDataColIdx].nDstPointTableIdx;
                        int nDstIdxOfLine = m_FieldInfos[colIdx - m_nDataColIdx].nDstLineTableIdx;

                        Range rgCur = ((Range)worksheet.Cells[rowIdx, colIdx]);
                        rgCur.Select();

                        //空值continue
                        if (rgCur.Value == null)
                        {
                            continue;
                        }
                        else
                        {
                            if (nDstIdxOfPt >= 0)
                            {
                                FieldStruct fsPt = m_ptFieldMgr.m_pointFields[nDstIdxOfPt];
                                drPt[nDstIdxOfPt] = rgCur.Value;
                                Console.WriteLine(rowIdx + " * " + colIdx);
                            }
                            if (nDstIdxOfLine >= 0)
                            {
                                //单独处理这个包含两个内容的字段
                                if (m_FieldInfos[colIdx - m_nDataColIdx].strSrcFieldName == "管孔数/未用孔数")
                                {
                                    int nHoleTotal = m_plFieldMgr.getDstIndex("孔数");
                                    int nHoleNoUsed = m_plFieldMgr.getDstIndex("未用孔数");
                                    
                                    if (rgCur.Value != null)
                                    {
                                        string strHole = rgCur.Value.ToString();
                                        string[] holeStrs = strHole.Split('/');
                                        if (holeStrs.Length == 2)
                                        {
                                            drLn[nHoleTotal] = holeStrs[0];
                                            drLn[nHoleNoUsed] = holeStrs[1];
                                        }
                                    }
                                }
                                else if (m_FieldInfos[colIdx - m_nDataColIdx].strSrcFieldName == "压力或电压(kV)")
                                {
                                    int nVotageIdx = m_plFieldMgr.getDstIndex("电压");
                                    int nPressIdx = m_plFieldMgr.getDstIndex("压力");

                                    if (rgCur.Value != null)
                                    {
                                        string strTemp = rgCur.Value.ToString();
                                        if (strTemp.ToUpper().Contains("KV"))
                                        {
                                            drLn[nVotageIdx] = strTemp;
                                            
                                        }
                                        else
                                        {
                                            drLn[nPressIdx] = strTemp;
                                        }
                                    }
                                
                                }else
                                {
                                    FieldStruct fsLine = m_plFieldMgr.m_lineFields[nDstIdxOfLine];
                                    drLn[nDstIdxOfLine] = rgCur.Value;
                                    Console.WriteLine(rowIdx + " * " + colIdx);
                                }
                            }
                        }
                    }
                    m_ptDataTable.Rows.Add(drPt);
                    m_lnDataTable.Rows.Add(drLn);


                    if (m_ptDataTable.Rows.Count - nWriteCountPer100 == 200)
                    {
                        String msgH = String.Format("已写入[{0}]表{1}条数据!", m_sheetName, m_ptDataTable.Rows.Count);
                        OutPut(msgH);
                        nWriteCountPer100 = m_ptDataTable.Rows.Count;
                    }
                    
                }
            }catch(Exception ex){
                MessageBox.Show(ex.Message);
            }
            String msgOver = String.Format("共写入[{0}]表{1}条数据!", m_sheetName, m_ptDataTable.Rows.Count);
            OutPut(msgOver);
        }

        public void ExecuteSql()
        {
            string strLnTableName = PipeLineTable.TableName;
            string strPtTableName = PipePointTable.TableName;

            m_mdbHelper.CreateTable(strPtTableName, PipePointTable);
            m_mdbHelper.DatatableToMdb(strPtTableName, PipePointTable);
            m_mdbHelper.CreateTable(strLnTableName, PipeLineTable);
            m_mdbHelper.DatatableToMdb(strLnTableName, PipeLineTable);

            string strPipeKindPt = "update " + strPtTableName + " set " + "管线性质" + " = \'" + PipeKind + "\'";
            string strPipeKindLn = "update " + strLnTableName + " set " + "管线性质" + " = \'" + PipeKind + "\'";
            string strOwnerDepLn = "update " + strLnTableName + " set " + "权属单位" + " = \'" + OwnerDep + "\'";

            m_mdbHelper.QueryReader(strPipeKindPt);
            m_mdbHelper.QueryReader(strPipeKindLn);
            m_mdbHelper.QueryReader(strOwnerDepLn);

            
            String sqlUpdatePipeGC  = String.Format("update {0},{1} set {0}.终点高程={1}.管线高程 where {0}.终点点号={1}.物探点号", strLnTableName, strPtTableName);
            String sqlUpdateStartMS = String.Format("update {0},{1} set {0}.起点埋深= Round({1}.地面高程 - {0}.起点高程, 3) where {0}.起点点号={1}.物探点号", strLnTableName, strPtTableName);
            String sqlUpdateEndMS = String.Format("update {0},{1}   set {0}.终点埋深= Round({1}.地面高程 - {0}.终点高程, 3) where {0}.终点点号={1}.物探点号", strLnTableName, strPtTableName);
            String sqlDelPipeGC = String.Format("alter table {0} drop COLUMN 管线高程", strPtTableName);
            String sqlAddPrifixPt = String.Format("UPDATE {0} SET 物探点号 = 物探点号 + '{1}'", strPtTableName, m_pointNoPrefix);
            String sqlAddPrifixLn = String.Format("UPDATE {0} SET 起点点号 = 起点点号 + '{1}', 终点点号 = 终点点号 + '{1}'", strLnTableName, m_pointNoPrefix);
            String sqlProNamePt = String.Format("update {0} set 项目名称= '{1}'", strPtTableName, ProjectName);
            String sqlProNameLn = String.Format("update {0} set 项目名称= '{1}'", strLnTableName, ProjectName);

            //更新线表的管线高程字段
            m_mdbHelper.QueryReader(sqlUpdatePipeGC);
            //更新线表的起点埋深字段
            m_mdbHelper.QueryReader(sqlUpdateStartMS);
            //更新线表的终点埋深字段
            m_mdbHelper.QueryReader(sqlUpdateEndMS);
            //删除点表的管线高程字段
            m_mdbHelper.QueryReader(sqlDelPipeGC);
            //点表的点号加随机前缀
            m_mdbHelper.QueryReader(sqlAddPrifixPt);
            //线表的点号加随机前缀
            m_mdbHelper.QueryReader(sqlAddPrifixLn);
            //更新点表的项目名称
            m_mdbHelper.QueryReader(sqlProNamePt);
            //更新线表的项目名称
            m_mdbHelper.QueryReader(sqlProNameLn);
        }
    }
}
