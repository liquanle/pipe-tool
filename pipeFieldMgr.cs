using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

namespace DataClient
{
    public enum FieldType { ft_String, ft_Decimal, ft_Integer, ft_Date };

    public struct FieldStruct{
        public string src;
        public string name;
        public FieldType fieldType;
        public int length;
        public int precision;
    }

    public class PipeLineFieldsMgr
    {
        public FieldStruct[] m_lineFields = new FieldStruct[30];

        public PipeLineFieldsMgr()
        {
            m_lineFields[0].src = "点号";
            m_lineFields[0].name = "起点点号";
            m_lineFields[0].fieldType = FieldType.ft_String;
            m_lineFields[0].length = 20;

            m_lineFields[1].src = "连接点号";
            m_lineFields[1].name = "终点点号";
            m_lineFields[1].fieldType = FieldType.ft_String;
            m_lineFields[1].length = 20;

            m_lineFields[2].src = "起点埋深";
            m_lineFields[2].name = "起点埋深";
            m_lineFields[2].fieldType = FieldType.ft_Decimal;
            m_lineFields[2].precision = 3;

            m_lineFields[3].src = "终点埋深";
            m_lineFields[3].name = "终点埋深";
            m_lineFields[3].fieldType = FieldType.ft_Decimal;
            m_lineFields[3].length = 8;
            m_lineFields[3].precision = 3;

            m_lineFields[4].src = "材质";
            m_lineFields[4].name = "材质";
            m_lineFields[4].fieldType = FieldType.ft_String;
            m_lineFields[4].length = 20;

            m_lineFields[5].src = "埋设方式";
            m_lineFields[5].name = "埋设方式";
            m_lineFields[5].fieldType = FieldType.ft_String;
            m_lineFields[5].length = 20;

            m_lineFields[6].src = "管径或断面尺寸(mm)";
            m_lineFields[6].name = "管径";
            m_lineFields[6].fieldType = FieldType.ft_String;
            m_lineFields[6].length = 20;

            m_lineFields[7].src = "埋设日期";
            m_lineFields[7].name = "埋设日期";
            m_lineFields[7].fieldType = FieldType.ft_Date;

            m_lineFields[8].src = "权属单位";
            m_lineFields[8].name = "权属单位";
            m_lineFields[8].fieldType = FieldType.ft_String;
            m_lineFields[8].length = 20;

            m_lineFields[9].src = "线型";
            m_lineFields[9].name = "线型";
            m_lineFields[9].fieldType = FieldType.ft_Integer;

            m_lineFields[10].src = "电缆根数";
            m_lineFields[10].name = "根数";
            m_lineFields[10].fieldType = FieldType.ft_String;
            m_lineFields[10].length = 20;

            m_lineFields[11].src = "压力或电压(kV)";
            m_lineFields[11].name = "电压";
            m_lineFields[11].fieldType = FieldType.ft_String;
            m_lineFields[11].length = 20;

            m_lineFields[12].src = "压力或电压(kV)";
            m_lineFields[12].name = "压力";
            m_lineFields[12].fieldType = FieldType.ft_String;
            m_lineFields[12].length = 20;

            m_lineFields[13].src = "管孔数/未用孔数";
            m_lineFields[13].name = "孔数";
            m_lineFields[13].fieldType = FieldType.ft_String;
            m_lineFields[13].length = 20;

            m_lineFields[14].src = "管孔数/未用孔数";
            m_lineFields[14].name = "未用孔数";
            m_lineFields[14].fieldType = FieldType.ft_String;
            m_lineFields[14].length = 20;

            m_lineFields[15].src = "套管尺寸";
            m_lineFields[15].name = "套管尺寸";
            m_lineFields[15].fieldType = FieldType.ft_String;
            m_lineFields[15].length = 20;

            m_lineFields[16].src = "道路名称";
            m_lineFields[16].name = "道路名称";
            m_lineFields[16].fieldType = FieldType.ft_String;
            m_lineFields[16].length = 20;

            m_lineFields[17].src = "流向";
            m_lineFields[17].name = "流向";
            m_lineFields[17].fieldType = FieldType.ft_Integer;

            m_lineFields[18].src = "备注";
            m_lineFields[18].name = "备注";
            m_lineFields[18].fieldType = FieldType.ft_String;
            m_lineFields[18].length = 20;

            m_lineFields[19].src = "关联编号";
            m_lineFields[19].name = "关联编号";
            m_lineFields[19].fieldType = FieldType.ft_String;
            m_lineFields[19].length = 20;

            m_lineFields[20].src = "SUR_DATE";
            m_lineFields[20].name = "SUR_DATE";
            m_lineFields[20].fieldType = FieldType.ft_Date;

            m_lineFields[21].src = "DB_DATE";
            m_lineFields[21].name = "DB_DATE";
            m_lineFields[21].fieldType = FieldType.ft_Date;

            m_lineFields[22].src = "E_CODE";
            m_lineFields[22].name = "E_CODE";
            m_lineFields[22].fieldType = FieldType.ft_String;
            m_lineFields[22].length = 20;

            m_lineFields[23].src = "P_ID";
            m_lineFields[23].name = "P_ID";
            m_lineFields[23].fieldType = FieldType.ft_String;
            m_lineFields[23].length = 20;

            m_lineFields[24].src = "管线性质";
            m_lineFields[24].name = "管线性质";
            m_lineFields[24].fieldType = FieldType.ft_String;
            m_lineFields[24].length = 20;

            m_lineFields[25].src = "LLENGTH";
            m_lineFields[25].name = "LLENGTH";
            m_lineFields[25].fieldType = FieldType.ft_Decimal;

            m_lineFields[26].src = "LayerName";
            m_lineFields[26].name = "LayerName";
            m_lineFields[26].fieldType = FieldType.ft_String;
            m_lineFields[26].length = 20;

            m_lineFields[27].src = "管线高程";
            m_lineFields[27].name = "起点高程";
            m_lineFields[27].fieldType = FieldType.ft_Decimal;

            m_lineFields[28].src = "终点高程";
            m_lineFields[28].name = "终点高程";
            m_lineFields[28].fieldType = FieldType.ft_Decimal;

            m_lineFields[29].src = "项目名称";
            m_lineFields[29].name = "项目名称";
            m_lineFields[29].fieldType = FieldType.ft_String;
            m_lineFields[29].length = 20;
        }

        public int getIndex(string strName)
        {
            for (int i = 0; i < m_lineFields.Length; i++)
            {
                FieldStruct fs = m_lineFields[i];
                if (fs.src == strName)
                {
                    return i;
                }
            }

            return -1;
        }

        //只有线表有这个功能
        public int getDstIndex(string strDstName)
        {
            for (int i = 0; i < m_lineFields.Length; i++)
            {
                FieldStruct fs = m_lineFields[i];
                if (fs.name == strDstName)
                {
                    return i;
                }
            }

            return -1;
        }
    }

    public class PipePointFieldsMgr
    {
        public FieldStruct[] m_pointFields = new FieldStruct[22];

        public PipePointFieldsMgr()
        {
            m_pointFields[0].src = "点号";
            m_pointFields[0].name = "物探点号";
            m_pointFields[0].fieldType = FieldType.ft_String;
            m_pointFields[0].length = 20;

            m_pointFields[1].src = "图上点号";
            m_pointFields[1].name = "图上点号";
            m_pointFields[1].fieldType = FieldType.ft_String;
            m_pointFields[1].length = 20;

            m_pointFields[2].src = "S_CODE";
            m_pointFields[2].name = "S_CODE";
            m_pointFields[2].fieldType = FieldType.ft_String;
            m_pointFields[2].length = 20;

            m_pointFields[3].src = "X坐标";
            m_pointFields[3].name = "X";
            m_pointFields[3].fieldType = FieldType.ft_Decimal;
            m_pointFields[3].precision = 3;

            m_pointFields[4].src = "Y坐标";
            m_pointFields[4].name = "Y";
            m_pointFields[4].fieldType = FieldType.ft_Decimal;
            m_pointFields[4].precision = 3;

            m_pointFields[5].src = "地面高程";
            m_pointFields[5].name = "地面高程";
            m_pointFields[5].fieldType = FieldType.ft_Decimal;
            m_pointFields[5].precision = 3;

            m_pointFields[6].src = "特征点";
            m_pointFields[6].name = "特征";
            m_pointFields[6].fieldType = FieldType.ft_String;
            m_pointFields[6].length = 20;

            m_pointFields[7].src = "附属物名称";
            m_pointFields[7].name = "附属物";
            m_pointFields[7].fieldType = FieldType.ft_String;
            m_pointFields[7].length = 20;

            m_pointFields[8].src = "偏心点号";
            m_pointFields[8].name = "偏心点号";
            m_pointFields[8].fieldType = FieldType.ft_String;
            m_pointFields[8].length = 20;

            m_pointFields[9].src = "MAPNO_X";
            m_pointFields[9].name = "MAPNO_X";
            m_pointFields[9].fieldType = FieldType.ft_Decimal;
            m_pointFields[9].length = 8;
            m_pointFields[9].precision = 3;

            m_pointFields[10].src = "MAPNO_Y";
            m_pointFields[10].name = "MAPNO_Y";
            m_pointFields[10].fieldType = FieldType.ft_Decimal;
            m_pointFields[10].length = 8;
            m_pointFields[10].precision = 3;

            m_pointFields[11].src = "图例角度";
            m_pointFields[11].name = "图例角度";
            m_pointFields[11].fieldType = FieldType.ft_Decimal;
            m_pointFields[11].length = 8;
            m_pointFields[11].precision = 3;

            m_pointFields[12].src = "图幅号";
            m_pointFields[12].name = "图幅号";
            m_pointFields[12].fieldType = FieldType.ft_String;
            m_pointFields[12].length = 20;

            m_pointFields[13].src = "备注";
            m_pointFields[13].name = "备注";
            m_pointFields[13].fieldType = FieldType.ft_String;
            m_pointFields[13].length = 20;

            m_pointFields[14].src = "SUR_DATE";
            m_pointFields[14].name = "SUR_DATE";
            m_pointFields[14].fieldType = FieldType.ft_Date;

            m_pointFields[15].src = "DB_DATE";
            m_pointFields[15].name = "DB_DATE";
            m_pointFields[15].fieldType = FieldType.ft_Date;

            m_pointFields[16].src = "E_CODE";
            m_pointFields[16].name = "E_CODE";
            m_pointFields[16].fieldType = FieldType.ft_String;
            m_pointFields[16].length = 20;

            m_pointFields[17].src = "P_ID";
            m_pointFields[17].name = "P_ID";
            m_pointFields[17].fieldType = FieldType.ft_String;
            m_pointFields[17].length = 20;

            m_pointFields[18].src = "管线性质";
            m_pointFields[18].name = "管线性质";
            m_pointFields[18].fieldType = FieldType.ft_String;
            m_pointFields[18].length = 20;

            m_pointFields[19].src = "LayerName";
            m_pointFields[19].name = "LayerName";
            m_pointFields[19].fieldType = FieldType.ft_String;
            m_pointFields[19].length = 20;

            m_pointFields[20].src = "管线高程";
            m_pointFields[20].name = "管线高程";
            m_pointFields[20].fieldType = FieldType.ft_Decimal;
            m_pointFields[20].precision = 3;

            m_pointFields[21].src = "项目名称";
            m_pointFields[21].name = "项目名称";
            m_pointFields[21].fieldType = FieldType.ft_String;
            m_pointFields[21].length = 20;
        }

        public int getIndex(string strName)
        {
            for (int i = 0; i < m_pointFields.Length; i++)
            {
                FieldStruct fs = m_pointFields[i];
                if (fs.src == strName)
                {
                    return i;
                }
            }

            return -1;
        }
    }
}
