using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace yingying
{
    public struct Business
    {
        public struct ZhuDong 
        {
            /// <summary>
            /// 检查单位数
            /// </summary>
            public int unitCount;
            /// <summary>
            /// 案件数
            /// </summary>
            public int caseFinishNum;
            /// <summary>
            /// 涉及人数
            /// </summary>
            public int personNum;
            /// <summary>
            /// 追发工资金额
            /// </summary>
            public double amount;
            public Reason reason;
        }
        /// <summary>
        /// 主动监察
        /// </summary>
        public ZhuDong zhudong;
        public struct TouSu
        {
            /// <summary>
            /// 立案案件数
            /// </summary>
            public int caseAllNum;
            /// <summary>
            /// 结案数量
            /// </summary>
            public int caseFinishNum;
            /// <summary>
            /// 涉及人数
            /// </summary>
            public int personNum;
            /// <summary>
            /// 追发工资金额
            /// </summary>
            public double amount;
            public Reason reason;
        }
        /// <summary>
        /// 投诉举报
        /// </summary>
        public TouSu tousu;
        public struct TuFa
        {
            /// <summary>
            /// 事件数
            /// </summary>
            public int eventNum;
            /// <summary>
            /// 其中30人以上
            /// </summary>
            public int bigEventNum;
            /// <summary>
            /// 结案数量
            /// </summary>
            public int caseFinishNum;
            /// <summary>
            /// 涉及人数
            /// </summary>
            public int personNum;
            /// <summary>
            /// 追发工资金额
            /// </summary>
            public double amount;
            public Reason reason;
        }
        /// <summary>
        /// 突发事件
        /// </summary>
        public TuFa tufa;
        public struct ChuLi
        {
            /// <summary>
            /// 责令改正
            /// </summary>
            public int correctNum;
            /// <summary>
            /// 做出行政处理
            /// </summary>
            public int dealNum;
            /// <summary>
            /// 做出行政处罚
            /// </summary>
            public int penalizeNum;
            /// <summary>
            /// 罚款金额
            /// </summary>
            public double penalizeAmount;
        }
        /// <summary>
        /// 案件处理情况
        /// </summary>
        public ChuLi chuli;
    }

    public struct BusinessSum
    {
        /// <summary>
        /// 行业名称
        /// </summary>
        public string Name;
        /// <summary>
        /// 结案数量
        /// </summary>
        public int caseFinishNum;
        /// <summary>
        /// 涉及人数
        /// </summary>
        public int personNum;
        /// <summary>
        /// 追发工资金额
        /// </summary>
        public double amount;
        public Reason reason;
    }

    public struct Reason
    {
        /// <summary>
        /// 三无工程
        /// </summary>
        public int SanWu;
        /// <summary>
        /// 拖欠工程款
        /// </summary>
        public int GongChengKuan;
        /// <summary>
        /// 结算纠纷
        /// </summary>
        public int JieSuan;
        /// <summary>
        /// 非法转包
        /// </summary>
        public int ZhuanBao;
        /// <summary>
        /// 使用零散工
        /// </summary>
        public int SanGong;
        /// <summary>
        /// 无故拖欠工资
        /// </summary>
        public int GongZi;
        /// <summary>
        /// 其他原因
        /// </summary>
        public int Other;

    }
}
