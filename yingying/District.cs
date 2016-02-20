using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace yingying
{
    class District
    {
        /// <summary>
        /// 区名
        /// </summary>
        public string Name;
        /// <summary>
        /// 加工制造业
        /// </summary>
        public Business JiaGong;
        /// <summary>
        /// 建筑施工业
        /// </summary>
        public Business JianZhu;
        /// <summary>
        /// 批发零售业
        /// </summary>
        public Business PiFa;
        /// <summary>
        /// 餐饮住宿业
        /// </summary>
        public Business CanYin;
        /// <summary>
        /// 居民服务业
        /// </summary>
        public Business FuWu;
        /// <summary>
        /// 其它
        /// </summary>
        public Business Other;
        public TuFaSum tuFaSum;
    }

    public struct TuFaSum
    {
        public int count;
        public int person;
    }
}
