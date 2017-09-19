using System;
using System.Drawing;
using System.Runtime.Serialization;
using System.Collections.Generic;

namespace LangWars
{
    [DataContract]
    class LangDataList
    {
        [DataMember]
        public LangData[] items{ get; set; }
    }

    [DataContract]
	class LangData
	{
        [DataMember]
        public string name{ get; set; }
        [DataMember]
        public int count{ get; set; }
	}
}

