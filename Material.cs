using ETABSv1;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ASCE_41
{
    [Serializable]
    public class Material
    {
        [JsonProperty]
        string name = "";
        [JsonProperty]
        eMatType type = new eMatType();
        [JsonProperty]
        int color = 0;
        [JsonProperty]
        string notes = "";
        [JsonProperty]
        string GUID = "";
        [JsonProperty]
        double knowledgeFactor = 1.00;

        public Material(string name, eMatType type, int color, string notes, string GUID)
        {
            this.name = name;
            this.type = type;
            this.color = color;
            this.notes = notes;
            this.GUID = GUID;
        }

        public string GetName()
        {
            return this.name;
        }

        public eMatType GetMatType()
        {
            return this.type;
        }

        public double GetKFactor()
        {
            return this.knowledgeFactor;
        }


        public void SetKFactor(double kFactor)
        {
            this.knowledgeFactor = kFactor;
        }
    }
}
