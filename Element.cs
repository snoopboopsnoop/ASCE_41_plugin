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
    public class Element
    {
        [JsonProperty]
        private string name = "";
        [JsonProperty]
        private eMatType matType;
        [JsonProperty]
        private string propName = "";
        // frame | area
        [JsonProperty]
        private string type = "";
        [JsonProperty]
        private string pierName = "";
        [JsonProperty]
        private string spandrelName = "";
        [JsonProperty]
        private double[] m = { 1.0, 1.0 };
        [JsonProperty]
        private double J = 1.0;
        [JsonProperty]
        private string eControl = "Force Controlled";
        [JsonProperty]
        private double[] factorAdjust = { 1.0, 1.0 };
        [JsonProperty]
        private double knowledgeFactor = 1.00;

        [JsonProperty]
        private bool customK = false;

        public Element()
        {

        }

        public Element(string name, string propName, string type, eMatType matType, double kFactor) 
        {
            this.name = name;
            this.propName = propName;
            this.type = type;
            this.matType = matType;
            this.knowledgeFactor = kFactor;
        }

        public Element(string name, string propName, string type, string pierName, string spandrelName, eMatType matType, double kFactor) : this(name, propName, type, matType, kFactor)
        {
            this.pierName = pierName;
            this.spandrelName = spandrelName;
        }

        public string GetName ()
        {
            return this.name;
        }

        public string GetPropName () { return this.propName; }

        public eMatType GetMatType()
        {
            return this.matType;
        }

        public string GetType()
        {
            return this.type;
        }

        public string GetPierName()
        {
            return this.pierName;
        }

        public string GetSpandrelName()
        {
            return this.spandrelName;
        }

        public double[] GetM()
        {
            return this.m;
        }

        public double GetJ()
        {
            return this.J;
        }

        public string GetEControl()
        {
            return this.eControl;
        }

        public double[] GetFactorAdj()
        {
            return this.factorAdjust;
        }

        public double GetKFactor()
        {
            return this.knowledgeFactor;
        }

        public bool KEdited()
        {
            return this.customK;
        }

        public void SetMatType(eMatType matType)
        {
            this.matType = matType; 
        }

        public void SetPropName(string name)
        {
            this.propName = name;
        }

        public void SetM1(double m)
        {
            this.m[0] = m;
        }

        public void SetM2(double m)
        {
            this.m[1] = m;
        }

        public void SetJ(double j)
        {
            this.J = j;
        }

        public void SetControl(string control)
        {
            this.eControl = control;
        }

        public void SetFactor1(double factor)
        {
            this.factorAdjust[0] = factor;
        }

        public void SetFactor2(double factor)
        {
            this.factorAdjust[1] = factor;
        }

        public void SetKFactor(double knowledgeFactor)
        {
            this.knowledgeFactor = knowledgeFactor;
        }

        public void SetEdited(bool edited)
        {
            this.customK = edited;
        }
    }
}
