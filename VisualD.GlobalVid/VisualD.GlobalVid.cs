using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisualD.GlobalVid
{
    public class TGlobalVid : ICloneable
    {

        public TGlobalVid()
        {
        }

        public object Clone()
        {
            return this.MemberwiseClone();
        }
    }

    public class TGlobalAddOnOptions
    {
        public string AddonId;
        public string Opciones;
        public string SQLUsers;
    }
}
